import { exec, spawn } from 'child_process';
import path from 'path';
import { promises as fs } from 'fs';
import os from 'os';

// BLOCKED/HIDDEN COMMANDS - prevent accidental logout
const BLOCKED_COMMANDS = ['logout', 'connection remove'];
const HIDDEN_COMMANDS = ['logout', 'connection remove'];

// Alias storage file
const ALIASES_FILE = path.join(os.homedir(), '.m365-connection-aliases.json');

// Alias type: { alias: string, connectionId: string, tenant: string, appId?: string }
interface ConnectionAlias {
    alias: string;
    connectionId: string;
    tenant: string;
    appId?: string;
}

async function loadAliases(): Promise<ConnectionAlias[]> {
    try {
        const content = await fs.readFile(ALIASES_FILE, 'utf-8');
        return JSON.parse(content);
    } catch {
        return [];
    }
}

async function saveAliases(aliases: ConnectionAlias[]): Promise<void> {
    await fs.writeFile(ALIASES_FILE, JSON.stringify(aliases, null, 2));
}

export async function setConnectionAlias(alias: string, connectionId: string, tenant: string, appId?: string): Promise<string> {
    const aliases = await loadAliases();
    const existing = aliases.findIndex(a => a.alias === alias);

    const newAlias: ConnectionAlias = { alias, connectionId, tenant, appId };

    if (existing >= 0) {
        aliases[existing] = newAlias;
    } else {
        aliases.push(newAlias);
    }

    await saveAliases(aliases);
    return `Alias '${alias}' set for connection '${connectionId}' (tenant: ${tenant}${appId ? `, appId: ${appId}` : ''})`;
}

export async function removeConnectionAlias(alias: string): Promise<string> {
    const aliases = await loadAliases();
    const filtered = aliases.filter(a => a.alias !== alias);

    if (filtered.length === aliases.length) {
        return `Alias '${alias}' not found`;
    }

    await saveAliases(filtered);
    return `Alias '${alias}' removed`;
}

export async function resolveConnectionName(nameOrAlias: string): Promise<string> {
    const aliases = await loadAliases();
    const found = aliases.find(a => a.alias === nameOrAlias);
    return found ? found.connectionId : nameOrAlias;
}

export async function getConnectionStatus(connectionId: string): Promise<any> {
    try {
        // Get status for specific connection WITHOUT switching
        const result = await runCliCommandRaw(`m365 status --connection "${connectionId}"`);
        return JSON.parse(result);
    } catch (error) {
        return { error: String(error), connectionId };
    }
}

export async function validateConnection(connectionNameOrAlias: string): Promise<string> {
    const aliases = await loadAliases();
    const alias = aliases.find(a => a.alias === connectionNameOrAlias);
    const connectionId = alias ? alias.connectionId : connectionNameOrAlias;

    try {
        const status = await getConnectionStatus(connectionId);

        if (status.error) {
            return JSON.stringify({
                valid: false,
                connectionId,
                alias: alias?.alias || null,
                error: status.error
            }, null, 2);
        }

        // Check if appId matches alias expectation
        const appIdMatch = !alias?.appId || alias.appId === status.appId;

        return JSON.stringify({
            valid: true,
            connectionId,
            alias: alias?.alias || null,
            connectedAs: status.connectedAs,
            appId: status.appId,
            appTenant: status.appTenant,
            expectedAppId: alias?.appId || null,
            appIdMatch
        }, null, 2);
    } catch (error) {
        return JSON.stringify({
            valid: false,
            connectionId,
            alias: alias?.alias || null,
            error: String(error)
        }, null, 2);
    }
}

export async function validateAllConnections(): Promise<string> {
    const [connectionsResult, aliases] = await Promise.all([
        runCliCommandRaw('m365 connection list'),
        loadAliases()
    ]);

    try {
        const connections = JSON.parse(connectionsResult);
        const results = [];

        for (const conn of connections) {
            const alias = aliases.find(a => a.connectionId === conn.name);

            try {
                const status = await getConnectionStatus(conn.name);

                if (status.error) {
                    results.push({
                        connectionId: conn.name,
                        alias: alias?.alias || null,
                        valid: false,
                        error: status.error
                    });
                } else {
                    const appIdMatch = !alias?.appId || alias.appId === status.appId;
                    results.push({
                        connectionId: conn.name,
                        alias: alias?.alias || null,
                        valid: true,
                        connectedAs: status.connectedAs,
                        appId: status.appId,
                        expectedAppId: alias?.appId || null,
                        appIdMatch
                    });
                }
            } catch (error) {
                results.push({
                    connectionId: conn.name,
                    alias: alias?.alias || null,
                    valid: false,
                    error: String(error)
                });
            }
        }

        return JSON.stringify(results, null, 2);
    } catch (error) {
        return `Failed to validate connections: ${error}`;
    }
}

export async function listConnectionsWithAliases(): Promise<string> {
    const [connectionsResult, aliases] = await Promise.all([
        runCliCommandRaw('m365 connection list'),
        loadAliases()
    ]);

    try {
        const connections = JSON.parse(connectionsResult);
        const enriched = connections.map((conn: any) => {
            const alias = aliases.find(a => a.connectionId === conn.name);
            // Remove 'active' field - doesn't make sense when connectionName is always required
            const { active, ...rest } = conn;
            return {
                ...rest,
                alias: alias?.alias || null,
                appId: alias?.appId || null
            };
        });
        return JSON.stringify(enriched, null, 2);
    } catch {
        return connectionsResult;
    }
}

// Login with device code flow - returns device code IMMEDIATELY for user visibility
export async function loginWithDeviceCode(alias: string, tenant: string, appId?: string): Promise<string> {
    const useAppId = appId || '31359c7f-bd7e-475c-86db-fdb8c937548e';
    let loginCmd = `m365 login --authType deviceCode --appId ${useAppId}`;
    if (tenant) {
        loginCmd += ` --tenant ${tenant}`;
    }

    return new Promise((resolve) => {
        const subprocess = spawn(loginCmd, {
            shell: true,
            stdio: ['ignore', 'pipe', 'pipe'],
            detached: true,
        });

        let resolved = false;

        const handleOutput = (data: Buffer) => {
            const chunk = data.toString();
            const codeMatch = chunk.match(/enter the code ([A-Z0-9]+) to authenticate/i);
            if (codeMatch && !resolved) {
                resolved = true;
                const deviceCode = codeMatch[1];
                // Detach subprocess and close streams so it doesn't block MCP response
                subprocess.stdout?.destroy();
                subprocess.stderr?.destroy();
                subprocess.unref();
                // Return IMMEDIATELY with super clear formatting
                resolve(`
████████████████████████████████████████████████████████████
██                                                        ██
██   DEVICE CODE: ${deviceCode.padEnd(12)}                       ██
██                                                        ██
██   Go to: https://microsoft.com/devicelogin            ██
██                                                        ██
████████████████████████████████████████████████████████████

Enter the code above, then come back here.
Alias "${alias}" will be created after successful login.
`);
            }
        };

        subprocess.stdout.on('data', handleOutput);
        subprocess.stderr.on('data', handleOutput);

        subprocess.on('close', async (code) => {
            if (resolved) return; // Already returned device code
            if (code === 0) {
                try {
                    const status = await runCliCommandRaw('m365 status');
                    const statusJson = JSON.parse(status);
                    await setConnectionAlias(alias, statusJson.connectionName, tenant, statusJson.appId);
                    resolve(`Login successful! Alias "${alias}" created for ${statusJson.connectedAs}`);
                } catch {
                    resolve('Login succeeded but failed to create alias');
                }
            } else {
                resolve('Login failed - no device code received');
            }
        });

        subprocess.on('error', (err) => {
            resolve(JSON.stringify({
                success: false,
                error: err.message
            }, null, 2));
        });
    });
}

// Raw command execution without alias resolution (for internal use)
async function runCliCommandRaw(command: string): Promise<string> {
    let fullCommand = command;
    if (!fullCommand.includes('--output')) {
        fullCommand += ' --output json';
    }

    return new Promise((resolve, reject) => {
        const subprocess = spawn(fullCommand, {
            shell: true,
            timeout: 120000,
        });

        let output = '';
        let error = '';

        subprocess.stdout.on('data', (data) => {
            output += data.toString();
        });

        subprocess.stderr.on('data', (data) => {
            error += data.toString();
        });

        subprocess.on('close', (code) => {
            if (code === 0) {
                resolve(output.trim());
            } else {
                reject(new Error(error.trim() || `Command failed with exit code ${code}`));
            }
        });

        subprocess.on('error', (err) => {
            reject(err);
        });
    });
}

export async function runCliCommand(command: string, connectionName?: string): Promise<string> {
    // Check for blocked commands
    const lowerCommand = command.toLowerCase();
    for (const blocked of BLOCKED_COMMANDS) {
        if (lowerCommand.includes(blocked)) {
            return `ERROR: '${blocked}' command is disabled to prevent accidental logout. Use the CLI directly if you really need this.`;
        }
    }

    // REQUIRE connectionName when multiple connections exist
    if (!connectionName) {
        try {
            const connectionsRaw = await runCliCommandRaw('m365 connection list');
            const connections = JSON.parse(connectionsRaw);
            if (connections.length > 1) {
                const aliases = await loadAliases();
                const available = connections.map((c: any) => {
                    const alias = aliases.find(a => a.connectionId === c.name);
                    return alias ? `"${alias.alias}"` : c.name;
                }).join(', ');
                return `ERROR: Multiple connections exist. You MUST specify connectionName. Available: ${available}`;
            }
        } catch {
            // If we can't check, continue - command will fail naturally if needed
        }
    }

    // If connectionName specified, resolve alias and add --connection flag (NO SWITCHING)
    let fullCommand = command;
    if (connectionName) {
        const aliases = await loadAliases();
        const alias = aliases.find(a => a.alias === connectionName);
        const resolvedConnection = alias ? alias.connectionId : connectionName;

        // Add --connection flag to target specific connection WITHOUT switching
        // This uses the connection directly without any "connection use" nonsense
        if (!fullCommand.includes('--connection')) {
            fullCommand = `${command} --connection "${resolvedConnection}"`;
        }
    }

    if (!fullCommand.includes('--output')) {
        const commandPart = fullCommand.split('--')[0].trim();
        fullCommand += commandPart.endsWith(' list') ? ' --output csv' : ' --output json';
    }

    return new Promise((resolve, reject) => {
        const subprocess = spawn(fullCommand, {
            shell: true,
            timeout: 120000,
        });

        let output = '';
        let error = '';

        subprocess.stdout.on('data', (data) => {
            output += data.toString();
        });

        subprocess.stderr.on('data', (data) => {
            error += data.toString();
        });

        subprocess.on('close', (code) => {
            if (code === 0) {
                resolve(output.trim());
            } else {
                reject(new Error(error.trim() || `Command failed with exit code ${code}`));
            }
        });

        subprocess.on('error', (err) => {
            if (err.message.includes('timeout')) {
                reject(new Error('Command timed out'));
            } else {
                reject(err);
            }
        });
    });
}

export async function getCommandDocs(commandName: string, docs: string): Promise<any> {
    try {
        const filePath = await checkGlobalPackage('@pnp/cli-microsoft365', `docs${path.sep}docs${path.sep}cmd${path.sep}${docs}`);
        if (!filePath) {
            throw new Error('@pnp/cli-microsoft365 npm package not found or command documentation file not found');
        }

        const fileExists = await CheckIfFileExists(filePath);
        if (!fileExists) {
            throw new Error(`Documentation file for command ${commandName} not found at ${filePath}`);
        }

        const fileContent = await fs.readFile(filePath, 'utf-8');
        return fileContent;
    } catch (error) {
        console.error('An error occurred:', error);
        return `Failed to retrieve documentation for command ${commandName}: ${error}`;
    }
}

export async function getAllCommands(): Promise<any[]> {
    let commands: any[] = [];
    try {
        const filePath = await checkGlobalPackage('@pnp/cli-microsoft365', 'allCommandsFull.json');
        if (!filePath)
            throw new Error('@pnp/cli-microsoft365 npm package not found or allCommandsFull.json file not found');

        const fileContent = await fs.readFile(filePath, 'utf-8');
        const cliCommands = JSON.parse(fileContent);
        commands = cliCommands
            .filter((command: any) => !HIDDEN_COMMANDS.some(hidden => command.name.toLowerCase().includes(hidden)))
            .map((command: any) => ({
                name: `m365 ${command.name}`,
                description: command.description,
                docs: command.help
            }));
    } catch (error) {
        console.error('An error occurred:', error);
        return [{
            error: `Failed to retrieve commands: ${error}`
        }];
    }
    return commands;
}

async function CheckIfFileExists(filePath: string): Promise<boolean> {
    try {
        await fs.access(filePath);
        return true;
    } catch {
        return false;
    }
}

async function checkGlobalPackage(packageName: string, filePath: string): Promise<string | null> {
    return new Promise((resolve) => {
        exec('npm list -g --depth=0', (error, stdout, stderr) => {
            if (error) {
                console.error('Error checking global packages:', error);
                resolve(null);
                return;
            }

            if (stdout.includes(packageName)) {
                exec('npm root -g', (err, npmRoot) => {
                    if (err) {
                        console.error('Error getting npm root:', err);
                        resolve(null);
                        return;
                    }

                    const fileFullPath = path.join(npmRoot.trim(), packageName, filePath);
                    resolve(fileFullPath);
                });
            } else {
                console.log(`Package ${packageName} not found in global packages`);
                resolve(null);
            }
        });
    });
}