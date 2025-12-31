import { exec, spawn } from 'child_process';
import path from 'path';
import { promises as fs } from 'fs';

// BLOCKED/HIDDEN COMMANDS - prevent accidental logout
const BLOCKED_COMMANDS = ['logout', 'connection remove'];
const HIDDEN_COMMANDS = ['logout', 'connection remove'];

export async function runCliCommand(command: string, connectionName?: string): Promise<string> {
    // Check for blocked commands
    const lowerCommand = command.toLowerCase();
    for (const blocked of BLOCKED_COMMANDS) {
        if (lowerCommand.includes(blocked)) {
            return `ERROR: '${blocked}' command is disabled to prevent accidental logout. Use the CLI directly if you really need this.`;
        }
    }

    // If connectionName specified, prepend connection switch
    let fullCommand = command;
    if (connectionName) {
        fullCommand = `m365 connection use --name "${connectionName}" && ${command}`;
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