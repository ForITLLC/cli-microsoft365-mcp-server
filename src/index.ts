#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import * as util from './util.js';


const server = new McpServer({
    name: "microsoft-365-mcp-server",
    version: "0.0.1",
});

server.registerTool(
    'm365_get_commands',
    {
        title: 'Retrieve CLI for Microsoft 365 commands',
        description: 'Gets all CLI for Microsoft 365 commands to be used by the Model Context Protocol to pick the right command for a given task',
        inputSchema: {}
    },
    async ({ }) => {
        const commands = await util.getAllCommands();
        return {
            content: [
                { type: 'text', text: "TIP: Before executing any of the command run the 'm365_get_command_docs' tool to retrieve more context about it" },
                { type: 'text', text: "TIP: avoid setting the '--output' option when running commands. The optimal output format is automatically selected in 'm365_run_command' tool based on the command type." },
                { type: 'text', text: JSON.stringify(commands, null, 2) }
            ]
        };
    }
);

server.registerTool(
    'm365_get_command_docs',
    {
        title: 'Retrieve CLI for Microsoft 365 command docs',
        description: 'Gets documentation for a specified CLI for Microsoft 365 command to be used by the Model Context Protocol to provide detailed information about the command along with examples, use cases, and option descriptions',
        inputSchema:
        {
            commandName: z.string().describe('command name which for which documentation is requested'),
            docs: z.string().describe('file path to command documentation')
        }
    },
    async ({ commandName, docs }) => ({
        content: [
            { type: 'text', text: "TIP: avoid setting the '--output' option when running commands. The optimal output format is automatically selected in 'm365_run_command' tool based on the command type." },
            { type: 'text', text: await util.getCommandDocs(commandName, docs) }
        ]
    })
);

server.registerTool(
    'm365_run_command',
    {
        title: 'Execute CLI for Microsoft 365 command',
        description: 'Runs a specified CLI for Microsoft 365 command to be used by the Model Context Protocol to execute the command and return the result and reason over the response. Can target specific tenant connection without manual switching.',
        inputSchema:
        {
            command: z.string().describe('command name which should be executed'),
            connectionName: z.string().optional().describe('Target a specific connection by name (from m365_list_connections) without switching. Omit to use active connection.')
        }
    },
    async ({ command, connectionName }) => ({
        content: [{ type: 'text', text: await util.runCliCommand(command, connectionName) }]
    })
);

server.registerTool(
    'm365_list_connections',
    {
        title: 'List CLI for Microsoft 365 connections',
        description: 'Lists all available tenant connections with their aliases. Use alias or connection name with m365_run_command connectionName parameter to target specific tenants.',
        inputSchema: {}
    },
    async ({}) => ({
        content: [{ type: 'text', text: await util.listConnectionsWithAliases() }]
    })
);

server.registerTool(
    'm365_set_connection_alias',
    {
        title: 'Set friendly alias for a connection',
        description: 'Creates a friendly name alias for a tenant connection. Use aliases instead of connection IDs for easier multi-tenant targeting.',
        inputSchema: {
            alias: z.string().describe('Friendly name for the connection (e.g., "ForIT", "ClientX")'),
            connectionId: z.string().describe('The connection ID/name from m365_list_connections'),
            tenant: z.string().describe('Tenant name or domain (e.g., "foritllc.onmicrosoft.com")'),
            appId: z.string().optional().describe('Optional: Azure AD app registration ID used for this connection')
        }
    },
    async ({ alias, connectionId, tenant, appId }) => ({
        content: [{ type: 'text', text: await util.setConnectionAlias(alias, connectionId, tenant, appId) }]
    })
);

server.registerTool(
    'm365_remove_connection_alias',
    {
        title: 'Remove a connection alias',
        description: 'Removes a friendly name alias for a tenant connection.',
        inputSchema: {
            alias: z.string().describe('The alias to remove')
        }
    },
    async ({ alias }) => ({
        content: [{ type: 'text', text: await util.removeConnectionAlias(alias) }]
    })
);

server.registerTool(
    'm365_validate_connection',
    {
        title: 'Validate a specific connection',
        description: 'Checks if a connection is valid and working. Verifies the appId matches the expected value if an alias with appId is configured.',
        inputSchema: {
            connectionNameOrAlias: z.string().describe('Connection ID or alias to validate')
        }
    },
    async ({ connectionNameOrAlias }) => ({
        content: [{ type: 'text', text: await util.validateConnection(connectionNameOrAlias) }]
    })
);

server.registerTool(
    'm365_validate_all_connections',
    {
        title: 'Validate all connections',
        description: 'Checks all connections for validity. Identifies broken connections and appId mismatches. Use this to find connections that need to be removed or re-authenticated.',
        inputSchema: {}
    },
    async ({}) => ({
        content: [{ type: 'text', text: await util.validateAllConnections() }]
    })
);

const transport = new StdioServerTransport();
await server.connect(transport);