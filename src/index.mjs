#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';

import { registerAuthenticationTools } from './auth/authentication.mjs';
import { registerDiscoveryTools } from './tools/discovery.mjs';
import { registerBrowsingTools } from './tools/browsing.mjs';
import { registerContentTools } from './tools/content.mjs';
import { registerSearchTools } from './tools/search.mjs';
import { registerActionTools } from './tools/actions.mjs';

dotenv.config();

const server = new McpServer(
  {
    name: "onenote",
    version: "2.0.0",
    description: "OneNote MCP Server - Unified Edition"
  },
  {
    capabilities: {
      tools: {
        listChanged: true
      }
    }
  }
);

registerAuthenticationTools(server);
registerDiscoveryTools(server);
registerBrowsingTools(server);
registerContentTools(server);
registerSearchTools(server);
registerActionTools(server);

async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);

    console.error('OneNote MCP Server started successfully (v2.0.0)');
    console.error('Use the "authenticate" tool to start the authentication flow.');

    process.on('SIGINT', () => {
      process.exit(0);
    });
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

main();
