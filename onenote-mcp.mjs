#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import { DeviceCodeCredential } from '@azure/identity';
import fetch from 'node-fetch';
import { z } from 'zod';
import open from 'open';

// Load environment variables
dotenv.config();

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Create the MCP server
const server = new McpServer(
  { 
    name: "onenote",
    version: "1.0.0",
    description: "OneNote MCP Server" 
  },
  {
    capabilities: {
      tools: {
        listChanged: true
      }
    }
  }
);

// Try to read the stored access token
let accessToken = null;
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error('Error reading access token file:', error.message);
}

// Alternatively, check if token is in environment variables
if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

let graphClient = null;

// Client ID for Microsoft Graph API access
const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Explorer client ID
const scopes = [
  'Notes.Read.All',
  'Notes.ReadWrite.All',
  'Sites.Read.All',     // For SharePoint sites and shared notebooks
  'Group.Read.All',     // For Teams notebooks
  'User.Read'
];

// Function to ensure Graph client is created
async function ensureGraphClient() {
  if (!graphClient) {
    // Read token from file if it exists
    try {
      if (fs.existsSync(tokenFilePath)) {
        const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
        try {
          // Try to parse as JSON first (new format)
          const parsedToken = JSON.parse(tokenData);
          accessToken = parsedToken.token;
        } catch (parseError) {
          // Fall back to using the raw token (old format)
          accessToken = tokenData;
        }
      }
    } catch (error) {
      console.error("Error reading token file:", error);
    }

    if (!accessToken) {
      throw new Error("Access token not found. Please save access token first.");
    }

    // Create Microsoft Graph client
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
  }
  return graphClient;
}

// Create graph client with device code auth or access token
async function createGraphClient() {
  if (accessToken) {
    // Use access token if available
    graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          return accessToken;
        }
      }
    });
    return { type: 'token', client: graphClient };
  } else {
    // Use device code flow
    const credential = new DeviceCodeCredential({
      clientId: clientId,
      userPromptCallback: (info) => {
        // This will be shown to the user with the URL and code
        console.error('\n' + info.message);
      }
    });

    try {
      // Get an access token using device code flow
      const tokenResponse = await credential.getToken(scopes);
      
      // Save the token for future use
      accessToken = tokenResponse.token;
      fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));
      
      // Initialize Graph client with the token
      graphClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            return accessToken;
          }
        }
      });
      
      return { type: 'device_code', client: graphClient };
    } catch (error) {
      console.error('Authentication error:', error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
}

// ============================================================================
// AUTHENTICATION
// ============================================================================

server.tool(
  "authenticate",
  "Start the authentication flow with Microsoft Graph. Use this first before accessing OneNote content.",
  async () => {
    try {
      const result = await createGraphClient();
      if (result.type === 'device_code') {
        return {
          content: [
            {
              type: "text",
              text: "Authentication started. Please check the console for the URL and code."
            }
          ]
        };
      } else {
        return {
          content: [
            {
              type: "text",
              text: "Already authenticated with an access token."
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in authentication:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
);

// ============================================================================
// DISCOVERY - Find your groups and explore structure
// ============================================================================

server.tool(
  "listGroups",
  "List all Microsoft 365 Groups and Teams you belong to (school, work, collaborative spaces). Use this to find group IDs needed for accessing shared notebooks.",
  async () => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/memberOf/$/microsoft.graph.group").get();

      const groups = response.value.map(group => ({
        id: group.id,
        displayName: group.displayName,
        description: group.description,
        mail: group.mail
      }));

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(groups, null, 2)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing groups:", error);
      throw new Error(`Failed to list groups: ${error.message}`);
    }
  }
);

// ============================================================================
// UNIFIED BROWSING - Works across personal and group notebooks
// ============================================================================

server.tool(
  "listAllNotebooks",
  "List ALL notebooks (personal and shared/Teams). Optionally filter by groupId. Use this when you want to see available notebooks. Returns notebooks with their source (personal or group name).",
  {
    groupId: z.string().optional().describe("Optional: Group ID to show only notebooks from that specific group")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { groupId } = params;

      const results = {
        personal: [],
        groups: []
      };

      // If groupId specified, only get that group's notebooks
      if (groupId) {
        try {
          const response = await graphClient.api(`/groups/${groupId}/onenote/notebooks`).get();

          // Get group name
          const groupInfo = await graphClient.api(`/groups/${groupId}`).select('displayName').get();

          results.groups.push({
            groupId: groupId,
            groupName: groupInfo.displayName,
            notebooks: response.value.map(nb => ({
              id: nb.id,
              name: nb.displayName,
              createdDateTime: nb.createdDateTime
            }))
          });
        } catch (error) {
          console.error(`Error getting notebooks for group ${groupId}:`, error.message);
        }

        return {
          content: [{
            type: "text",
            text: JSON.stringify(results, null, 2)
          }]
        };
      }

      // Otherwise, get all notebooks (personal + all groups)

      // Get personal notebooks
      try {
        const personalNbs = await graphClient.api("/me/onenote/notebooks").get();
        results.personal = personalNbs.value.map(nb => ({
          id: nb.id,
          name: nb.displayName,
          createdDateTime: nb.createdDateTime
        }));
      } catch (error) {
        console.error("Error getting personal notebooks:", error.message);
      }

      // Get all group notebooks
      try {
        const groupsResponse = await graphClient.api("/me/memberOf/$/microsoft.graph.group").get();

        for (const group of groupsResponse.value) {
          try {
            const groupNbs = await graphClient.api(`/groups/${group.id}/onenote/notebooks`).get();

            if (groupNbs.value.length > 0) {
              results.groups.push({
                groupId: group.id,
                groupName: group.displayName,
                notebooks: groupNbs.value.map(nb => ({
                  id: nb.id,
                  name: nb.displayName,
                  createdDateTime: nb.createdDateTime
                }))
              });
            }
          } catch (error) {
            console.error(`Error getting notebooks for group ${group.id}:`, error.message);
          }
        }
      } catch (error) {
        console.error("Error getting groups:", error.message);
      }

      const totalPersonal = results.personal.length;
      const totalGroups = results.groups.reduce((sum, g) => sum + g.notebooks.length, 0);

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            summary: {
              totalPersonal: totalPersonal,
              totalGroups: totalGroups,
              totalNotebooks: totalPersonal + totalGroups
            },
            results: results
          }, null, 2)
        }]
      };

    } catch (error) {
      console.error("Error in listAllNotebooks:", error);
      throw new Error(`Failed to list notebooks: ${error.message}`);
    }
  }
);

server.tool(
  "listAllSections",
  "List sections from notebooks. Can show personal sections, all sections from a group, or sections from a specific notebook. Also shows section groups (organizational folders like '_Indholdsbibliotek').",
  {
    groupId: z.string().optional().describe("Optional: Group ID to filter by group"),
    notebookId: z.string().optional().describe("Optional: Notebook ID to show only sections from that notebook"),
    sectionGroupId: z.string().optional().describe("Optional: Section Group ID to show sections inside that folder")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { groupId, notebookId, sectionGroupId } = params;

      // Case 1: Section Group ID provided → show sections inside that folder
      if (sectionGroupId) {
        if (!groupId) {
          throw new Error("groupId is required when using sectionGroupId");
        }

        const sections = await graphClient
          .api(`/groups/${groupId}/onenote/sectionGroups/${sectionGroupId}/sections`)
          .get();

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              source: "sectionGroup",
              groupId: groupId,
              sectionGroupId: sectionGroupId,
              count: sections.value.length,
              sections: sections.value.map(s => ({
                id: s.id,
                name: s.displayName
              }))
            }, null, 2)
          }]
        };
      }

      // Case 2: Group + Notebook ID → show sections AND section groups from that notebook
      if (groupId && notebookId) {
        const [sections, sectionGroups] = await Promise.all([
          graphClient.api(`/groups/${groupId}/onenote/notebooks/${notebookId}/sections`).get(),
          graphClient.api(`/groups/${groupId}/onenote/notebooks/${notebookId}/sectionGroups`).get()
        ]);

        const items = [
          ...sectionGroups.value.map(sg => ({
            type: "sectionGroup",
            id: sg.id,
            name: sg.displayName,
            icon: "📁"
          })),
          ...sections.value.map(sec => ({
            type: "section",
            id: sec.id,
            name: sec.displayName,
            icon: "📄"
          }))
        ];

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              source: "group_notebook",
              groupId: groupId,
              notebookId: notebookId,
              count: items.length,
              items: items
            }, null, 2)
          }]
        };
      }

      // Case 3: Only Group ID → show all sections and section groups across all notebooks in that group
      if (groupId) {
        const allItems = [];
        const notebooks = await graphClient.api(`/groups/${groupId}/onenote/notebooks`).get();

        for (const notebook of notebooks.value) {
          try {
            const [sections, sectionGroups] = await Promise.all([
              graphClient.api(`/groups/${groupId}/onenote/notebooks/${notebook.id}/sections`).get(),
              graphClient.api(`/groups/${groupId}/onenote/notebooks/${notebook.id}/sectionGroups`).get()
            ]);

            sectionGroups.value.forEach(sg => {
              allItems.push({
                type: "sectionGroup",
                notebookId: notebook.id,
                notebookName: notebook.displayName,
                sectionGroupId: sg.id,
                sectionGroupName: sg.displayName,
                icon: "📁"
              });
            });

            sections.value.forEach(section => {
              allItems.push({
                type: "section",
                notebookId: notebook.id,
                notebookName: notebook.displayName,
                sectionId: section.id,
                sectionName: section.displayName,
                icon: "📄"
              });
            });
          } catch (error) {
            console.error(`Error getting sections for notebook ${notebook.id}:`, error.message);
          }
        }

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              source: "group",
              groupId: groupId,
              count: allItems.length,
              items: allItems
            }, null, 2)
          }]
        };
      }

      // Case 4: Notebook ID only (personal) → show sections
      if (notebookId) {
        const sections = await graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).get();

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              source: "personal_notebook",
              notebookId: notebookId,
              count: sections.value.length,
              sections: sections.value.map(s => ({
                id: s.id,
                name: s.displayName
              }))
            }, null, 2)
          }]
        };
      }

      // Case 5: No params → show all personal sections
      const response = await graphClient.api(`/me/onenote/sections`).get();

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            source: "personal",
            count: response.value.length,
            sections: response.value.map(s => ({
              id: s.id,
              name: s.displayName
            }))
          }, null, 2)
        }]
      };

    } catch (error) {
      console.error("Error in listAllSections:", error);
      throw new Error(`Failed to list sections: ${error.message}`);
    }
  }
);

server.tool(
  "listAllPages",
  "List pages from sections. Can show all personal pages, all pages from a group, or pages from a specific section.",
  {
    groupId: z.string().optional().describe("Optional: Group ID to filter by group"),
    sectionId: z.string().optional().describe("Optional: Section ID to show only pages from that section")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { groupId, sectionId } = params;

      // Case 1: Group + Section ID → pages from that group section
      if (groupId && sectionId) {
        const pages = await graphClient
          .api(`/groups/${groupId}/onenote/sections/${sectionId}/pages`)
          .get();

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              source: "group_section",
              groupId: groupId,
              sectionId: sectionId,
              count: pages.value.length,
              pages: pages.value.map(page => ({
                id: page.id,
                title: page.title,
                created: page.createdDateTime,
                modified: page.lastModifiedDateTime
              }))
            }, null, 2)
          }]
        };
      }

      // Case 2: Only Section ID (personal) → pages from that personal section
      if (sectionId) {
        const pages = await graphClient
          .api(`/me/onenote/sections/${sectionId}/pages`)
          .get();

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              source: "personal_section",
              sectionId: sectionId,
              count: pages.value.length,
              pages: pages.value.map(page => ({
                id: page.id,
                title: page.title,
                created: page.createdDateTime,
                modified: page.lastModifiedDateTime
              }))
            }, null, 2)
          }]
        };
      }

      // Case 3: Only Group ID → all pages from that group
      if (groupId) {
        const pages = await graphClient
          .api(`/groups/${groupId}/onenote/pages`)
          .get();

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              source: "group",
              groupId: groupId,
              count: pages.value.length,
              pages: pages.value.map(page => ({
                id: page.id,
                title: page.title,
                created: page.createdDateTime,
                modified: page.lastModifiedDateTime
              }))
            }, null, 2)
          }]
        };
      }

      // Case 4: No params → all personal pages
      const pages = await graphClient.api("/me/onenote/pages").get();

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            source: "personal",
            count: pages.value.length,
            pages: pages.value.map(page => ({
              id: page.id,
              title: page.title,
              created: page.createdDateTime,
              modified: page.lastModifiedDateTime
            }))
          }, null, 2)
        }]
      };

    } catch (error) {
      console.error("Error in listAllPages:", error);
      throw new Error(`Failed to list pages: ${error.message}`);
    }
  }
);

server.tool(
  "browseOneNote",
  "Navigate through OneNote content step-by-step (notebooks → sections/section groups → pages). Provide no parameters to see all notebooks (personal + groups), then drill down by adding groupId, notebookId, sectionGroupId, or sectionId. Perfect for exploring structure layer by layer.",
  {
    groupId: z.string().optional().describe("Optional: Group ID to browse that group's content"),
    notebookId: z.string().optional().describe("Optional: Notebook ID to see sections and section groups"),
    sectionGroupId: z.string().optional().describe("Optional: Section Group ID to see sections inside that folder"),
    sectionId: z.string().optional().describe("Optional: Section ID to see pages in that section")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { groupId, notebookId, sectionGroupId, sectionId } = params;

      // No params → show ALL notebooks (personal + all groups)
      if (!groupId && !notebookId && !sectionGroupId && !sectionId) {
        const results = {
          personal: [],
          groups: []
        };

        // Get personal notebooks
        try {
          const personalNbs = await graphClient.api("/me/onenote/notebooks").get();
          results.personal = personalNbs.value.map(nb => ({
            type: "notebook",
            source: "personal",
            id: nb.id,
            name: nb.displayName
          }));
        } catch (error) {
          console.error("Error getting personal notebooks:", error.message);
        }

        // Get all group notebooks
        try {
          const groupsResponse = await graphClient.api("/me/memberOf/$/microsoft.graph.group").get();

          for (const group of groupsResponse.value) {
            try {
              const groupNbs = await graphClient.api(`/groups/${group.id}/onenote/notebooks`).get();

              if (groupNbs.value.length > 0) {
                results.groups.push({
                  groupId: group.id,
                  groupName: group.displayName,
                  notebooks: groupNbs.value.map(nb => ({
                    type: "notebook",
                    id: nb.id,
                    name: nb.displayName
                  }))
                });
              }
            } catch (error) {
              console.error(`Error getting notebooks for group ${group.id}:`, error.message);
            }
          }
        } catch (error) {
          console.error("Error getting groups:", error.message);
        }

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              level: "all_notebooks",
              results: results
            }, null, 2)
          }]
        };
      }

      // groupId only → show notebooks in that group
      if (groupId && !notebookId && !sectionGroupId && !sectionId) {
        const notebooks = await graphClient.api(`/groups/${groupId}/onenote/notebooks`).get();
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              level: "notebooks",
              groupId: groupId,
              count: notebooks.value.length,
              items: notebooks.value.map(nb => ({
                type: "notebook",
                id: nb.id,
                name: nb.displayName
              }))
            }, null, 2)
          }]
        };
      }

      // groupId + notebookId → show sections AND section groups
      if ((groupId || notebookId) && notebookId && !sectionGroupId && !sectionId) {
        const endpoint = groupId
          ? `/groups/${groupId}/onenote/notebooks/${notebookId}`
          : `/me/onenote/notebooks/${notebookId}`;

        const [sections, sectionGroups] = await Promise.all([
          graphClient.api(`${endpoint}/sections`).get(),
          graphClient.api(`${endpoint}/sectionGroups`).get()
        ]);

        const items = [
          ...sectionGroups.value.map(sg => ({
            type: "sectionGroup",
            id: sg.id,
            name: sg.displayName,
            icon: "📁"
          })),
          ...sections.value.map(sec => ({
            type: "section",
            id: sec.id,
            name: sec.displayName,
            icon: "📄"
          }))
        ];

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              level: "sections_and_groups",
              groupId: groupId || "personal",
              notebookId: notebookId,
              count: items.length,
              items: items
            }, null, 2)
          }]
        };
      }

      // sectionGroupId → show sections inside section group
      if (sectionGroupId && !sectionId) {
        const endpoint = groupId
          ? `/groups/${groupId}/onenote/sectionGroups/${sectionGroupId}`
          : `/me/onenote/sectionGroups/${sectionGroupId}`;

        const sections = await graphClient.api(`${endpoint}/sections`).get();

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              level: "sections_in_group",
              groupId: groupId || "personal",
              sectionGroupId: sectionGroupId,
              count: sections.value.length,
              items: sections.value.map(sec => ({
                type: "section",
                id: sec.id,
                name: sec.displayName
              }))
            }, null, 2)
          }]
        };
      }

      // sectionId → show pages
      if (sectionId) {
        const endpoint = groupId
          ? `/groups/${groupId}/onenote/sections/${sectionId}`
          : `/me/onenote/sections/${sectionId}`;

        const pages = await graphClient.api(`${endpoint}/pages`).get();

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              level: "pages",
              groupId: groupId || "personal",
              sectionId: sectionId,
              count: pages.value.length,
              items: pages.value.map(page => ({
                type: "page",
                id: page.id,
                title: page.title,
                created: page.createdDateTime,
                modified: page.lastModifiedDateTime
              }))
            }, null, 2)
          }]
        };
      }

    } catch (error) {
      console.error("Error browsing OneNote:", error);
      throw new Error(`Failed to browse OneNote: ${error.message}`);
    }
  }
);

// ============================================================================
// CONTENT READING
// ============================================================================

server.tool(
  "readPage",
  "Read the full text content of any OneNote page - automatically detects if it's personal or from a shared/Teams notebook. Provide pageId (and optionally groupId for faster lookup).",
  {
    pageId: z.string().describe("The ID of the page to read"),
    groupId: z.string().optional().describe("Optional: Group ID if known (makes it faster)")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { pageId, groupId } = params;

      // If groupId provided, use it directly (faster)
      if (groupId) {
        const pageDetails = await graphClient
          .api(`/groups/${groupId}/onenote/pages/${pageId}`)
          .get();

        const contentResponse = await fetch(pageDetails.contentUrl, {
          headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const htmlContent = await contentResponse.text();
        const { JSDOM } = await import('jsdom');
        const dom = new JSDOM(htmlContent);
        const bodyText = dom.window.document.body.textContent || '';
        const cleanText = bodyText
          .split('\n')
          .map(line => line.trim())
          .filter(line => line.length > 0)
          .join('\n');

        return {
          content: [{
            type: "text",
            text: `[Group: ${groupId}]\nTitle: ${pageDetails.title}\nCreated: ${pageDetails.createdDateTime}\nLast Modified: ${pageDetails.lastModifiedDateTime}\n\n--- Content ---\n\n${cleanText}`
          }]
        };
      }

      // Try personal first
      try {
        const pageDetails = await graphClient
          .api(`/me/onenote/pages/${pageId}`)
          .get();

        const contentResponse = await fetch(pageDetails.contentUrl, {
          headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const htmlContent = await contentResponse.text();
        const { JSDOM } = await import('jsdom');
        const dom = new JSDOM(htmlContent);
        const bodyText = dom.window.document.body.textContent || '';
        const cleanText = bodyText
          .split('\n')
          .map(line => line.trim())
          .filter(line => line.length > 0)
          .join('\n');

        return {
          content: [{
            type: "text",
            text: `[Personal]\nTitle: ${pageDetails.title}\nCreated: ${pageDetails.createdDateTime}\nLast Modified: ${pageDetails.lastModifiedDateTime}\n\n--- Content ---\n\n${cleanText}`
          }]
        };
      } catch (personalError) {
        // Not personal, search in groups
        const groupsResponse = await graphClient
          .api("/me/memberOf/$/microsoft.graph.group")
          .get();

        for (const group of groupsResponse.value) {
          try {
            const pageDetails = await graphClient
              .api(`/groups/${group.id}/onenote/pages/${pageId}`)
              .get();

            const contentResponse = await fetch(pageDetails.contentUrl, {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            });

            const htmlContent = await contentResponse.text();
            const { JSDOM } = await import('jsdom');
            const dom = new JSDOM(htmlContent);
            const bodyText = dom.window.document.body.textContent || '';
            const cleanText = bodyText
              .split('\n')
              .map(line => line.trim())
              .filter(line => line.length > 0)
              .join('\n');

            return {
              content: [{
                type: "text",
                text: `[Group: ${group.displayName}]\nTitle: ${pageDetails.title}\nCreated: ${pageDetails.createdDateTime}\nLast Modified: ${pageDetails.lastModifiedDateTime}\n\n--- Content ---\n\n${cleanText}`
              }]
            };
          } catch (groupError) {
            continue;
          }
        }

        throw new Error("Page not found in personal notebooks or any group notebooks");
      }

    } catch (error) {
      console.error("Error in readPage:", error);
      throw new Error(`Failed to read page: ${error.message}`);
    }
  }
);

// ============================================================================
// SEARCH
// ============================================================================

server.tool(
  "searchNotebooks",
  "Search for notebooks by name across both personal and all shared/Teams notebooks. Returns all matching notebooks with their location (personal or group name). Case-insensitive search.",
  {
    searchTerm: z.string().describe("Name or part of notebook name to search for")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { searchTerm } = params;
      const searchLower = searchTerm.toLowerCase();

      const results = {
        personal: [],
        groups: []
      };

      // Search personal notebooks
      try {
        const personalNbs = await graphClient.api("/me/onenote/notebooks").get();
        results.personal = personalNbs.value
          .filter(nb => nb.displayName.toLowerCase().includes(searchLower))
          .map(nb => ({
            id: nb.id,
            name: nb.displayName,
            createdDateTime: nb.createdDateTime
          }));
      } catch (error) {
        console.error("Error searching personal notebooks:", error.message);
      }

      // Search all group notebooks
      try {
        const groupsResponse = await graphClient
          .api("/me/memberOf/$/microsoft.graph.group")
          .get();

        for (const group of groupsResponse.value) {
          try {
            const groupNbs = await graphClient
              .api(`/groups/${group.id}/onenote/notebooks`)
              .get();

            const matches = groupNbs.value
              .filter(nb => nb.displayName.toLowerCase().includes(searchLower))
              .map(nb => ({
                id: nb.id,
                name: nb.displayName,
                createdDateTime: nb.createdDateTime
              }));

            if (matches.length > 0) {
              results.groups.push({
                groupId: group.id,
                groupName: group.displayName,
                notebooks: matches
              });
            }
          } catch (error) {
            console.error(`Error searching group ${group.id}:`, error.message);
          }
        }
      } catch (error) {
        console.error("Error searching groups:", error.message);
      }

      const totalCount = results.personal.length +
        results.groups.reduce((sum, g) => sum + g.notebooks.length, 0);

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            searchTerm: searchTerm,
            totalMatches: totalCount,
            results: results
          }, null, 2)
        }]
      };

    } catch (error) {
      console.error("Error in searchNotebooks:", error);
      throw new Error(`Failed to search notebooks: ${error.message}`);
    }
  }
);

server.tool(
  "searchAllPages",
  "Search for pages by title across both personal and all shared/Teams notebooks. Returns all matching pages with their location (personal or group name). Case-insensitive search.",
  {
    searchTerm: z.string().describe("Text to search for in page titles")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { searchTerm } = params;
      const searchLower = searchTerm.toLowerCase();

      const results = {
        personal: [],
        groups: []
      };

      // Search personal pages
      try {
        const personalPages = await graphClient.api("/me/onenote/pages").get();
        results.personal = personalPages.value
          .filter(page => page.title && page.title.toLowerCase().includes(searchLower))
          .map(page => ({
            id: page.id,
            title: page.title,
            createdDateTime: page.createdDateTime,
            lastModifiedDateTime: page.lastModifiedDateTime
          }));
      } catch (error) {
        console.error("Error searching personal pages:", error.message);
      }

      // Search all group pages
      try {
        const groupsResponse = await graphClient
          .api("/me/memberOf/$/microsoft.graph.group")
          .get();

        for (const group of groupsResponse.value) {
          try {
            const groupPages = await graphClient
              .api(`/groups/${group.id}/onenote/pages`)
              .get();

            const matches = groupPages.value
              .filter(page => page.title && page.title.toLowerCase().includes(searchLower))
              .map(page => ({
                id: page.id,
                title: page.title,
                createdDateTime: page.createdDateTime,
                lastModifiedDateTime: page.lastModifiedDateTime
              }));

            if (matches.length > 0) {
              results.groups.push({
                groupId: group.id,
                groupName: group.displayName,
                pages: matches
              });
            }
          } catch (error) {
            console.error(`Error searching pages in group ${group.id}:`, error.message);
          }
        }
      } catch (error) {
        console.error("Error searching group pages:", error.message);
      }

      const totalCount = results.personal.length +
        results.groups.reduce((sum, g) => sum + g.pages.length, 0);

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            searchTerm: searchTerm,
            totalMatches: totalCount,
            results: results
          }, null, 2)
        }]
      };

    } catch (error) {
      console.error("Error in searchAllPages:", error);
      throw new Error(`Failed to search pages: ${error.message}`);
    }
  }
);

// ============================================================================
// ACTIONS
// ============================================================================

server.tool(
  "openPage",
  "Open any OneNote page in the OneNote app or web browser. Automatically detects if page is personal or from a group. Optionally specify 'web' to force opening in browser.",
  {
    pageId: z.string().describe("The ID of the page to open"),
    groupId: z.string().optional().describe("Optional: Group ID if known (makes it faster)"),
    preferWeb: z.boolean().optional().describe("Optional: Set to true to force opening in web browser instead of desktop app")
  },
  async (params) => {
    try {
      await ensureGraphClient();
      const { pageId, groupId, preferWeb } = params;

      let pageDetails;

      // If groupId provided, use it
      if (groupId) {
        pageDetails = await graphClient
          .api(`/groups/${groupId}/onenote/pages/${pageId}`)
          .get();
      } else {
        // Try personal first
        try {
          pageDetails = await graphClient
            .api(`/me/onenote/pages/${pageId}`)
            .get();
        } catch (personalError) {
          // Search in groups
          const groupsResponse = await graphClient
            .api("/me/memberOf/$/microsoft.graph.group")
            .get();

          for (const group of groupsResponse.value) {
            try {
              pageDetails = await graphClient
                .api(`/groups/${group.id}/onenote/pages/${pageId}`)
                .get();
              break;
            } catch (groupError) {
              continue;
            }
          }

          if (!pageDetails) {
            throw new Error("Page not found in personal or any group notebooks");
          }
        }
      }

      // Choose URL based on preference
      const urlToOpen = preferWeb
        ? pageDetails.links?.oneNoteWebUrl?.href
        : (pageDetails.links?.oneNoteClientUrl?.href || pageDetails.links?.oneNoteWebUrl?.href);

      if (!urlToOpen) {
        throw new Error("Could not find URL to open page");
      }

      await open(urlToOpen);

      return {
        content: [{
          type: "text",
          text: `Opened page "${pageDetails.title}" ${preferWeb ? 'in web browser' : 'in OneNote'}.\nURL: ${urlToOpen}`
        }]
      };

    } catch (error) {
      console.error("Error opening page:", error);
      throw new Error(`Failed to open page: ${error.message}`);
    }
  }
);

// Connect to stdio and start server
async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('Server started successfully.');
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

