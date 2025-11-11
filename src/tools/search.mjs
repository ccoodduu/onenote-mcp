import { z } from 'zod';
import { ensureGraphClient, graphClient } from '../utils/graph-client.mjs';

export function registerSearchTools(server) {
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
}
