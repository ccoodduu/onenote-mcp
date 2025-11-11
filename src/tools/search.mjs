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
    "Search for pages by title across both personal and all shared/Teams notebooks. Returns all matching pages with their location (personal or group name). Case-insensitive search. OPTIMIZED: Uses parallel searching and limits results for fast performance.",
    {
      searchTerm: z.string().describe("Text to search for in page titles"),
      maxResults: z.number().optional().describe("Optional: Maximum number of results to return (default: 100). Use lower values for faster searches.")
    },
    async (params) => {
      try {
        await ensureGraphClient();
        const { searchTerm, maxResults = 100 } = params;
        const searchLower = searchTerm.toLowerCase();

        const results = {
          personal: [],
          groups: []
        };

        // Fetch more pages than needed to ensure we find enough matches
        const fetchLimit = Math.max(maxResults * 3, 300);

        // Search personal pages
        const personalPromise = graphClient
          .api("/me/onenote/pages")
          .top(fetchLimit)
          .select("id,title,createdDateTime,lastModifiedDateTime")
          .orderby("lastModifiedDateTime desc")
          .get()
          .then(personalPages => {
            results.personal = personalPages.value
              .filter(page => page.title && page.title.toLowerCase().includes(searchLower))
              .slice(0, maxResults)
              .map(page => ({
                id: page.id,
                title: page.title,
                createdDateTime: page.createdDateTime,
                lastModifiedDateTime: page.lastModifiedDateTime
              }));
          })
          .catch(error => {
            console.error("Error searching personal pages:", error.message);
          });

        // Search groups in parallel
        const groupsPromise = graphClient
          .api("/me/memberOf/$/microsoft.graph.group")
          .get()
          .then(async (groupsResponse) => {
            const groupSearchPromises = groupsResponse.value.map(async (group) => {
              try {
                const groupPages = await graphClient
                  .api(`/groups/${group.id}/onenote/pages`)
                  .top(fetchLimit)
                  .select("id,title,createdDateTime,lastModifiedDateTime")
                  .orderby("lastModifiedDateTime desc")
                  .get();

                const matches = groupPages.value
                  .filter(page => page.title && page.title.toLowerCase().includes(searchLower))
                  .slice(0, maxResults)
                  .map(page => ({
                    id: page.id,
                    title: page.title,
                    createdDateTime: page.createdDateTime,
                    lastModifiedDateTime: page.lastModifiedDateTime
                  }));

                if (matches.length > 0) {
                  return {
                    groupId: group.id,
                    groupName: group.displayName,
                    pages: matches
                  };
                }
                return null;
              } catch (error) {
                console.error(`Error searching pages in group ${group.id}:`, error.message);
                return null;
              }
            });

            const groupResults = await Promise.all(groupSearchPromises);
            results.groups = groupResults.filter(g => g !== null);
          })
          .catch(error => {
            console.error("Error searching groups:", error.message);
          });

        // Wait for both personal and group searches to complete
        await Promise.all([personalPromise, groupsPromise]);

        const totalCount = results.personal.length +
          results.groups.reduce((sum, g) => sum + g.pages.length, 0);

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              searchTerm: searchTerm,
              maxResults: maxResults,
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
