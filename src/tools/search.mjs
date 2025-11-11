import { z } from 'zod';
import { ensureGraphClient, graphClient } from '../utils/graph-client.mjs';

// Helper function to get all pages by iterating through notebooks, section groups, and sections
async function getAllPagesFromSource(apiPath, sourceName) {
  const allPages = [];

  try {
    // Get all notebooks
    const notebooks = await graphClient.api(`${apiPath}/notebooks`).get();

    // For each notebook, get section groups and sections
    for (const notebook of notebooks.value) {
      try {
        // Get section groups in this notebook (like "_Indholdsbibliotek")
        try {
          const sectionGroups = await graphClient.api(`${apiPath}/notebooks/${notebook.id}/sectionGroups`).get();

          // For each section group, get sections
          for (const sectionGroup of sectionGroups.value) {
            try {
              const sections = await graphClient.api(`${apiPath}/sectionGroups/${sectionGroup.id}/sections`).get();

              // For each section in section group, get pages
              for (const section of sections.value) {
                try {
                  const pages = await graphClient.api(`${apiPath}/sections/${section.id}/pages`).get();
                  allPages.push(...pages.value);
                } catch (error) {
                  console.error(`Error fetching pages from section ${section.displayName}:`, error.message);
                }
              }
            } catch (error) {
              console.error(`Error fetching sections from section group ${sectionGroup.displayName}:`, error.message);
            }
          }
        } catch (error) {
          // No section groups or error - that's ok, continue to regular sections
        }

        // Get sections directly in this notebook (not in section groups)
        const sections = await graphClient.api(`${apiPath}/notebooks/${notebook.id}/sections`).get();

        // For each section, get pages
        for (const section of sections.value) {
          try {
            const pages = await graphClient.api(`${apiPath}/sections/${section.id}/pages`).get();
            allPages.push(...pages.value);
          } catch (error) {
            console.error(`Error fetching pages from section ${section.displayName}:`, error.message);
          }
        }
      } catch (error) {
        console.error(`Error fetching sections from notebook ${notebook.displayName}:`, error.message);
      }
    }
  } catch (error) {
    console.error(`Error fetching notebooks from ${sourceName}:`, error.message);
  }

  return allPages;
}

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
    "Search for pages by title across both personal and all shared/Teams notebooks. Returns all matching pages with their location (personal or group name). Case-insensitive search. Handles large notebooks by searching section-by-section.",
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

        // Search personal pages using section-by-section approach
        try {
          console.error('Searching personal pages...');
          const personalPages = await getAllPagesFromSource('/me/onenote', 'personal');
          console.error(`Fetched ${personalPages.length} personal pages total`);

          results.personal = personalPages
            .filter(page => page.title && page.title.toLowerCase().includes(searchLower))
            .map(page => ({
              id: page.id,
              title: page.title,
              createdDateTime: page.createdDateTime,
              lastModifiedDateTime: page.lastModifiedDateTime
            }));

          console.error(`Found ${results.personal.length} personal matches`);
        } catch (error) {
          console.error("Error searching personal pages:", error.message);
        }

        // Search all group pages IN PARALLEL using section-by-section approach
        try {
          const groupsResponse = await graphClient
            .api("/me/memberOf/$/microsoft.graph.group")
            .get();

          console.error(`Searching ${groupsResponse.value.length} groups...`);

          // Create promises for parallel execution
          const groupSearchPromises = groupsResponse.value.map(async (group) => {
            try {
              const groupPages = await getAllPagesFromSource(`/groups/${group.id}/onenote`, group.displayName);

              const matches = groupPages
                .filter(page => page.title && page.title.toLowerCase().includes(searchLower))
                .map(page => ({
                  id: page.id,
                  title: page.title,
                  createdDateTime: page.createdDateTime,
                  lastModifiedDateTime: page.lastModifiedDateTime
                }));

              if (matches.length > 0) {
                console.error(`Found ${matches.length} matches in group: ${group.displayName}`);
                return {
                  groupId: group.id,
                  groupName: group.displayName,
                  pages: matches
                };
              }
              return null;
            } catch (error) {
              console.error(`Error searching pages in group ${group.displayName}:`, error.message);
              return null;
            }
          });

          // Wait for all to complete
          const groupResults = await Promise.all(groupSearchPromises);
          results.groups = groupResults.filter(g => g !== null);
        } catch (error) {
          console.error("Error searching group pages:", error.message);
        }

        const totalCount = results.personal.length +
          results.groups.reduce((sum, g) => sum + g.pages.length, 0);

        console.error(`Total matches found: ${totalCount}`);

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
