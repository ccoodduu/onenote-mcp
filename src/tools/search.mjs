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
    "searchPagesInNotebook",
    "Search for pages by title within a specific notebook. Much faster than searching everywhere. Use listAllNotebooks first to find the notebookId, then search within that notebook. Case-insensitive search.",
    {
      notebookId: z.string().describe("The ID of the notebook to search in (get from listAllNotebooks)"),
      searchTerm: z.string().describe("Text to search for in page titles"),
      groupId: z.string().optional().describe("Optional: Group ID if this is a group notebook (makes it faster)")
    },
    async (params) => {
      try {
        await ensureGraphClient();
        const { notebookId, searchTerm, groupId } = params;
        const searchLower = searchTerm.toLowerCase();

        console.error(`Searching in notebook ${notebookId}...`);

        // Determine API path based on whether it's a group or personal notebook
        let apiPath;
        let notebookInfo;

        if (groupId) {
          // Group notebook
          apiPath = `/groups/${groupId}/onenote`;
          const notebook = await graphClient.api(`${apiPath}/notebooks/${notebookId}`).get();
          notebookInfo = {
            id: notebook.id,
            name: notebook.displayName,
            location: "group",
            groupId: groupId
          };
        } else {
          // Try personal first
          try {
            apiPath = `/me/onenote`;
            const notebook = await graphClient.api(`${apiPath}/notebooks/${notebookId}`).get();
            notebookInfo = {
              id: notebook.id,
              name: notebook.displayName,
              location: "personal"
            };
          } catch (personalError) {
            // Try to find it in groups
            const groupsResponse = await graphClient
              .api("/me/memberOf/$/microsoft.graph.group")
              .get();

            for (const group of groupsResponse.value) {
              try {
                apiPath = `/groups/${group.id}/onenote`;
                const notebook = await graphClient.api(`${apiPath}/notebooks/${notebookId}`).get();
                notebookInfo = {
                  id: notebook.id,
                  name: notebook.displayName,
                  location: "group",
                  groupId: group.id,
                  groupName: group.displayName
                };
                break;
              } catch (groupError) {
                continue;
              }
            }

            if (!notebookInfo) {
              throw new Error("Notebook not found in personal or any group notebooks");
            }
          }
        }

        console.error(`Found notebook: ${notebookInfo.name}`);

        // Get all pages from this notebook
        const allPages = [];

        try {
          // Get section groups
          try {
            const sectionGroups = await graphClient.api(`${apiPath}/notebooks/${notebookId}/sectionGroups`).get();

            for (const sectionGroup of sectionGroups.value) {
              try {
                const sections = await graphClient.api(`${apiPath}/sectionGroups/${sectionGroup.id}/sections`).get();

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
            // No section groups - that's ok
          }

          // Get sections directly in notebook
          const sections = await graphClient.api(`${apiPath}/notebooks/${notebookId}/sections`).get();

          for (const section of sections.value) {
            try {
              const pages = await graphClient.api(`${apiPath}/sections/${section.id}/pages`).get();
              allPages.push(...pages.value);
            } catch (error) {
              console.error(`Error fetching pages from section ${section.displayName}:`, error.message);
            }
          }
        } catch (error) {
          console.error(`Error fetching sections from notebook:`, error.message);
        }

        console.error(`Fetched ${allPages.length} total pages`);

        // Filter by search term
        const matches = allPages
          .filter(page => page.title && page.title.toLowerCase().includes(searchLower))
          .map(page => ({
            id: page.id,
            title: page.title,
            createdDateTime: page.createdDateTime,
            lastModifiedDateTime: page.lastModifiedDateTime
          }));

        console.error(`Found ${matches.length} matches`);

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              notebook: notebookInfo,
              searchTerm: searchTerm,
              totalMatches: matches.length,
              matches: matches
            }, null, 2)
          }]
        };

      } catch (error) {
        console.error("Error in searchPagesInNotebook:", error);
        throw new Error(`Failed to search pages in notebook: ${error.message}`);
      }
    }
  );
}
