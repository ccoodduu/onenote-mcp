import { z } from 'zod';
import { ensureGraphClient, graphClient } from '../utils/graph-client.mjs';

export function registerBrowsingTools(server) {
  // listAllNotebooks
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

        if (groupId) {
          try {
            const response = await graphClient.api(`/groups/${groupId}/onenote/notebooks`).get();
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

  // listAllSections
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

  // listAllPages
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

  // browseOneNote - fortsætter...
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

        if (!groupId && !notebookId && !sectionGroupId && !sectionId) {
          const results = {
            personal: [],
            groups: []
          };

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
}
