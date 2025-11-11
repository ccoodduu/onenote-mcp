import { z } from 'zod';
import open from 'open';
import { ensureGraphClient, graphClient } from '../utils/graph-client.mjs';

export function registerActionTools(server) {
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

        if (groupId) {
          pageDetails = await graphClient
            .api(`/groups/${groupId}/onenote/pages/${pageId}`)
            .get();
        } else {
          try {
            pageDetails = await graphClient
              .api(`/me/onenote/pages/${pageId}`)
              .get();
          } catch (personalError) {
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
}
