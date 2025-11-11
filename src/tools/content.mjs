import { z } from 'zod';
import fetch from 'node-fetch';
import { ensureGraphClient, graphClient, accessToken } from '../utils/graph-client.mjs';

export function registerContentTools(server) {
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
}
