import { z } from 'zod';
import fetch from 'node-fetch';
import { ensureGraphClient, graphClient, accessToken } from '../utils/graph-client.mjs';

// Helper function to extract file attachments from HTML content
function extractAttachments(htmlContent) {
  const attachments = [];
  const objectMatches = htmlContent.match(/<object[^>]*>/gi) || [];

  objectMatches.forEach(tag => {
    const nameMatch = tag.match(/data-attachment="([^"]*)"/i);
    const typeMatch = tag.match(/type="([^"]*)"/i);
    const dataMatch = tag.match(/data="([^"]*)"/i);

    if (nameMatch || dataMatch) {
      attachments.push({
        name: nameMatch ? nameMatch[1] : 'Unknown file',
        type: typeMatch ? typeMatch[1] : 'unknown',
        url: dataMatch ? dataMatch[1] : null
      });
    }
  });

  return attachments;
}

// Helper function to extract images from HTML content
function extractImages(htmlContent) {
  const images = [];
  const imgMatches = htmlContent.match(/<img[^>]*>/gi) || [];

  imgMatches.forEach((tag, index) => {
    const srcMatch = tag.match(/src="([^"]*)"/i);
    const altMatch = tag.match(/alt="([^"]*)"/i);
    const widthMatch = tag.match(/width="([^"]*)"/i);
    const heightMatch = tag.match(/height="([^"]*)"/i);

    if (srcMatch) {
      images.push({
        index: index + 1,
        url: srcMatch[1],
        alt: altMatch ? altMatch[1] : 'Image',
        width: widthMatch ? widthMatch[1] : null,
        height: heightMatch ? heightMatch[1] : null
      });
    }
  });

  return images;
}

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

          // Extract attachments
          const attachments = extractAttachments(htmlContent);
          const attachmentText = attachments.length > 0
            ? `\n\n--- Attachments (${attachments.length}) ---\n\n` +
              attachments.map((att, idx) => `📎 ${idx + 1}. ${att.name}${att.type !== 'unknown' ? ` (${att.type})` : ''}`).join('\n') +
              '\n\nUse getPageAttachments tool to read specific files.'
            : '';

          // Extract images
          const images = extractImages(htmlContent);
          const imageText = images.length > 0
            ? `\n\n--- Images (${images.length}) ---\n\n` +
              images.map(img => {
                const sizeText = img.width && img.height ? ` (${img.width}x${img.height}px)` : '';
                return `🖼️ ${img.index}. ${img.alt}${sizeText}`;
              }).join('\n') +
              '\n\nUse getPageImages tool to view specific images.'
            : '';

          return {
            content: [{
              type: "text",
              text: `[Group: ${groupId}]\nTitle: ${pageDetails.title}\nCreated: ${pageDetails.createdDateTime}\nLast Modified: ${pageDetails.lastModifiedDateTime}\n\n--- Content ---\n\n${cleanText}${attachmentText}${imageText}`
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

          // Extract attachments
          const attachments = extractAttachments(htmlContent);
          const attachmentText = attachments.length > 0
            ? `\n\n--- Attachments (${attachments.length}) ---\n\n` +
              attachments.map((att, idx) => `📎 ${idx + 1}. ${att.name}${att.type !== 'unknown' ? ` (${att.type})` : ''}`).join('\n') +
              '\n\nUse getPageAttachments tool to read specific files.'
            : '';

          // Extract images
          const images = extractImages(htmlContent);
          const imageText = images.length > 0
            ? `\n\n--- Images (${images.length}) ---\n\n` +
              images.map(img => {
                const sizeText = img.width && img.height ? ` (${img.width}x${img.height}px)` : '';
                return `🖼️ ${img.index}. ${img.alt}${sizeText}`;
              }).join('\n') +
              '\n\nUse getPageImages tool to view specific images.'
            : '';

          return {
            content: [{
              type: "text",
              text: `[Personal]\nTitle: ${pageDetails.title}\nCreated: ${pageDetails.createdDateTime}\nLast Modified: ${pageDetails.lastModifiedDateTime}\n\n--- Content ---\n\n${cleanText}${attachmentText}${imageText}`
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

              // Extract attachments
              const attachments = extractAttachments(htmlContent);
              const attachmentText = attachments.length > 0
                ? `\n\n--- Attachments (${attachments.length}) ---\n\n` +
                  attachments.map(att => `📎 ${att.name}${att.type !== 'unknown' ? ` (${att.type})` : ''}`).join('\n')
                : '';

              // Extract images
              const images = extractImages(htmlContent);
              const imageText = images.length > 0
                ? `\n\n--- Images (${images.length}) ---\n\n` +
                  images.map(img => {
                    const sizeText = img.width && img.height ? ` (${img.width}x${img.height}px)` : '';
                    return `🖼️ ${img.index}. ${img.alt}${sizeText}`;
                  }).join('\n') +
                  '\n\nUse getPageImages tool to view specific images.'
                : '';

              return {
                content: [{
                  type: "text",
                  text: `[Group: ${group.displayName}]\nTitle: ${pageDetails.title}\nCreated: ${pageDetails.createdDateTime}\nLast Modified: ${pageDetails.lastModifiedDateTime}\n\n--- Content ---\n\n${cleanText}${attachmentText}${imageText}`
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

  server.tool(
    "getPageImages",
    "Fetch and view specific images from a OneNote page. Returns images as base64-encoded data that Claude can see. Use this after readPage shows you which images are available.",
    {
      pageId: z.string().describe("The ID of the page containing the images"),
      imageIndices: z.array(z.number()).describe("Array of image numbers to fetch (e.g., [1, 2] for first and second image)"),
      groupId: z.string().optional().describe("Optional: Group ID if known (makes it faster)")
    },
    async (params) => {
      try {
        await ensureGraphClient();
        const { pageId, imageIndices, groupId } = params;

        // Find the page and get its content
        let htmlContent;
        let pageTitle;

        if (groupId) {
          const pageDetails = await graphClient
            .api(`/groups/${groupId}/onenote/pages/${pageId}`)
            .get();

          const contentResponse = await fetch(pageDetails.contentUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          });
          htmlContent = await contentResponse.text();
          pageTitle = pageDetails.title;
        } else {
          // Try personal first, then groups
          try {
            const pageDetails = await graphClient
              .api(`/me/onenote/pages/${pageId}`)
              .get();

            const contentResponse = await fetch(pageDetails.contentUrl, {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            htmlContent = await contentResponse.text();
            pageTitle = pageDetails.title;
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
                htmlContent = await contentResponse.text();
                pageTitle = pageDetails.title;
                break;
              } catch (groupError) {
                continue;
              }
            }

            if (!htmlContent) {
              throw new Error("Page not found in personal notebooks or any group notebooks");
            }
          }
        }

        // Extract all images from the page
        const allImages = extractImages(htmlContent);

        if (allImages.length === 0) {
          return {
            content: [{
              type: "text",
              text: "No images found on this page."
            }]
          };
        }

        // Fetch the requested images
        const content = [{
          type: "text",
          text: `Images from "${pageTitle}":\n`
        }];

        for (const index of imageIndices) {
          if (index < 1 || index > allImages.length) {
            content.push({
              type: "text",
              text: `\n⚠️ Image ${index} not found (page has ${allImages.length} images)`
            });
            continue;
          }

          const image = allImages[index - 1];

          try {
            // Fetch the image with authentication
            const imageResponse = await fetch(image.url, {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            });

            if (!imageResponse.ok) {
              content.push({
                type: "text",
                text: `\n❌ Failed to fetch image ${index}: ${imageResponse.statusText}`
              });
              continue;
            }

            // Get image data as buffer
            const arrayBuffer = await imageResponse.arrayBuffer();
            const imageBuffer = Buffer.from(arrayBuffer);
            const base64Data = imageBuffer.toString('base64');

            // Determine mime type from response headers or URL
            let mimeType = imageResponse.headers.get('content-type');

            // If content-type is generic or missing, try to determine from URL or default to PNG
            if (!mimeType || mimeType === 'application/octet-stream') {
              if (image.url.includes('.jpg') || image.url.includes('.jpeg')) {
                mimeType = 'image/jpeg';
              } else if (image.url.includes('.gif')) {
                mimeType = 'image/gif';
              } else if (image.url.includes('.webp')) {
                mimeType = 'image/webp';
              } else {
                mimeType = 'image/png';
              }
            }

            content.push({
              type: "text",
              text: `\n🖼️ Image ${index}: ${image.alt}`
            });

            content.push({
              type: "image",
              data: base64Data,
              mimeType: mimeType
            });

          } catch (error) {
            content.push({
              type: "text",
              text: `\n❌ Error fetching image ${index}: ${error.message}`
            });
          }
        }

        return { content };

      } catch (error) {
        console.error("Error in getPageImages:", error);
        throw new Error(`Failed to get page images: ${error.message}`);
      }
    }
  );

  server.tool(
    "getPageAttachments",
    "Fetch and view specific file attachments from a OneNote page (PDFs, Word docs, Excel files, etc.). Returns files as base64-encoded data that Claude can read. Use this after readPage shows you which attachments are available.",
    {
      pageId: z.string().describe("The ID of the page containing the attachments"),
      attachmentIndices: z.array(z.number()).describe("Array of attachment numbers to fetch (e.g., [1, 2] for first and second attachment)"),
      groupId: z.string().optional().describe("Optional: Group ID if known (makes it faster)")
    },
    async (params) => {
      try {
        await ensureGraphClient();
        const { pageId, attachmentIndices, groupId } = params;

        // Find the page and get its content
        let htmlContent;
        let pageTitle;

        if (groupId) {
          const pageDetails = await graphClient
            .api(`/groups/${groupId}/onenote/pages/${pageId}`)
            .get();

          const contentResponse = await fetch(pageDetails.contentUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          });
          htmlContent = await contentResponse.text();
          pageTitle = pageDetails.title;
        } else {
          // Try personal first, then groups
          try {
            const pageDetails = await graphClient
              .api(`/me/onenote/pages/${pageId}`)
              .get();

            const contentResponse = await fetch(pageDetails.contentUrl, {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            htmlContent = await contentResponse.text();
            pageTitle = pageDetails.title;
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
                htmlContent = await contentResponse.text();
                pageTitle = pageDetails.title;
                break;
              } catch (groupError) {
                continue;
              }
            }

            if (!htmlContent) {
              throw new Error("Page not found in personal notebooks or any group notebooks");
            }
          }
        }

        // Extract all attachments from the page
        const allAttachments = extractAttachments(htmlContent);

        if (allAttachments.length === 0) {
          return {
            content: [{
              type: "text",
              text: "No attachments found on this page."
            }]
          };
        }

        // Fetch the requested attachments
        const content = [{
          type: "text",
          text: `Attachments from "${pageTitle}":\n`
        }];

        for (const index of attachmentIndices) {
          if (index < 1 || index > allAttachments.length) {
            content.push({
              type: "text",
              text: `\n⚠️ Attachment ${index} not found (page has ${allAttachments.length} attachments)`
            });
            continue;
          }

          const attachment = allAttachments[index - 1];

          try {
            // Fetch the attachment with authentication
            const attachmentResponse = await fetch(attachment.url, {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            });

            if (!attachmentResponse.ok) {
              content.push({
                type: "text",
                text: `\n❌ Failed to fetch attachment ${index}: ${attachmentResponse.statusText}`
              });
              continue;
            }

            // Get file data as buffer
            const arrayBuffer = await attachmentResponse.arrayBuffer();
            const fileBuffer = Buffer.from(arrayBuffer);
            const base64Data = fileBuffer.toString('base64');

            // Get mime type
            let mimeType = attachment.type;
            if (!mimeType || mimeType === 'unknown') {
              mimeType = attachmentResponse.headers.get('content-type') || 'application/octet-stream';
            }

            content.push({
              type: "text",
              text: `\n📎 Attachment ${index}: ${attachment.name} (${(fileBuffer.length / 1024).toFixed(1)} KB)`
            });

            content.push({
              type: "resource",
              resource: {
                blob: base64Data,
                mimeType: mimeType
              }
            });

          } catch (error) {
            content.push({
              type: "text",
              text: `\n❌ Error fetching attachment ${index}: ${error.message}`
            });
          }
        }

        return { content };

      } catch (error) {
        console.error("Error in getPageAttachments:", error);
        throw new Error(`Failed to get page attachments: ${error.message}`);
      }
    }
  );
}
