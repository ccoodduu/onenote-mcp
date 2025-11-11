import { createGraphClient } from '../utils/graph-client.mjs';

export function registerAuthenticationTools(server) {
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
}
