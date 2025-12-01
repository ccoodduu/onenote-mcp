import { createGraphClient, checkAuthStatus, logout } from '../utils/graph-client.mjs';

export function registerAuthenticationTools(server) {
  server.tool(
    "authenticate",
    "Start the authentication flow with Microsoft Graph. Use this first before accessing OneNote content. IMPORTANT: After calling this tool, you MUST show the URL and code to the user, then call 'waitForAuth' to wait for them to complete authentication.",
    async () => {
      try {
        const result = await createGraphClient();
        if (result.type === 'device_code') {
          const info = result.deviceCodeInfo;
          let message = "Authentication started!\n\n";

          if (info) {
            message += `Please authenticate with Microsoft:\n\n`;
            message += `🌐 URL: ${info.verificationUri}\n`;
            message += `🔑 Code: ${info.userCode}\n\n`;
            message += `✅ Code has been copied to your clipboard!\n`;
            message += `The browser should open automatically. If not, click the URL above and paste the code.\n\n`;
            message += `After authenticating, use the "waitForAuth" tool to wait for completion.`;
          } else {
            message += "Please check the console for the URL and code.";
          }

          return {
            content: [
              {
                type: "text",
                text: message
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

  server.tool(
    "waitForAuth",
    "Wait for authentication to complete. Use this after calling 'authenticate' to ensure the user has completed the login process.",
    async () => {
      try {
        const maxAttempts = 60; // 60 attempts * 2 seconds = 2 minutes max wait
        const delayMs = 2000; // Check every 2 seconds

        for (let attempt = 0; attempt < maxAttempts; attempt++) {
          const isAuthenticated = await checkAuthStatus();

          if (isAuthenticated) {
            return {
              content: [
                {
                  type: "text",
                  text: "✅ Authentication completed successfully! You can now use OneNote tools."
                }
              ]
            };
          }

          // Wait before next check
          await new Promise(resolve => setTimeout(resolve, delayMs));
        }

        // Timeout
        return {
          content: [
            {
              type: "text",
              text: "⏱️ Authentication timeout. Please try calling 'authenticate' again."
            }
          ]
        };
      } catch (error) {
        console.error("Error waiting for authentication:", error);
        throw new Error(`Failed to wait for authentication: ${error.message}`);
      }
    }
  );

  server.tool(
    "logout",
    "Log out and clear the stored authentication token. Use this to sign out or switch accounts.",
    async () => {
      try {
        const result = logout();
        return {
          content: [
            {
              type: "text",
              text: result.success
                ? "✅ Logged out successfully. Token has been cleared."
                : "⚠️ No active session found."
            }
          ]
        };
      } catch (error) {
        console.error("Error during logout:", error);
        throw new Error(`Logout failed: ${error.message}`);
      }
    }
  );
}
