import { ensureGraphClient, graphClient } from '../utils/graph-client.mjs';

export function registerDiscoveryTools(server) {
  server.tool(
    "listGroups",
    "List all Microsoft 365 Groups and Teams you belong to (school, work, collaborative spaces). Use this to find group IDs needed for accessing shared notebooks.",
    async () => {
      try {
        await ensureGraphClient();
        const response = await graphClient.api("/me/memberOf/$/microsoft.graph.group").get();

        const groups = response.value.map(group => ({
          id: group.id,
          displayName: group.displayName,
          description: group.description,
          mail: group.mail
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(groups, null, 2)
            }
          ]
        };
      } catch (error) {
        console.error("Error listing groups:", error);
        throw new Error(`Failed to list groups: ${error.message}`);
      }
    }
  );
}
