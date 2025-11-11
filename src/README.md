# OneNote MCP - Source Structure

This directory contains the modular implementation of the OneNote MCP server.

## Directory Structure

```
src/
├── index.mjs                   # Main entry point, registers all tools
├── auth/
│   └── authentication.mjs      # Authentication tool
├── tools/
│   ├── discovery.mjs           # listGroups
│   ├── browsing.mjs            # listAllNotebooks, listAllSections, listAllPages, browseOneNote
│   ├── content.mjs             # readPage
│   ├── search.mjs              # searchNotebooks, searchAllPages
│   └── actions.mjs             # openPage
└── utils/
    └── graph-client.mjs        # Microsoft Graph API client setup and utilities
```

## Module Descriptions

### `auth/authentication.mjs`
Registers the `authenticate` tool for device code flow authentication with Microsoft Graph.

### `tools/discovery.mjs`
Provides the `listGroups` tool to discover Microsoft 365 Groups and Teams.

### `tools/browsing.mjs`
Core browsing functionality:
- `listAllNotebooks` - List all notebooks (personal + groups)
- `listAllSections` - List sections with support for section groups
- `listAllPages` - List pages from sections
- `browseOneNote` - Smart layer-by-layer navigation

### `tools/content.mjs`
Content reading functionality:
- `readPage` - Read page content with auto-detection of personal/group notebooks

### `tools/search.mjs`
Search functionality across all notebooks:
- `searchNotebooks` - Search for notebooks by name
- `searchAllPages` - Search for pages by title

### `tools/actions.mjs`
Action tools:
- `openPage` - Open pages in OneNote app or web browser

### `utils/graph-client.mjs`
Shared utilities for:
- Microsoft Graph API client initialization
- Access token management
- Authentication state handling

## Adding New Tools

To add a new tool:

1. Create or edit the appropriate file in `tools/`
2. Export a `register*Tools(server)` function
3. Import and call it in `src/index.mjs`

Example:
```javascript
// tools/my-new-tools.mjs
export function registerMyNewTools(server) {
  server.tool("myTool", "Description", async (params) => {
    // Implementation
  });
}

// src/index.mjs
import { registerMyNewTools } from './tools/my-new-tools.mjs';
// ...
registerMyNewTools(server);
```
