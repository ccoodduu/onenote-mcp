# OneNote MCP Server

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that lets AI assistants read, search, and navigate your Microsoft OneNote notebooks — including personal and shared/Teams notebooks.

No Azure setup required. Uses Microsoft's device code authentication flow.

> Based on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks.

## Features

- **Browse** notebooks, section groups, sections, and pages (personal + Teams/shared)
- **Read** full page content with text extraction from HTML
- **View images** embedded in pages (returned as base64 for AI analysis)
- **Download attachments** (PDFs, Word docs, Excel, etc.) for analysis
- **Search** notebooks by name and pages by title
- **Open pages** directly in OneNote app or browser
- **Smart caching** to minimize API calls
- **Device code auth** — no Azure app registration needed

## Quick Start

### Install

```bash
npm install -g mcp-server-onenote
```

Or use directly with npx:

```bash
npx mcp-server-onenote
```

### Configure Your AI Client

Add the server to your MCP client configuration:

<details>
<summary><strong>Claude Desktop</strong></summary>

Edit `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "npx",
      "args": ["-y", "mcp-server-onenote"]
    }
  }
}
```
</details>

<details>
<summary><strong>Claude Code</strong></summary>

```bash
claude mcp add onenote -- npx -y mcp-server-onenote
```
</details>

<details>
<summary><strong>Cursor</strong></summary>

Add to your MCP settings (Ctrl+, → MCP tab):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "npx",
      "args": ["-y", "mcp-server-onenote"]
    }
  }
}
```
</details>

<details>
<summary><strong>Other MCP Clients</strong></summary>

Use the command `npx -y mcp-server-onenote` with stdio transport.
</details>

### Authenticate

The first time you use a OneNote tool, the AI will guide you through authentication:

1. Ask the AI to authenticate with OneNote
2. A device code is copied to your clipboard and your browser opens automatically
3. Paste the code at the Microsoft login page and sign in
4. Tell the AI to check authentication status (calls `waitForAuth`)
5. Done — your token is saved locally for future sessions

## Available Tools

### Authentication

| Tool | Description |
|------|-------------|
| `authenticate` | Start Microsoft device code login flow |
| `waitForAuth` | Poll for authentication completion (call after `authenticate`) |
| `logout` | Clear stored token and sign out |

### Discovery & Browsing

| Tool | Description |
|------|-------------|
| `listGroups` | List Microsoft 365 Groups and Teams you belong to |
| `listAllNotebooks` | List all notebooks (personal + shared). Optional: filter by `groupId` |
| `listAllSections` | List sections in notebooks. Optional: filter by `groupId`, `notebookId`, `sectionGroupId` |
| `listAllPages` | List pages. Optional: filter by `groupId`, `sectionId` |
| `browseOneNote` | Interactive drill-down navigation (notebooks → sections → pages) |

### Content

| Tool | Description |
|------|-------------|
| `readPage` | Read full text content of a page. Params: `pageId`, optional `groupId` |
| `getPageImages` | Fetch embedded images as base64. Params: `pageId`, `imageIndices`, optional `groupId` |
| `getPageAttachments` | Download file attachments. Params: `pageId`, `attachmentIndices`, optional `groupId` |

### Search

| Tool | Description |
|------|-------------|
| `searchNotebooks` | Search notebooks by name |
| `searchPagesInNotebook` | Search pages by title within a notebook |

### Actions

| Tool | Description |
|------|-------------|
| `openPage` | Open a page in OneNote app or browser |
| `clearCache` | Clear cached API responses to see fresh data |

## Example Interactions

```
You: Show me my OneNote notebooks
AI: [calls listAllNotebooks] You have 3 notebooks: "Work", "Personal", and "Team Notes" (shared)

You: What's in my Work notebook?
AI: [calls browseOneNote] Your Work notebook has sections: "Meetings", "Projects", "Notes"

You: Read my latest meeting notes
AI: [calls readPage] Here's the content from "Sprint Planning 2026-03-15"...

You: Are there any images in that page?
AI: [calls getPageImages] Yes, here's the architecture diagram from the page...

You: Search for anything about "roadmap" in my Projects notebook
AI: [calls searchPagesInNotebook] Found 3 pages matching "roadmap"...

You: Open the Q2 roadmap in my browser
AI: [calls openPage] Opened "Q2 Roadmap" in your browser
```

## Security

- Auth tokens are stored locally in `.access-token.txt` (git-ignored)
- Uses Microsoft's public OAuth app — no secrets to manage
- Tokens expire automatically and require re-authentication
- You can sign out anytime with the `logout` tool

## Troubleshooting

**Authentication fails:** Make sure you're using a browser without aggressive tracking prevention. Try clearing cookies, or use a private/incognito window.

**Token expired:** The AI will tell you if the token is invalid. Just ask it to authenticate again.

**Stale data:** OneNote API responses are cached. Ask the AI to clear the cache (`clearCache`) or wait for the cache TTL to expire (5–60 minutes depending on data type).

**Can't see shared notebooks:** Make sure you have access to the Microsoft 365 Group. Use `listGroups` to see which groups are available.

## Development

```bash
git clone https://github.com/ccoodduu/onenote-mcp-unified.git
cd onenote-mcp-unified
npm install
npm start
```

## Credits

Built on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks.

## License

MIT
