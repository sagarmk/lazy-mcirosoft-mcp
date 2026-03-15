<p align="center">
  <img src="assets/logo.svg" width="128" height="128" alt="Lazy MS Graph MCP" />
</p>

<h1 align="center">Lazy MS Graph MCP</h1>

<p align="center">
  Access Microsoft 365 — mail, calendar, users, files, teams, and more — directly from Claude, Cursor, or VS Code.
</p>

<p align="center">
  <a href="https://github.com/sagarmk/lazy-mcirosoft-mcp/releases/latest">
    <img src="https://img.shields.io/github/v/release/sagarmk/lazy-mcirosoft-mcp?label=Download%20DMG&style=for-the-badge" alt="Download" />
  </a>
</p>

---

## Quick Start (macOS App)

**[Download the latest DMG](https://github.com/sagarmk/lazy-mcirosoft-mcp/releases/latest)** and drag it to Applications. That's it.

The app gives you:
- One-click server start/stop
- Encrypted credential storage
- Auto-install buttons for Claude Code and Cursor
- Built-in Azure AD setup guide
- Live server logs

---

## Quick Start (npm)

If you prefer running the MCP server directly:

### 1. Clone and build

```bash
git clone https://github.com/sagarmk/lazy-mcirosoft-mcp.git
cd lazy-ms-graph-mcp
npm install
npm run build
```

### 2. Set up Azure AD credentials

You need an Azure AD app registration with **Application permissions** (not Delegated). See [Azure Setup](#azure-ad-setup) below.

### 3. Run the server

```bash
MSGRAPH_CLIENT_ID="your-client-id" \
MSGRAPH_CLIENT_SECRET="your-client-secret" \
MSGRAPH_TENANT_ID="your-tenant-id" \
npm start
```

### 4. Add to your IDE

**Claude Code** — add to `~/.claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "lazy-ms-graph": {
      "command": "node",
      "args": ["/path/to/lazy-ms-graph-mcp/packages/server/dist/index.js"],
      "env": {
        "MSGRAPH_CLIENT_ID": "your-client-id",
        "MSGRAPH_CLIENT_SECRET": "your-client-secret",
        "MSGRAPH_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

**Cursor** — add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "lazy-ms-graph": {
      "command": "node",
      "args": ["/path/to/lazy-ms-graph-mcp/packages/server/dist/index.js"],
      "env": {
        "MSGRAPH_CLIENT_ID": "your-client-id",
        "MSGRAPH_CLIENT_SECRET": "your-client-secret",
        "MSGRAPH_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

**VS Code** — add to `.vscode/mcp.json`:

```json
{
  "servers": {
    "lazy-ms-graph": {
      "type": "stdio",
      "command": "node",
      "args": ["/path/to/lazy-ms-graph-mcp/packages/server/dist/index.js"],
      "env": {
        "MSGRAPH_CLIENT_ID": "your-client-id",
        "MSGRAPH_CLIENT_SECRET": "your-client-secret",
        "MSGRAPH_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

---

## What's Included

53 tools across 10 Microsoft Graph categories:

| Category | Tools | Examples |
|----------|-------|---------|
| **Mail** | 5 | Send emails, list inbox, read messages |
| **Calendar** | 6 | Create events, find available times, manage meetings |
| **Users** | 5 | List users, get profiles, create/update accounts |
| **Files** | 5 | Browse OneDrive, upload/download files, search |
| **Teams** | 5 | List teams, channels, members |
| **Contacts** | 5 | Manage contacts |
| **Groups** | 5 | Create groups, manage membership |
| **Planner** | 4 | Manage tasks and plans |
| **SharePoint** | 3 | Browse sites and lists |
| **OneNote** | 2 | List notebooks, create pages |

The server exposes 2 MCP tools: `search_tools` (discover available tools) and `execute_tool` (run them). Your AI assistant uses search to find the right tool, then executes it.

---

## Azure AD Setup

### 1. Register an app

1. Go to [portal.azure.com](https://portal.azure.com) > **App registrations** > **New registration**
2. Name it anything (e.g. "Lazy MS Graph MCP")
3. Click **Register**
4. Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page

### 2. Create a client secret

1. Go to **Certificates & secrets** > **New client secret**
2. Copy the **Value** immediately (shown only once)

### 3. Add Application permissions

> **IMPORTANT**: You must add **Application** permissions, not **Delegated** permissions. Delegated permissions will not work and all API calls will return 403 errors.

1. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**
2. Add the permissions you need:
   - `User.Read.All` — for user tools
   - `Mail.Read`, `Mail.Send` — for mail tools
   - `Calendars.ReadWrite` — for calendar tools
   - `Files.ReadWrite.All` — for file tools
   - `Team.ReadBasic.All` — for teams tools
3. Click **Grant admin consent**

---

## Environment Variables

| Variable | Fallback | Required |
|----------|----------|----------|
| `MSGRAPH_CLIENT_ID` | `MICROSOFT_MCP_CLIENT_ID` | Yes |
| `MSGRAPH_CLIENT_SECRET` | `MICROSOFT_MCP_CLIENT_SECRET` | Yes |
| `MSGRAPH_TENANT_ID` | `MICROSOFT_MCP_TENANT_ID` | Yes |
| `MSGRAPH_MCP_PORT` | `MICROSOFT_MCP_PORT` | No (default: 3100) |
| `MSGRAPH_MCP_TRANSPORT` | `MICROSOFT_MCP_TRANSPORT` | No (default: stdio) |

---

## License

MIT
