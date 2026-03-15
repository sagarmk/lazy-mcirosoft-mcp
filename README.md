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

## How It Works

The server exposes **2 MCP tools** to your AI assistant:

1. **`search_tools`** — discover available Microsoft Graph tools by category or keyword
2. **`execute_tool`** — run a tool by name with parameters

Your AI assistant browses by category, picks the right tool, and executes it:

```
You:    "Send an email to john@company.com about the meeting"

AI:     → search_tools({query: "*", category: "mail"})
        ← returns: mail_list_messages, mail_get_message, mail_send,
                   mail_create_draft, mail_delete_message

AI:     → execute_tool({tool_name: "mail_send", parameters: {
            userId: "me",
            subject: "About the meeting",
            body: "...",
            toRecipients: ["john@company.com"]
          }})
        ← email sent
```

This keeps the MCP surface area small (just 2 tools) while giving access to all 53 Microsoft Graph operations — important when you have multiple MCP servers loaded.

---

## All Available Tools

53 tools across 10 categories:

| Category | Tools | Operations |
|----------|-------|-----------|
| **users** | 5 | `users_list`, `users_get`, `users_create`, `users_update`, `users_delete` |
| **mail** | 5 | `mail_list_messages`, `mail_get_message`, `mail_send`, `mail_create_draft`, `mail_delete_message` |
| **calendar** | 6 | `calendar_list_events`, `calendar_get_event`, `calendar_create_event`, `calendar_update_event`, `calendar_delete_event`, `calendar_find_available_times` |
| **contacts** | 5 | `contacts_list`, `contacts_get`, `contacts_create`, `contacts_update`, `contacts_delete` |
| **files** | 5 | `files_list_in_drive`, `files_get`, `files_upload`, `files_delete`, `files_search` |
| **teams** | 5 | `teams_list`, `teams_get`, `teams_create`, `teams_list_members`, `teams_list_channels` |
| **groups** | 5 | `groups_list`, `groups_get`, `groups_create`, `groups_add_member`, `groups_list_members` |
| **planner** | 4 | `planner_list_tasks`, `planner_get_task`, `planner_create_task`, `planner_update_task` |
| **sharepoint** | 3 | `sharepoint_list_sites`, `sharepoint_get_site`, `sharepoint_list_items` |
| **onenote** | 2 | `onenote_list_notebooks`, `onenote_create_page` |

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
