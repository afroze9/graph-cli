# graph-cli

A .NET global tool for interacting with Microsoft Graph — manage emails, calendar events, Teams chats, Teams channel messages, To Do tasks, presence, user directory, and OneDrive/SharePoint files from the command line or via [MCP](https://modelcontextprotocol.io/) for AI assistants like Claude. Output is JSON by default (`--format table` for human-readable output).

## Installation

Requires [.NET 10 SDK](https://dotnet.microsoft.com/download) or later.

```bash
# Install from NuGet
dotnet tool install -g graph-cli
```

Or install from source:

```bash
git clone https://github.com/afroze9/graph-cli.git
cd graph-cli
dotnet pack -o ./nupkg
dotnet tool install -g graph-cli --add-source ./nupkg
```

## Setup (First-Time Users)

### 1. Register an Azure AD App

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and click **New registration**
2. Name it something like `graph-cli`, set **Supported account types** to "Single tenant", and set **Redirect URI** to `http://localhost` (type: Public client/native)
3. After creation, copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page
4. Go to **API permissions → Add a permission → Microsoft Graph → Delegated permissions** and add:
   - `User.Read`, `User.ReadBasic.All`
   - `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read.Shared`, `Calendars.ReadWrite`
   - `Chat.Create`, `Chat.ReadWrite`, `ChatMessage.Read`, `ChatMessage.Send`
   - `Presence.Read.All`
   - `Tasks.ReadWrite`
   - `Files.Read.All`, `Sites.ReadWrite.All` (or `Sites.ReadWrite.All` for page create/update/publish)
   - `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send` _(required for Teams channel commands)_
5. Click **Grant admin consent** (or ask your tenant admin to do this)

### 2. Configure graph-cli

Create `~/.graph-cli/config.json`:

```json
{
  "tenantId": "<your-tenant-id>",
  "clientId": "<your-client-id>"
}
```

Or set environment variables `GRAPH_CLI_TENANT_ID` and `GRAPH_CLI_CLIENT_ID` instead.

### 3. Authenticate

```bash
graph-cli auth login    # Opens browser for interactive auth (only needed once)
graph-cli auth status   # Check if authenticated
graph-cli auth logout   # Clear cached tokens
```

Tokens are cached at `~/.graph-cli/token-cache.bin` and auto-refresh silently — you won't need to log in again unless you explicitly log out.

## Using as an MCP Server (Claude Desktop, Cursor, etc.)

graph-cli includes a built-in [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server, so AI assistants like Claude can interact with your Microsoft 365 data directly.

### 1. Register an Azure AD App

Follow [step 1 from Setup](#1-register-an-azure-ad-app) above if you haven't already.

### 2. Add to Claude Desktop

Edit your Claude Desktop config file:

- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "graph-cli": {
      "command": "graph-cli",
      "args": ["mcp"],
      "env": {
        "GRAPH_CLI_TENANT_ID": "<your-tenant-id>",
        "GRAPH_CLI_CLIENT_ID": "<your-client-id>",
        "GRAPH_CLI_SCOPES": "User.Read,User.ReadBasic.All,Mail.ReadWrite,Mail.Send,Calendars.Read.Shared,Calendars.ReadWrite,Chat.Create,Chat.ReadWrite,ChatMessage.Read,ChatMessage.Send,Presence.Read.All,Tasks.ReadWrite,Files.Read.All,Sites.ReadWrite.All,Team.ReadBasic.All,Channel.ReadBasic.All,ChannelMessage.Read.All,ChannelMessage.Send"
      }
    }
  }
}
```

No `config.json` needed — the MCP config provides everything. On first tool call, a browser window will open for Microsoft login. After that, tokens are cached and refresh automatically.

### 3. Restart Claude Desktop

You should see graph-cli tools available. Ask Claude something like "What's my email address?" or "Show my calendar for today" to verify.

### Troubleshooting

- **"Authentication required"** — Run `graph-cli auth login` in a terminal, then retry.
- **"No configuration found"** — Check that `GRAPH_CLI_TENANT_ID` and `GRAPH_CLI_CLIENT_ID` are set in your MCP config, or create `~/.graph-cli/config.json` (see [Setup](#2-configure-graph-cli)).
- **Tools not appearing** — Run `graph-cli auth status` in a terminal to check that auth is working.

## Quick Start

```bash
# Check your profile
graph-cli user me --format table

# See latest emails
graph-cli mail list --top 5 --format table

# Check today's calendar
graph-cli calendar events --format table

# Allow a contact before sending
graph-cli contacts allow jane@company.com --actions email,chat
graph-cli mail send --to jane@company.com --subject "Hello" --body "Hi Jane"
```

## Commands

### Command Availability

All commands are available via both the CLI and the MCP server unless noted otherwise.

| Command | CLI | MCP | MCP Tool Name |
|---|:---:|:---:|---|
| **Auth** | | | |
| auth login | ✅ | — | — |
| auth status | ✅ | ✅ | `auth_status` |
| auth logout | ✅ | — | — |
| **User** | | | |
| user me | ✅ | ✅ | `user_me` |
| user get | ✅ | ✅ | `user_get` |
| user search | ✅ | ✅ | `user_search` |
| user manager | ✅ | ✅ | `user_manager` |
| user reports | ✅ | ✅ | `user_reports` |
| **Mail** | | | |
| mail list | ✅ | ✅ | `mail_list` |
| mail get | ✅ | ✅ | `mail_get` |
| mail search | ✅ | ✅ | `mail_search` |
| mail send | ✅ | ✅ | `mail_send` |
| mail draft | ✅ | ✅ | `mail_draft` |
| mail send-draft | ✅ | ✅ | `mail_send_draft` |
| mail reply | ✅ | ✅ | `mail_reply` |
| mail reply-all | ✅ | ✅ | `mail_reply` (with replyAll=true) |
| mail forward | ✅ | ✅ | `mail_forward` |
| mail move | ✅ | ✅ | `mail_move` |
| mail delete | ✅ | ✅ | `mail_delete` |
| mail mark-read | ✅ | ✅ | `mail_mark_read` |
| mail folders | ✅ | ✅ | `mail_folders` |
| mail attachments | ✅ | ✅ | `mail_attachments` |
| mail download-attachment | ✅ | — | — |
| **Calendar** | | | |
| calendar list | ✅ | ✅ | `calendar_list` |
| calendar events | ✅ | ✅ | `calendar_events` |
| calendar get-event | ✅ | ✅ | `calendar_get_event` |
| calendar create-event | ✅ | ✅ | `calendar_create_event` |
| calendar update-event | ✅ | ✅ | `calendar_update_event` |
| calendar delete-event | ✅ | ✅ | `calendar_delete_event` |
| calendar respond | ✅ | ✅ | `calendar_respond` |
| calendar find-times | ✅ | ✅ | `calendar_find_times` |
| calendar schedule | ✅ | ✅ | `calendar_schedule` |
| **Chat (Teams)** | | | |
| chat list | ✅ | ✅ | `chat_list` |
| chat search | ✅ | ✅ | `chat_search` |
| chat find-with | ✅ | ✅ | `chat_find_with` |
| chat get | ✅ | ✅ | `chat_get` |
| chat create | ✅ | ✅ | `chat_create` |
| chat members | ✅ | ✅ | `chat_members` |
| chat messages | ✅ | ✅ | `chat_messages` |
| chat send | ✅ | ✅ | `chat_send` |
| chat send-image | ✅ | ✅ | `chat_send_image` |
| chat reply | ✅ | ✅ | `chat_reply` |
| **Teams Channels** | | | |
| teams list | ✅ | ✅ | `teams_list` |
| teams channels | ✅ | ✅ | `teams_channels_list` |
| teams messages | ✅ | ✅ | `teams_channel_messages` |
| teams send | ✅ | ✅ | `teams_channel_send` |
| teams reply | ✅ | ✅ | `teams_channel_reply` |
| **Presence** | | | |
| presence me | ✅ | ✅ | `presence_me` |
| presence get | ✅ | ✅ | `presence_get` |
| presence batch | ✅ | ✅ | `presence_batch` |
| **Tasks (To Do)** | | | |
| task lists | ✅ | ✅ | `task_lists` |
| task list | ✅ | ✅ | `task_list` |
| task create | ✅ | ✅ | `task_create` |
| task update | ✅ | ✅ | `task_update` |
| task delete | ✅ | ✅ | `task_delete` |
| task complete | ✅ | ✅ | `task_complete` |
| **Files (OneDrive/SharePoint)** | | | |
| files list | ✅ | ✅ | `files_list` |
| files get | ✅ | ✅ | `files_get` |
| files search | ✅ | ✅ | `files_search` |
| files share | ✅ | ✅ | `files_share` |
| files download | ✅ | ✅ | `files_download` |
| **Contacts (Allow-List)** | | | |
| contacts list | ✅ | ✅ | `contacts_list` |
| contacts allow | ✅ | ✅ | `contacts_allow` |
| contacts remove | ✅ | ✅ | `contacts_remove` |
| **Sites (SharePoint)** | | | |
| sites search | ✅ | ✅ | `sites_search` |
| **Pages (SharePoint)** | | | |
| pages list | ✅ | ✅ | `pages_list` |
| pages get | ✅ | ✅ | `pages_get` |
| pages create | ✅ | ✅ | `pages_create` |
| pages update | ✅ | ✅ | `pages_update` |
| pages publish | ✅ | ✅ | `pages_publish` |
| pages sections list | ✅ | ✅ | `pages_sections_list` |
| pages sections get | ✅ | ✅ | `pages_sections_get` |
| pages sections create | ✅ | ✅ | `pages_sections_create` |
| pages sections update | ✅ | ✅ | `pages_sections_update` |
| pages sections delete | ✅ | ✅ | `pages_sections_delete` |
| pages columns list | ✅ | ✅ | `pages_columns_list` |
| pages webparts list | ✅ | ✅ | `pages_webparts_list` |
| pages webparts get | ✅ | ✅ | `pages_webparts_get` |
| pages webparts create | ✅ | ✅ | `pages_webparts_create` |
| pages webparts update | ✅ | ✅ | `pages_webparts_update` |
| pages webparts delete | ✅ | ✅ | `pages_webparts_delete` |
| **Lists (SharePoint)** | | | |
| lists list | ✅ | ✅ | `lists_list` |
| lists items | ✅ | ✅ | `lists_items` |
| lists columns | ✅ | ✅ | `lists_columns` |

> Commands marked **—** in MCP are either interactive-only (`auth login`, `auth logout`) or involve file system write operations (`download-attachment`) that require explicit user confirmation outside the MCP context.

### Mail

```bash
graph-cli mail list [--top <n>] [--folder <name>]
graph-cli mail get <message-id>
graph-cli mail search --query <text> [--top <n>]
graph-cli mail send --to <emails> --subject <text> --body <text> [--cc <emails>] [--content-type text|html] [--attachment <file> ...]
graph-cli mail draft --to <emails> --subject <text> --body <text>
graph-cli mail send-draft <message-id>
graph-cli mail reply <message-id> --body <text> [--cc <emails>] [--bcc <emails>] [--content-type text|html] [--attachment <file> ...] [--draft]
graph-cli mail reply-all <message-id> --body <text> [--cc <emails>] [--bcc <emails>] [--content-type text|html] [--attachment <file> ...] [--draft]
graph-cli mail forward <message-id> --to <emails> --body <text> [--cc <emails>] [--bcc <emails>] [--content-type text|html] [--attachment <file> ...] [--draft]
graph-cli mail mark-read <message-id> [<id2> ...] [--unread]
graph-cli mail move <message-id> [<id2> ...] --folder <folder-id-or-name>
graph-cli mail delete <message-id> [<id2> ...]
graph-cli mail folders [--parent <folder-id-or-name>]
graph-cli mail attachments <message-id>
graph-cli mail download-attachment <message-id> <attachment-id> [--out <path>]
```

### Calendar

```bash
graph-cli calendar list
graph-cli calendar events [--start <iso-date>] [--end <iso-date>] [--calendar-id <id>] [--top <n>]
graph-cli calendar get-event <event-id>
graph-cli calendar create-event --subject <text> --start <iso-datetime> --end <iso-datetime> \
    [--attendees <emails>] [--body <text>] [--content-type text|html] \
    [--categories <names>] [--location <text>] [--online-meeting] [--calendar-id <id>]
graph-cli calendar update-event <event-id> [--subject <text>] [--start <datetime>] [--end <datetime>] \
    [--body <text>] [--content-type text|html] [--categories <names>]
graph-cli calendar delete-event <event-id>
graph-cli calendar respond <event-id> --action accept|decline|tentative [--comment <text>]
graph-cli calendar find-times --attendees <emails> --duration <minutes> [--start <iso-datetime>] [--end <iso-datetime>]
graph-cli calendar schedule --users <emails> --start <iso-datetime> --end <iso-datetime>
```

### Chat (Teams)

```bash
graph-cli chat list [--top <n>]
graph-cli chat search --query <text> [--top <n>] [--refresh] [--max-depth <n>]
graph-cli chat find-with --user <email-or-id> [--type oneOnOne|group|all] [--top <n>]
graph-cli chat get <chat-id>
graph-cli chat create --members <emails> [--topic <text>] [--type oneOnOne|group]
graph-cli chat members <chat-id>
graph-cli chat messages <chat-id> [--top <n>]
graph-cli chat send <chat-id> --message <text> [--content-type text|html] [--mentions <emails>]
graph-cli chat send-image <chat-id> --image <path> [--caption <text>]
graph-cli chat reply <chat-id> <message-id> --message <text> [--content-type text|html] [--mentions <emails>]
```

### Teams Channels

> **Note:** Teams channel commands use a different Graph API endpoint (`/teams/{teamId}/channels/...`) from `chat` commands which use the `Me/Chats` API. Use `chat` for 1:1 and group chats; use `teams` for messages inside a Teams channel.

Additional Graph scopes required: `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`

```bash
graph-cli teams list [--top <n>]
graph-cli teams channels <team-id> [--top <n>]
graph-cli teams messages <team-id> <channel-id> [--top <n>]
graph-cli teams send <team-id> <channel-id> --message <text> [--content-type text|html] [--mentions <emails>]
graph-cli teams reply <team-id> <channel-id> <message-id> --message <text> [--content-type text|html] [--mentions <emails>]
```

### Presence

```bash
graph-cli presence me
graph-cli presence get <user-id>
graph-cli presence batch --user-ids <comma-separated-ids>
```

### Tasks (Microsoft To Do)

```bash
graph-cli task lists
graph-cli task list <list-id> [--status notStarted|inProgress|completed]
graph-cli task create <list-id> --title <text> [--due <iso-date>] [--importance low|normal|high] [--body <text>]
graph-cli task update <list-id> <task-id> [--title <text>] [--status notStarted|inProgress|completed] [--due <date>] [--importance low|normal|high]
graph-cli task complete <list-id> <task-id>
graph-cli task delete <list-id> <task-id>
```

### Files (OneDrive & SharePoint)

```bash
graph-cli files list [--folder <item-id>] [--drive-id <id>] [--site <site-id-or-hostname>] [--top <n>]
graph-cli files get <item-id-or-sharing-url> [--drive-id <id>] [--site <site>]
graph-cli files download <item-id-or-sharing-url> [--out <path>] [--drive-id <id>] [--site <site>]
graph-cli files search <query> [--drive-id <id>] [--site <site>] [--top <n>] [--refresh]
graph-cli files share <item-id-or-sharing-url> --recipients <emails> [--role read|write|owner] [--message <text>] [--drive-id <id>] [--site <site>]
```

### User Directory

```bash
graph-cli user me
graph-cli user get <user-id-or-email>
graph-cli user search --query <text>
graph-cli user manager
graph-cli user reports
```

### Contacts Allow-List

Outbound actions (mail send, chat send, file share, etc.) are gated by an allowed contacts list for safety.

```bash
graph-cli contacts allow <email-or-group> --actions email,chat,share [--name "Display Name"] [--type user|group]
graph-cli contacts list [--type user|group]
graph-cli contacts remove <email-or-group>
```

### SharePoint Sites, Pages & Lists

```bash
# Sites
graph-cli sites search <query> [--top <n>] [--refresh]

# Pages — CRUD and publish
graph-cli pages list --site <site> [--top <n>] [--search <text>]
graph-cli pages get <page-id> --site <site> [--expand-content]
graph-cli pages create --site <site> --title <text> [--name <filename.aspx>] [--content <html>] [--publish]
graph-cli pages update <page-id> --site <site> [--title <text>] [--content <html>] [--publish]
graph-cli pages publish <page-id> --site <site>

# Sections — manage page layout sections
graph-cli pages sections list --site <site> --page-id <id>
graph-cli pages sections get <section-id> --site <site> --page-id <id>
graph-cli pages sections create --site <site> --page-id <id> --layout fullWidth|twoColumn|threeColumn|oneThirdLeftColumn|oneThirdRightColumn [--emphasis none|neutral|soft|strong]
graph-cli pages sections update <section-id> --site <site> --page-id <id> [--layout <layout>] [--emphasis <emphasis>]
graph-cli pages sections delete <section-id> --site <site> --page-id <id>

# Columns — inspect columns within a section
graph-cli pages columns list --site <site> --page-id <id> --section-id <id>

# WebParts — manage content within columns
graph-cli pages webparts list --site <site> --page-id <id> [--section-id <id> --column-id <id>]
graph-cli pages webparts get <webpart-id> --site <site> --page-id <id> --section-id <id> --column-id <id>
graph-cli pages webparts create --site <site> --page-id <id> --section-id <id> --column-id <id> --inner-html <html>
graph-cli pages webparts create --site <site> --page-id <id> --section-id <id> --column-id <id> --webpart-type <guid> [--data-json <json>]
graph-cli pages webparts update <webpart-id> --site <site> --page-id <id> --section-id <id> --column-id <id> [--inner-html <html>] [--data-json <json>]
graph-cli pages webparts delete <webpart-id> --site <site> --page-id <id> --section-id <id> --column-id <id>

# Lists
graph-cli lists list --site <site>
graph-cli lists items <list-id> --site <site> [--top <n>] [--fields <field1,field2>] [--filter <odata-filter>] [--expand-lookups <col1,col2>]
graph-cli lists columns <list-id> --site <site>   # column definitions: displayName + internal name
```

## Architecture

graph-cli uses a layered architecture where core business logic lives in shared services, exposed via both the CLI and MCP interfaces:

```
Services/           ← Core logic (Graph SDK calls)
  ├── UserService.cs
  ├── MailService.cs
  ├── CalendarService.cs
  ├── ChatService.cs      (1:1 and group chats — Me/Chats API)
  ├── TeamsService.cs     (Teams channel messages — /teams API)
  ├── FileService.cs
  ├── ...
  │
Commands/           ← CLI layer (arg parsing, output formatting)
  ├── UserCommands.cs
  ├── MailCommands.cs
  ├── ChatCommands.cs
  ├── TeamsCommands.cs
  ├── ...
  │
McpTools/           ← MCP layer (JSON serialization, tool registration)
  ├── UserTools.cs
  ├── MailTools.cs
  ├── ChatTools.cs
  ├── TeamsTools.cs
  ├── FilesTools.cs
  ├── ...
```

## Global Options

| Option | Description |
|---|---|
| `--format json\|table` | Output format (default: `json`) |
| `--timezone <tz>` | Timezone for datetime I/O — accepts IANA (e.g. `Asia/Karachi`) or Windows IDs (e.g. `Pakistan Standard Time`). Defaults to system local timezone. |

## Environment Variables

| Variable | Description |
|---|---|
| `GRAPH_CLI_TENANT_ID` | Azure AD tenant ID (alternative to config.json) |
| `GRAPH_CLI_CLIENT_ID` | Azure AD client/application ID (alternative to config.json) |
| `GRAPH_CLI_SCOPES` | Comma-separated Microsoft Graph scopes (default: all required scopes) |
| `GRAPH_CLI_SKIP_ALLOWLIST` | Set to `true` to bypass the allowed contacts check for outbound actions (send email, chat, share) |

## License

MIT
