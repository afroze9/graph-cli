# graph-cli

A .NET global tool for interacting with Microsoft Graph — manage emails, calendar events, Teams chats, To Do tasks, presence, user directory, and OneDrive/SharePoint files from the command line or via [MCP](https://modelcontextprotocol.io/) for AI assistants like Claude. Output is JSON by default (`--format table` for human-readable output).

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
graph-cli auth list     # List all configured profiles
graph-cli auth logout   # Clear cached tokens
```

Tokens are cached at `~/.graph-cli/token-cache.bin` and auto-refresh silently — you won't need to log in again unless you explicitly log out.

### Profiles (multiple accounts)

Every account is a named **profile**. If you never name one, everything uses the profile `default`, so single-account users can ignore this entirely.

There is **no stateful "active account" to switch** — you select a profile per invocation, and it never mutates shared state. This means two shells (or two MCP servers) can use two different accounts at the same time without interfering. Selection precedence:

```
--profile <name>   >   GRAPH_CLI_PROFILE env var   >   "default"
```

```bash
# Log in a second account under a named profile
graph-cli auth login --profile work --tenant <tenant-id> --client <client-id>

# Use a specific profile for one command
graph-cli mail list --profile work

# Or set it for a whole shell session (no persisted state, no switch-back)
export GRAPH_CLI_PROFILE=work
graph-cli mail list

# See every profile and which one the current invocation resolves to
graph-cli auth list
```

Each profile is pinned to its Microsoft identity at login, so profiles that share one app registration still resolve deterministically.

## Using as an MCP Server (Claude Desktop, Cursor, etc.)

graph-cli includes a built-in [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server, so AI assistants like Claude can interact with your Microsoft 365 data directly.

### 1. Register an Azure AD App

Follow [step 1 from Setup](#1-register-an-azure-ad-app) above if you haven't already.

### 2. Log in first

Authenticate each account once from a terminal so its profile and tokens are cached (see [Profiles](#profiles-multiple-accounts)):

```bash
graph-cli auth login                                                   # → "default" profile
graph-cli auth login --profile work --tenant <id> --client <id>       # → a second account
```

The MCP server reuses these cached profiles; it never prompts for login itself.

### 3. Add to Claude Desktop

Edit your Claude Desktop config file:

- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

Bind **one server per account** and select the profile with `--profile`:

```json
{
  "mcpServers": {
    "graph-cli": {
      "command": "graph-cli",
      "args": ["mcp", "--profile", "default"]
    },
    "graph-cli-work": {
      "command": "graph-cli",
      "args": ["mcp", "--profile", "work"]
    }
  }
}
```

Each server is pinned to a single account for its whole lifetime — no state switching. Their tools are namespaced by the server label (`graph-cli` → `mcp__graph-cli__*`, `graph-cli-work` → `mcp__graph-cli-work__*`), so the assistant picks the account by choosing the tool. If you only have one account, keep the single `graph-cli` entry and drop `--profile` (it defaults to `default`).

> **Zero-config bootstrap:** if you'd rather not run `auth login` first, you can instead supply `GRAPH_CLI_TENANT_ID` and `GRAPH_CLI_CLIENT_ID` in the server's `env` block. A browser window opens for login on the first tool call. This path is single-account only; use profiles (above) for multiple accounts.

### 4. Add to Claude Code

Register the server with the `claude mcp add` command — same binary, same `--profile` flag:

```bash
# Default account
claude mcp add graph-cli -- graph-cli mcp --profile default

# A second account, as its own server
claude mcp add graph-cli-work -- graph-cli mcp --profile work
```

This writes to your Claude Code MCP config (use `-s user` to make it global across projects). Verify with `claude mcp list`. As in Desktop, each entry is one account and tools are namespaced by the server name.

### 5. Restart the client

Restart Claude Desktop (or reload Claude Code) and you should see graph-cli tools available. Ask something like "What's my email address?" or "Show my calendar for today" to verify.

### Troubleshooting

- **"Authentication required" / "No cached token"** — Run `graph-cli auth login` (add `--profile <name>` for a named account) in a terminal, then retry.
- **"Profile '<name>' not found"** — The server's `--profile` doesn't match a logged-in profile. Run `graph-cli auth list` to see configured profiles, and `graph-cli auth login --profile <name> --tenant <id> --client <id>` to add it.
- **"isn't pinned to an identity"** — A profile migrated from an older single-account config has multiple identities in the token cache. Run `graph-cli auth login --profile <name>` once to pin the correct one.
- **Tools not appearing** — Run `graph-cli auth status --profile <name>` in a terminal to confirm that profile authenticates.

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
| auth list | ✅ | — | — |
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

> **Note:** `chat reply` works for both 1:1 and group chats. It sends a quoted reply using a `messageReference` attachment, which Teams renders as a native reply with the original message quoted above.

> **Note:** `chat messages` includes a `reactions` array on each message, listing any emoji reactions (e.g. a 👍 thumbs-up). Each reaction reports `reactionType` (the emoji), `displayName` (a friendly label such as `Like`), `createdDateTime`, and `userId` (the AAD id of who reacted). This lets you detect when someone acknowledged a message with a reaction rather than a text reply.

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
  ├── ChatService.cs
  ├── FileService.cs
  ├── ...
  │
Commands/           ← CLI layer (arg parsing, output formatting)
  ├── UserCommands.cs
  ├── MailCommands.cs
  ├── ChatCommands.cs
  ├── ...
  │
McpTools/           ← MCP layer (JSON serialization, tool registration)
  ├── UserTools.cs
  ├── MailTools.cs
  ├── ChatTools.cs
  ├── FilesTools.cs
  ├── ...
```

## Global Options

| Option | Description |
|---|---|
| `--format json\|table` | Output format (default: `json`) |
| `--timezone <tz>` | Timezone for datetime I/O — accepts IANA (e.g. `Asia/Karachi`) or Windows IDs (e.g. `Pakistan Standard Time`). Defaults to system local timezone. |
| `--profile <name>` | Account profile to use for this invocation. Overrides `GRAPH_CLI_PROFILE`; defaults to `default`. |

## Environment Variables

| Variable | Description |
|---|---|
| `GRAPH_CLI_PROFILE` | Account profile to use when `--profile` is not passed (default: `default`) |
| `GRAPH_CLI_TENANT_ID` | Azure AD tenant ID (alternative to config.json) |
| `GRAPH_CLI_CLIENT_ID` | Azure AD client/application ID (alternative to config.json) |
| `GRAPH_CLI_SCOPES` | Comma-separated Microsoft Graph scopes (default: all required scopes) |
| `GRAPH_CLI_SKIP_ALLOWLIST` | Set to `true` to bypass the allowed contacts check for outbound actions (send email, chat, share) |

## License

MIT
