# graph-cli

A .NET global tool for interacting with Microsoft Graph — manage emails, calendar events, Teams chats, To Do tasks, presence, and user directory from the command line. Output is JSON by default (`--format table` for human-readable output).

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
   - `Files.Read.All`, `Sites.Read.All`
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

### Mail

```bash
graph-cli mail list [--top <n>] [--folder <name>]
graph-cli mail get <message-id>
graph-cli mail search --query <text> [--top <n>]
graph-cli mail send --to <emails> --subject <text> --body <text> [--cc <emails>] [--content-type text|html]
graph-cli mail draft --to <emails> --subject <text> --body <text>
graph-cli mail send-draft <message-id>
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
graph-cli chat get <chat-id>
graph-cli chat create --members <emails> [--topic <text>] [--type oneOnOne|group]
graph-cli chat members <chat-id>
graph-cli chat messages <chat-id> [--top <n>]
graph-cli chat send <chat-id> --message <text> [--content-type text|html] [--mentions <id:email,...>]
graph-cli chat reply <chat-id> <message-id> --message <text> [--mentions <id:email,...>]
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

### User Directory

```bash
graph-cli user me
graph-cli user get <user-id-or-email>
graph-cli user search --query <text>
graph-cli user manager
graph-cli user reports
```

### Contacts Allow-List

Outbound actions (mail send, chat send, etc.) are gated by an allowed contacts list.

```bash
graph-cli contacts allow <email-or-group> --actions email,chat [--name "Display Name"] [--type user|group]
graph-cli contacts list [--type user|group]
graph-cli contacts remove <email-or-group>
```

## Global Options

| Option | Description |
|---|---|
| `--format json\|table` | Output format (default: `json`) |
| `--timezone <tz>` | Timezone for datetime I/O — accepts IANA (e.g. `Asia/Karachi`) or Windows IDs (e.g. `Pakistan Standard Time`). Defaults to system local timezone. |

## License

MIT
