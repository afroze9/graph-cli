# graph-cli

A .NET global tool for interacting with Microsoft Graph — manage emails, calendar events, Teams chats, To Do tasks, presence, and user directory from the command line. Output is JSON by default (`--format table` for human-readable output).

## Installation

Requires [.NET 10 SDK](https://dotnet.microsoft.com/download) or later.

```bash
# Clone and install
git clone https://github.com/afroze9/graph-cli.git
cd graph-cli
dotnet pack -o ./nupkg
dotnet tool install -g graph-cli --add-source ./nupkg
```

## Authentication

Uses device-code flow with Microsoft Identity. Tokens are cached at `~/.graph-cli/token-cache.bin` and auto-refresh silently.

```bash
graph-cli auth login    # Interactive browser auth (only needed once)
graph-cli auth status   # Check if authenticated
graph-cli auth logout   # Clear cached tokens
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
graph-cli mail mark-read <message-id> [--unread]
graph-cli mail move <message-id> --folder <folder-id-or-name>
graph-cli mail delete <message-id>
graph-cli mail folders
```

### Calendar

```bash
graph-cli calendar list
graph-cli calendar events [--start <iso-date>] [--end <iso-date>] [--calendar-id <id>] [--top <n>]
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
graph-cli chat send <chat-id> --message <text> [--content-type text|html]
graph-cli chat reply <chat-id> <message-id> --message <text>
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
