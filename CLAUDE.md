# MyOffice CLI

## Overview

A command-line interface for personal Microsoft 365 access via Microsoft Graph API. Uses delegated authentication - users authenticate as themselves and can only access their own data.

**Project Management:**
- Planner Plan ID: `_gFp0xaWK06WnP5z5BEPmZgAE7LQ` (MyOffice CLI)
- On startup, check the plan for current tasks and priorities using: `npx @awolve/myoffice@latest planner tasks --plan-id "_gFp0xaWK06WnP5z5BEPmZgAE7LQ" --json`

**IMPORTANT:** Before using the `myoffice` CLI, ALWAYS check if relevant Awolve skills are available first:
- `awolve-general:myoffice-login` - For authentication issues
- `awolve-general:myoffice-setup` - For initial setup
- `awolve-general:awolve-myoffice` - For general MyOffice operations

These skills handle authentication, configuration, and common issues more gracefully than direct CLI calls.

## Architecture

```
src/
├── cli.ts             # CLI entry point (Commander.js)
├── core/
│   └── handler.ts     # Shared command dispatch logic
├── cli/
│   └── formatter.ts   # Human-readable output formatting
├── auth/
│   ├── config.ts      # Azure AD client config and scopes
│   ├── device-code.ts # Device code authentication flow
│   ├── token-manager.ts # Token caching and refresh
│   ├── login.ts       # CLI login script
│   └── index.ts       # Auth exports
├── tools/             # Graph API tool implementations
│   ├── mail.ts        # Email operations
│   ├── calendar.ts    # Calendar events
│   ├── tasks.ts       # Microsoft To Do
│   ├── planner.ts     # Microsoft Planner (plans, buckets, tasks)
│   ├── onedrive.ts    # OneDrive files + shared files
│   ├── sharepoint.ts  # SharePoint sites and document libraries
│   ├── contacts.ts    # Contacts
│   ├── teams.ts       # Teams channels and messages
│   ├── chats.ts       # 1:1 and group chats
│   └── index.ts       # Tool exports
└── utils/
    ├── graph-client.ts # Authenticated Graph API client
    └── version.ts      # Dynamic version from package.json
```

## Key Files

- `src/cli.ts` - CLI entry point with all commands
- `src/core/handler.ts` - Command dispatch logic
- `src/cli/formatter.ts` - Human-readable output formatting
- `src/auth/config.ts` - Azure AD scopes (add new permissions here)
- `src/utils/graph-client.ts` - All Graph API calls go through this
- `src/tools/*.ts` - Individual tool implementations

## Authentication

- Device code flow for initial authentication (`npm run login`)
- Tokens cached at `~/.config/myoffice-mcp/token.json`
- Requires Azure AD app registration with delegated permissions
- Environment variables: `M365_CLIENT_ID` (required), `M365_TENANT_ID` (optional, defaults to "common")

## Adding New Commands

1. Create or update tool module in `src/tools/`
2. Define Zod schema for input validation
3. Implement function that calls Graph API via `graphClient.fetch()`
4. Add command to `src/cli.ts`
5. Add case to switch statement in `src/core/handler.ts`
6. If new permissions needed, add scopes to `src/auth/config.ts`

## Development

```bash
npm run dev    # Run with tsx (TypeScript directly)
npm run build  # Compile to dist/
npm run login  # Authenticate with Microsoft
```

## CLI Usage

The `myoffice` CLI provides terminal access to all Microsoft 365 tools.

### Installation & Authentication

```bash
npm install -g awolve-myoffice-cli

# First-time setup: provide your Azure AD client ID (saved for future use)
myoffice login --client-id <your-azure-app-client-id>

# Subsequent logins (client ID remembered)
myoffice login
```

**Configuration priority:**
1. Environment variable `M365_CLIENT_ID` (if set)
2. Stored config file `~/.config/myoffice-mcp/config.json`

You can also set config without logging in:
```bash
myoffice config set --client-id <id> --tenant-id <id>
myoffice config show
```

### Commands

| Command | Description |
|---------|-------------|
| `myoffice login [--client-id <id>]` | Authenticate with Microsoft 365 |
| `myoffice status` | Check authentication status |
| `myoffice debug` | Show server and auth info |
| `myoffice config show` | Show current configuration |
| `myoffice config set --client-id <id>` | Save client ID to config file |

**Mail:**
- `myoffice mail list [--folder <name>] [--limit <n>] [--skip <n>] [--unread] [--mailbox <address>]` - List emails (supports custom folder names; `--skip` pages through a folder; `--mailbox` reads a shared mailbox you have Full Access to, e.g. `test@awolve.ai`). Rows include `isDraft`, `to`/`toNames`, and `lastModified` — a draft has no sender, so `to` is what identifies it
- `myoffice mail read <id> [--mailbox <address>]` - Read email
- `myoffice mail search <query> [--mailbox <address>]` - Search emails (KQL, e.g. `to:test+abc@awolve.ai` matches only that plus-address; results span all folders of the mailbox)
- `myoffice mail send --to <addr> --subject <subj> --body <body> [--attach <files...>]` - Send email with optional attachments
- `myoffice mail draft --to <addr> --subject <subj> --body <body> [--attach <files...>]` - Create draft email with optional attachments
- `myoffice mail draft-update --id <id> [--to <addr...>] [--cc <addr...>] [--subject <subj>] [--body <body>]` - Update an existing draft in place; only the options given are replaced, so omitting `--body` keeps the draft's original HTML byte for byte
- `myoffice mail draft-send --id <id>` - Send an existing draft exactly as it stands (no body rewriting). Both refuse anything that is not an unsent draft
- `myoffice mail reply <id> --body <body> [--all]` - Reply to email
- `myoffice mail forward --id <id> --to <addr...> [--body <comment>]` - Forward email, optionally with a comment above the forwarded message
- `myoffice mail delete <id>` - Delete email
- `myoffice mail mark <id> [--unread]` - Mark as read/unread
- `myoffice mail move --id <id> --folder <name>` - Move email to folder (creates folder if needed)
- `myoffice mail attachments --id <id>` - List email attachments
- `myoffice mail download-attachment --id <id> --attachment-id <attachId> --output <path>` - Download attachment

**Calendar:**
- `myoffice calendar list [--start <date>] [--end <date>]` - List events
- `myoffice calendar get <id>` - Get event details
- `myoffice calendar create --subject <subj> --start <dt> --end <dt>` - Create event
- `myoffice calendar update <id> [--subject <s>] [--start <dt>]` - Update event
- `myoffice calendar delete <id>` - Delete event

**Tasks (Microsoft To Do):**
- `myoffice tasks lists` - List task lists
- `myoffice tasks list [--list <id>] [--completed]` - List tasks
- `myoffice tasks create <title> [--list <id>] [--due <date>]` - Create task
- `myoffice tasks update <id> [--title <t>] [--due <date>]` - Update task
- `myoffice tasks complete <id>` - Mark task complete
- `myoffice tasks delete <id>` - Delete task

**Files (OneDrive):**
- `myoffice files list [path]` - List files
- `myoffice files get <path>` - Get file metadata
- `myoffice files search <query>` - Search files
- `myoffice files read <path>` - Read text file content
- `myoffice files mkdir <name> [--parent <path>]` - Create folder
- `myoffice files shared` - List files shared with me
- `myoffice files upload --file <path> [--dest <path>]` - Upload local file (any size)
- `myoffice files download --path <path> --output <path>` - Download file to local path

**SharePoint:**
- `myoffice sharepoint sites [--search <query>]` - List sites
- `myoffice sharepoint site <id>` - Get site details
- `myoffice sharepoint drives <siteId>` - List document libraries
- `myoffice sharepoint files <driveId> [path]` - List files
- `myoffice sharepoint file <driveId> <path>` - Get file metadata
- `myoffice sharepoint read <driveId> <path>` - Read file content
- `myoffice sharepoint search <driveId> <query>` - Search files
- `myoffice sharepoint download --url <url> --output <path>` - Download file from SharePoint URL
- `myoffice sharepoint download-file --drive-id <id> --path <path> --output <path>` - Download file by drive ID
- `myoffice sharepoint upload --drive-id <id> --file <path> --dest <path>` - Upload file to SharePoint

**Contacts:**
- `myoffice contacts list` - List contacts
- `myoffice contacts search <query>` - Search contacts
- `myoffice contacts get <id>` - Get contact details
- `myoffice contacts create [--given-name <n>] [--surname <n>] [--email <e>] [--birthday <date>]` - Create contact
- `myoffice contacts update --id <id> [--given-name <n>] [--email <e>] [--birthday <date>]` - Update contact

**Teams:**
- `myoffice teams list` - List teams
- `myoffice teams channels <teamId>` - List channels
- `myoffice teams messages <teamId> <channelId>` - List channel messages
- `myoffice teams post <teamId> <channelId> <message>` - Post message

**Chats:**
- `myoffice chats list` - List chats
- `myoffice chats messages <chatId>` - List chat messages
- `myoffice chats send <chatId> <message>` - Send message
- `myoffice chats create <email> [message]` - Create/start chat

**Planner:**
- `myoffice planner plans [--group <id>]` - List plans
- `myoffice planner plan <id>` - Get plan details
- `myoffice planner buckets <planId>` - List buckets
- `myoffice planner bucket-create <planId> <name>` - Create bucket
- `myoffice planner bucket-update <id> <name>` - Update bucket
- `myoffice planner bucket-delete <id>` - Delete bucket
- `myoffice planner my-tasks [--status <s>] [--limit <n>]` - List all tasks assigned to me across all plans
- `myoffice planner tasks <planId> [--bucket <id>]` - List tasks
- `myoffice planner task <id>` - Get task details
- `myoffice planner task-create <planId> <title> [--bucket <id>]` - Create task
- `myoffice planner task-update <id> [--title <t>] [--progress <p>]` - Update task
- `myoffice planner task-delete <id>` - Delete task
- `myoffice planner task-details <id>` - Get task details (description, checklist, attachments)
- `myoffice planner task-details-update <id> [--description <d>]` - Update details
- `myoffice planner attach --id <taskId> --url <url> [--alias <name>]` - Add link/attachment to task
- `myoffice planner detach --id <taskId> --url <url>` - Remove attachment from task
- `myoffice planner upload --id <taskId> --file <path> [--alias <name>]` - Upload file and attach to task
- `myoffice planner checklist-add --id <taskId> --title <title> [--checked]` - Add checklist item
- `myoffice planner checklist-remove --id <taskId> --item <itemId>` - Remove checklist item
- `myoffice planner checklist-toggle --id <taskId> --item <itemId>` - Toggle checklist item

### Output Format

By default, CLI outputs human-readable tables. Use `--json` flag for JSON output:

```bash
myoffice mail list --json
myoffice --json calendar list
```

### Shared Mailbox Examples

`--mailbox <address>` on `mail list`, `mail read`, and `mail search` reads a shared mailbox the signed-in user has Full Access to (Graph `/users/<address>/...` instead of `/me/...`; scope `Mail.ReadWrite.Shared`). Read-only: send, reply, move, and delete stay on your own mailbox.

```bash
# Newest mail in the shared test mailbox
myoffice mail list --mailbox test@awolve.ai --limit 10 --json

# Only mail sent to one plus-address (KQL to: matches the exact address)
myoffice mail search --mailbox test@awolve.ai --query "to:test+exec42-candidate@awolve.ai" --json

# Read one of them
myoffice mail read --id <message-id> --mailbox test@awolve.ai --json
```

No access, or the mailbox does not exist: Graph returns 403/404 and the CLI never falls back to your own mailbox. Only real shared mailboxes work: an M365 Group address (e.g. `hello@awolve.ai`) fails with `403 Group Shard is used in non-Groups URI`, because group mail lives under `/groups/<id>/conversations`, not `/users`. A freshly created shared mailbox answers `404 Default folder Inbox not found` for the first ~15 minutes.

### Email Folder Examples

List emails from custom folders:

```bash
# List from inbox (built-in folder)
myoffice mail list --folder inbox

# List from custom folders
myoffice mail list --folder "Under Processing"
myoffice mail list --folder "Captured"
myoffice mail list --folder "Skipped"

# List unread emails from custom folder
myoffice mail list --folder "Under Processing" --unread --json
```

### Email Attachment Examples

Send emails with attachments:

```bash
# Send email with single attachment
myoffice mail send --to user@example.com --subject "Report" --body "See attached" --attach report.pdf

# Send email with multiple attachments
myoffice mail send --to user@example.com --subject "Documents" --body "Here are the files" --attach file1.pdf file2.xlsx

# Create draft with attachments
myoffice mail draft --to user@example.com --subject "Draft with files" --body "Review these" --attach document.docx

# Download attachments from received email
myoffice mail attachments --id <message-id> --json
myoffice mail download-attachment --id <message-id> --attachment-id <attach-id> --output ~/Downloads/file.pdf
```

## Testing

`npm test` runs the unit tests in `tests/` (Node's built-in `node:test` via tsx, no Graph access needed). `tests/tasks.test.ts` is an older integration script that hits Graph; run it by hand with `npx tsx tests/tasks.test.ts`. Everything else: test manually by running CLI commands.

Use `myoffice debug` to check version and auth state.

## Versioning

- Version is tracked in `package.json`
- Run `npm run build` after bumping the version
- Publish with `npm publish`

## Specs

Feature specs live in the awolve-cortex context repo, not in this repo:
`~/code/awolve-cortex/awolve-context/operations/tools/myoffice/specs/`

Do NOT create specs in this repo.

## Common Tasks

### On session startup
1. Check the Planner plan for current tasks and priorities:
   ```bash
   npx @awolve/myoffice@latest planner tasks --plan-id "_gFp0xaWK06WnP5z5BEPmZgAE7LQ" --json
   ```
2. Review open tasks to understand current priorities

### Before pushing new features
1. Bump the version in `package.json`
2. Check `~/code/awolve-cortex/awolve-context/operations/tools/myoffice/specs/` for a spec that describes the current changes
3. If no relevant spec exists, run `/retro-spec` to document the work (spec goes in awolve-cortex)
4. Commit and push

### Add new Graph API permission
Add scope to `SCOPES` in `src/auth/config.ts`. Every user must re-authenticate with `myoffice login`: the silent token refresh requests the full scope set, so after a scope is added *every* command fails with `AADSTS65001` until the user logs in again — not only the commands that use the new scope (seen when `Mail.ReadWrite.Shared` was added in v2.13.0).

### Debug issues
Run `myoffice debug` to see version, environment variables, and auth status.

## Planner Integration

Microsoft Planner provides team-oriented task management. Key concepts:

- **Plans** - Read-only. Plans belong to M365 Groups, users can only see plans in groups they're members of.
- **Buckets** - Columns within a plan. Full CRUD supported.
- **Tasks** - Items within buckets. Full CRUD with assignments, due dates, priority, progress.
- **Task Details** - Extended info: description, checklist items, references (attachments).
- **References/Attachments** - Links to files or URLs. Planner doesn't store files directly; it stores references (URLs) to files in OneDrive, SharePoint, or external sites.

**Important notes:**
- All updates/deletes require ETags (handled internally - no user action needed)
- Task assignments accept email addresses (resolved to user IDs automatically)
- **Read assignees from `assignees`, not `assignedTo`** (v2.12.0+). `tasks`, `task`
  and `my-tasks` all return `assignees: [{userId, displayName, email}]`, resolved
  via `/users/{id}` and cached per session. `assignedTo` is kept for backward
  compatibility and is inconsistent by history: bare id strings from `tasks`/`task`,
  objects-or-strings from `my-tasks`. It was not changed in place because every
  consumer runs `npx @awolve/myoffice@latest`, so a shape change lands on every
  machine before the code that reads it does.
- A `assignees` entry always has `userId`; `displayName`/`email` are absent when the
  user cannot be looked up (deleted or external). That distinguishes "unknown who"
  from "unassigned" — do not render the raw `userId` as if it were a name.
- Progress values: `notStarted`, `inProgress`, `completed`
- Priority values: `urgent`, `important`, `medium`, `low`
- Plans cannot be created via Graph API (would require M365 Group creation)
- Attachments are stored as "references" (URLs). The `planner upload` command uploads to the plan's SharePoint site (accessible to all plan members) at `Planner Attachments/<Plan Name>/<filename>`. File type is auto-detected from the URL.
