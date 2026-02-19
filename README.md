# Planner MCP Server

A small Node.js MCP server that integrates with Microsoft Planner using the Microsoft Graph API.

## Features

| Tool | Description |
|---|---|
| `list_groups` | List Microsoft 365 groups (teams) that have Planner plans. Use this to discover the `groupId` needed for `list_plans`. |
| `list_plans` | List all Planner plans for a Microsoft 365 group (team). |
| `list_buckets` | List all buckets in a Planner plan. |
| `list_tasks` | List all tasks in a Planner plan. |
| `get_task` | Get details of a specific Planner task by its ID (includes task details/description). |
| `create_task` | Create a new Planner task with title, dates, priority, assignments, and categories. |
| `update_task` | Update an existing Planner task (title, dates, progress, priority, bucket, assignments) and its details. |
| `complete_task` | Mark a Planner task as complete (sets `percentComplete` to 100). |
| `list_categories` | List the category labels (colored tags) defined in a Planner plan. |
| `delete_task` | Delete a Planner task permanently. |
| `list_users` | List users in the Microsoft 365 directory (useful to discover user IDs for assignments). |

## Prerequisites

1. An **Azure App Registration** with the following Application permissions in Microsoft Graph:
   - `Group.Read.All`
   - `Tasks.ReadWrite.All`
   - `User.Read.All`
2. A **Client Secret** generated for the App Registration.
3. Node.js 18 or newer.

## Setup

1. Install dependencies:

```bash
npm install
```

2. Copy and edit the environment file:

```bash
cp .env.example .env
```

Fill these environment variables in `.env`:

```env
CLIENT_ID=<App Registration Client ID>
TENANT_ID=<Azure AD Tenant ID>
CLIENT_SECRET=<Client Secret>
# Optional defaults used by some tools
GROUP_ID=<optional default group id>
PLAN_ID=<optional default plan id>
```

3. Build the project:

```bash
npm run build
```

## Usage

Run the server as an MCP stdio server. Example `mcp.json` entry:

```json
{
  "servers": {
    "planner": {
      "command": "node",
      "args": ["/absolute/path/to/planner-mcp-server/dist/index.js"],
      "env": {
        "CLIENT_ID": "...",
        "TENANT_ID": "...",
        "CLIENT_SECRET": "...",
        "GROUP_ID": "...", // optional default group id for tools that require it
        "PLAN_ID": "..." // optional default plan id for tools that require it
      }
    }
  }
}
```

For local development you can run:

```bash
npm run dev
```

## Testing

Tests are written with [Vitest](https://vitest.dev/) and live in `src/__tests__/`.

| File | Coverage |
|---|---|
| `utils.test.ts` | `ok` / `err` response helpers |
| `planner-service.test.ts` | `PlannerService` — all CRUD operations against a mocked Graph client |

Run the test suite:

```bash
# Run once
npm test

# Watch mode (re-runs on file changes)
npm run test:watch

# With coverage report
npm run test:coverage
```

No real Azure credentials are needed to run the tests — the Microsoft Graph client is fully mocked.

## Project Structure

```
src/
  graph-client.ts    # Creates an authenticated Microsoft Graph client
  planner-service.ts # Service implementing Planner CRUD operations
  utils.ts           # Response helpers (ok / err)
  index.ts           # Entry point — registers tools and starts the MCP server
  __tests__/
    utils.test.ts            # Unit tests for response helpers
    planner-service.test.ts  # Unit tests for PlannerService
```

## Example Prompts

- "List the plans in group abc123"
- "Create a task 'Review document' in plan xyz456 with urgent priority and due 2025-04-01"
- "Mark task task789 as completed"
- "Update task task789 description to 'Document reviewed and approved'"