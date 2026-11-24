# Brillienta Copilot

One Microsoft 365 Copilot agent for every Brillienta app a customer uses. Employees ask in Copilot or Teams about
their helpdesk requests, the equipment assigned to them and the people they work with, in one conversation.

The agent calls Microsoft Graph **as each signed-in person**, so SharePoint permissions still decide what everyone
sees. There is no Brillienta server or database: the setup page builds the agent package in the admin's browser, and
the agent talks straight to the customer's SharePoint lists.

## What it can do (phase 1)

| App | Read | Write | How writes are applied |
| --- | --- | --- | --- |
| Helpdesk | Tickets, conversation, knowledge base, categories, statuses, priorities, agents | Raise a ticket, reply, 17 ticket actions (resolve, reopen, assign, approve...) | Queued in `HD_CopilotActions` and applied by Helpdesk's own `CopilotProcessor` with its normal rules |
| Asset Management | My equipment, who has a tag, equipment requests, asset types | Request or borrow equipment, report a fault, return, cancel a request, confirm receipt | Requests queued in `AM_CopilotActions` and applied by Asset Management's `CopilotProcessor` as the person who asked; receipts via `AM_Acknowledgements`. Needs Asset Management with Copilot requests (schema 2) and its feature switch on; read-only otherwise |
| Employee Directory | People by department, office, country, manager, direct reports; departments | None | Read-only by design |

Every write asks the person to confirm first. Directory reads request only directory-card fields (never mobile
phone, hire date or employee number) and skip people hidden from the directory.

## How it works

```
Setup page (SPFx web part, this repo)
  1. Finds Brillienta apps on the sites you add (lists matched by name, any prefix)
  2. Reads Helpdesk's settings (ticket prefix, "Process Copilot requests")
  3. Builds the agent package in the browser:
       manifest.json, declarativeAgent.json,
       ai-plugin-core.json + openapi-core.json          (who is asking: /me)
       ai-plugin-<app>.json + openapi-<app>.json         (one per app found)
  4. Admin downloads the zip and uploads it in Microsoft 365 admin center
```

Settings (agent name, stable app id, OAuth registration id, sites) are kept in the list `BrillCP_Settings` on the
site that hosts the setup page, following the `Brill<AppShort>_` naming convention.

| Path | What it is |
| --- | --- |
| `src/webparts/brillientaCopilot/core/` | Pure, unit-tested agent builder: `compose.ts` (merges apps), `apps/*.ts` (one module per app), `discovery.ts`, `openapi.ts`, `package.ts` |
| `src/webparts/brillientaCopilot/services/` | SharePoint REST: `SiteProbe.ts` (read-only discovery), `SettingsStore.ts` (`BrillCP_Settings`) |
| `src/webparts/brillientaCopilot/components/SetupPage.tsx` | The four-step setup page |
| `tests/` | `node:test` against `lib-commonjs` |

## Build and test

Requires Node 22 (SPFx 1.23), the same as the other Brillienta apps.

```bash
npm install
npm test          # heft build + 20 tests
npm run build     # production bundle + sharepoint/solution/Brillienta-Copilot.sppkg
```

The tests:

- validate every generated file against Microsoft's published schemas (Teams manifest v1.19, declarative agent v1.3,
  API plugin v2.2, shipped in `@microsoft/app-manifest`);
- check every plugin function has its OpenAPI operation, no action sends headers, every write has a confirmation,
  and the instructions fit the 8,000-character limit for every combination of apps;
- **drift guards**: when the product repos are checked out next to this one (`../brillientahelpdeskwebpart`,
  `../brillienta-assetmanagement`, `../brillientaemployeedirectory`), they check that every column the agent filters
  on is still indexed, every column it requests still exists, and the Helpdesk action list matches Helpdesk's
  `services/copilot/actions.ts`. A product change that would break the agent fails this build.

## Setup for customers

See [docs/setup.md](docs/setup.md).

## Phase 2 (each a reviewed branch in the product repo)

- [x] Asset Management: `AM_CopilotActions` queue + processor (request, return, report a fault, cancel) and receipt confirmation (branch `claude/copilot-actions` in Brillienta-AssetManagement).
- [ ] HRMS: index `EmployeeId` on Leave and Expenses, then leave balance, apply for leave and approvals via a queue.
- [ ] TaskManager: index `AssigneeEmail`, then "my tasks", create and update tasks via a queue.
- [ ] Fill in `core/publisher.ts` after Brillienta's one-time Entra + OAuth registration, so customers only consent.

## Not verified here

Built and tested without a Microsoft 365 tenant. Before release, upload the package to a dev tenant and try each
conversation starter, a requester and an agent account in Helpdesk, and a user without access to the Assets list.
