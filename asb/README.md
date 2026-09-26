# ASB Athlete App

One app for Akanksha Singh Basketball athletes, parents and staff. It replaces the daily
WhatsApp roster message. It is built module by module. Module 1 is the schedule and
roster.

## Status

| Phase 1 step | State |
|---|---|
| 1. Read-only importer: Roster 2026 → events + review report | **Built**: [`roster_importer/`](roster_importer/README.md). Waiting for Nityn's sign-off on the Jun–Sep 2026 report |
| 2. Membership table (athlete → squad / class / PT group / team) | Seed comes out of step 1 (`memberships.csv`). Still to add: attendance register and development squads |
| 3. Calendars (central + personal, ICS feeds) | Not started |
| 4. Daily digest | Not started (the 9 am roster email already runs separately) |
| 5. Source of truth decision | The sheet stays the source. The app re-imports it |

## Privacy

Many athletes are minors. **This repository is public**, so it holds code only. Roster
exports, import output and `aliases.csv` (which holds athletes' names) are in
`.gitignore`. Build the app itself in a **private** repository.
