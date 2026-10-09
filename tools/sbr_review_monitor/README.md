# SBR Group Review Monitor

A scheduled bot that watches the internet for new reviews of **SBR Group**
(Bengaluru real-estate developer, sbrgroup.in) and its projects, and emails
Nityn (nitynb@gmail.com) whenever a new review appears.

## How it runs

The bot is **not** a daemon in this codebase — it is a Claude Code scheduled
routine bound to the session that created it. Once a month (the 9th, ~02:42
UTC / 8:12 AM IST; daily from launch until 9 Oct 2026, when Nityn asked to
switch to monthly) the routine wakes the session, which:

1. Reads its baseline state (the set of already-seen review pages and
   per-source rating counts). The authoritative copy of that state lives in
   the routine's own prompt; `state.json` in this directory is a mirror
   committed back to this branch when repo access allows.
2. Runs the fixed search-query battery defined in `RUNBOOK.md` via web search.
3. Diffs the results against the baseline:
   - a review page / forum thread URL not in the baseline → **new**
   - a changed rating/review count on a tracked source (e.g. Justdial) → **new**
   - a search snippet showing a review title/date not seen before → **new**
4. If anything new is found, sends a summary email to nitynb@gmail.com
   (source, project, sentiment, link). If nothing is new, it updates the
   baseline silently.

## Why search-based

This environment's network egress proxy blocks direct fetches of the main
review portals (MouthShut, NoBroker, Square Yards, Justdial return
`EGRESS_BLOCKED`), so the bot detects new reviews from web-search results
(titles, URLs, snippets) rather than by scraping pages. That means detection
granularity is per review page / thread / snippet, not per individual on-page
review — a new review on an already-tracked page is caught when it surfaces in
search snippets or changes a tracked count.

## Watched projects

SBR One Residence (Whitefield; pre-launch codename "Windy Ridge"), SBR Horizon
(Whitefield), SBR Magnus (Katamnallur), SBR Florenso (Whitefield), SBR Minara
(Seegehalli/KR Puram), SBR Keerthi (Kattanallur/Sannatammanahalli), SBR Global
Queens Ville (Kumbalgodu), SBR Earth and Sky (Whitefield), SBR Tejas
(Aavalahalli), SBR Gokulam (Kannamangala/Whitefield), SBR Nest (Kannamangala),
SBR Pravanika (Budigere Cross), SBR Beverly Hills (Electronic City), plus the
builder itself.

To add/remove projects, edit the query battery in `RUNBOOK.md` and in the
routine prompt.

## Changing cadence or stopping

The schedule lives in the Claude Code routine (not in this repo). Ask the
session to list routines (`list_triggers`) and update or delete the
`SBR review monitor` routine to change frequency or stop the bot.
