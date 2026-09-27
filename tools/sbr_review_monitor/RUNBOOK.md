# SBR Review Monitor — Per-Run Procedure

Follow these steps exactly on every scheduled run. The authoritative baseline
lives in the routine prompt (updated via `update_trigger` each run); this repo
holds a mirror in `state.json`.

## 1. Sync state

```bash
git fetch origin claude/sbr-review-monitor-bot-x3kvnn
git checkout claude/sbr-review-monitor-bot-x3kvnn
git pull origin claude/sbr-review-monitor-bot-x3kvnn
```

Read the baseline from the routine prompt (primary) and
`tools/sbr_review_monitor/state.json` (mirror).

## 2. Run the query battery (WebSearch)

Run every query. Do not reword them — stable queries make diffs meaningful.

1. `"SBR Group" Bangalore review`
2. `"SBR Group" Bangalore complaint OR "buyer experience"`
3. `"SBR One Residence" OR "SBR Horizon" OR "SBR Magnus" review`
4. `"SBR Florenso" OR "SBR Minara" OR "SBR Keerthi" OR "SBR Queens Ville" OR "SBR Tejas" OR "SBR Gokulam" OR "SBR Windy Ridge" OR "SBR Nest" OR "SBR Pravanika" OR "SBR Beverly Hills" review`
5. `SBR Group Bangalore builder reddit OR forum review`
6. `SBR Group MouthShut OR Justdial new review`

Optionally attempt `WebFetch` on any newly discovered review URL; expect
`EGRESS_BLOCKED` for most portals and fall back to the search snippet.

## 3. Diff against state

A result is review-relevant if its title/snippet concerns a review, rating,
complaint, or buyer experience of SBR Group or a watched project. Ignore pure
listings/brochure/price pages, builder-site testimonial copy, employee-review
pages (Glassdoor), and false positives (e.g. "Shanghai Business Review",
styrene-butadiene rubber, Gokulam Kerala FC).

Flag as **NEW** when:
- a review-relevant URL is not in the baseline (normalize: strip query params,
  trailing slashes, and `www.`);
- a snippet contains a review title, reviewer, or date not previously recorded
  for a URL already in the baseline;
- a tracked metric visibly changed (e.g. Justdial "204 Ratings" becomes a
  higher number in the search snippet).

Paraphrased positives with no reviewer/date/URL that read as variants of
already-recorded praise fold into the baseline silently.

## 4. Notify (only if something is NEW)

Send an email via the Gmail tool:
- **To**: nitynb@gmail.com
- **Subject**: `SBR review monitor: N new review(s) — <date>`
- **Body**: for each new item — project, source site, link, review
  date if known, and a 1–2 line gist with sentiment (positive/negative/mixed).
  End with a one-line reminder that asking Claude to delete the
  `SBR review monitor` routine stops the bot.

If nothing is new, send nothing. Never email "no news".

## 5. Persist state

Primary: rewrite the routine prompt via `update_trigger` with new items folded
into the baseline, counts updated, and a run-log line appended.

Mirror: update `state.json` (seen URLs, `tracked_counts`, `last_run`,
`run_log` — keep the latest 30 entries), then:

```bash
git add tools/sbr_review_monitor/state.json
git commit -m "sbr-review-monitor: run <date>, <N> new"
git push -u origin claude/sbr-review-monitor-bot-x3kvnn
```

Retry the push up to 4 times with exponential backoff (2s/4s/8s/16s) on
network failure. If the push fails with a permissions error (403), skip the
mirror update for the day — the routine prompt remains authoritative — and
note it in the run summary.
