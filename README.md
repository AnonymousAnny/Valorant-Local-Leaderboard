# Valorant Local Leaderboard

A Google Sheet that keeps a live Valorant rank leaderboard for a group of friends. A Google Apps Script
(`Code.gs`) reads Riot IDs from the sheet, asks the [HenrikDev API](https://docs.henrikdev.xyz/) for each
player's current rank, and writes the results back, sorted from best to worst.

<a href="https://docs.google.com/spreadsheets/d/1ZEW9ZiodScvkAmRtjblUzVHhWzH39gIGYMmhucP2zYY/edit?usp=sharing">Link to the public Google Sheet</a> (view only, make a copy to use it yourself)

## What it does

For every row in the sheet it fetches the player's rank and fills in:

| Column | Meaning |
|---|---|
| Rank | Current tier, for example `Platinum 2` |
| Rank image | The rank icon (`=IMAGE(...)`) |
| RR | Ranked rating inside the tier |
| ELO | Used for sorting (hidden column) |
| Highest Rank / season | Peak rank and the act it was reached in |
| Games needed | Placement games still needed for a rating (hidden column) |

Then it:

- **Sorts** the leaderboard: players with a current rating first, ordered by ELO (highest first).
- **Greys out inactive accounts.** If an account still needs any games for a rating (`games needed > 0`),
  its rank cells are greyed and the account sits below all active players.
- **Crowns the leader.** Exactly one 👑 is kept on the top row of the Owner column.
- **Reports problems per row**, for example `API Error: 404 (Account not found)` or `Missing name or tag`,
  without stopping the other rows.
- **Refreshes** from the menu (**Valorant → Refresh Valorant Stats**) or automatically when the sheet is
  opened (**Valorant → Enable auto-refresh on open**).
- Keeps a **Discord** sidebar and menu.

## Reliability features

- **Several API keys.** Requests are spread across all keys you configure, and a retry after a rate limit
  switches to the next key.
- **Rate-limit handling.** Requests go out in small batches. A `429` or `5xx` response is retried (up to 4
  times), waiting for the API's own reset time (`x-ratelimit-reset` / `Retry-After`) when it is sent.
- **No data loss on temporary failures.** If a refresh fails for a transient reason (rate limit, server
  error, network), that row keeps its previous rank instead of being wiped. A definitive error such as
  "account not found" does replace it.
- **Safe concurrency.** A script lock stops two refreshes from running at the same time.
- Names and tags are URL-encoded, a leading `#` on the tag is optional, and blank rows are skipped.

## Setup

1. Make a copy of the sheet, then open **Extensions → Apps Script** and paste in [`Code.gs`](Code.gs).
2. Get a free HenrikDev key (step-by-step guide: [docs/GETTING_AN_API_KEY.md](docs/GETTING_AN_API_KEY.md)), then open
   **Project Settings → Script properties** and add:
   - `API_KEY` (required): your HenrikDev API key. To use several keys, add more properties named
     `API_KEY1`, `API_KEY2`, `API_KEY_3`, and so on, or list keys separated by commas in one value.
   - `REGION` (optional): defaults to `ap`.
   - `SHEET_NAME` (optional): defaults to the first tab.
3. Reload the sheet, open **Valorant → Refresh Valorant Stats** and approve the Google permission prompt.
4. Optional: **Valorant → Enable auto-refresh on open** installs a trigger so the sheet refreshes whenever
   it is opened. (A plain `onOpen` trigger cannot make API calls, so this separate trigger is needed.)
5. If the rank icons show `#REF!`, click **Allow access** on the yellow banner once.

Never put API keys or account passwords in the sheet or in this repository. Keys belong in Script properties.

## Sheet layout

Row 1 is a header. Columns A and B hold the Riot ID (name and tag, with or without `#`). The script finds the
column headed **Owner** (the 👑 goes there) and uses the seven columns after it for its output. If no
`Owner` header exists, it assumes Owner is column E, so a sheet with two extra columns before it (as in the
private version of this project) works as well as the public one.

```
Riot ID | Tag | Owner | Rank | (icon) | RR | ELO* | Highest Rank | Season | Games needed*
                                         * hidden by the script
```

## Tests

`node test/run.js` runs `Code.gs` against mocked Apps Script services and a fake API, with no network or key
needed. It checks sorting, greying and bottom placement of inactive accounts, retries and key rotation,
preserving data on transient errors, layout detection, the menu and trigger, and edge cases such as an
empty sheet.

## Contributors

Thanks to @Henrik-3 for the [HenrikDev API](https://docs.henrikdev.xyz/).
