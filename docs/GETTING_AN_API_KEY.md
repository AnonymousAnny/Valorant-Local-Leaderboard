# Getting a HenrikDev API key and adding it to the sheet

The leaderboard script asks the [HenrikDev API](https://docs.henrikdev.xyz/) for each player's rank, and the
API needs a key. A key is free. Each key is limited to **30 requests per minute**, so a sheet with many players
works best with **two keys** (see step 4).

> **Never paste a key into a cell, a sheet tab, a README or a commit.** A sheet that is shared publicly shows
> every tab to everybody. Keys belong only in the Apps Script *Script properties* described below.

## 1. Request a key

1. Open the dashboard: <https://api.henrikdev.xyz/dashboard/api-keys> and sign in.
2. On the **API Keys** page, request a new key. When asked, choose the **Valorant** key type and give it a
   recognisable name, for example `ValorantMMRCheckGoogleSheet`.
3. The key appears as a card with its name, type badges, a status, the **Access token** and its **rate limits**.
   - The status may read **Pending** at first. If the script gets `401 Unauthorized` or `403` straight away,
     wait for the key to be approved and try again.
   - The token starts with `HDEV-`. It is masked on the card: use the eye icon to reveal it and the copy icon to
     copy it. Reveal it only when you are about to paste it.
   - The card also has **Regenerate** (new token, old one stops working) and **Delete**.
4. Optional but recommended: request a **second key** the same way. The script alternates between all your
   keys, so two keys give about 60 requests per minute and a rate-limited retry switches to the other key.

The dashboard layout may change over time. If a button has moved, the HenrikDev
[documentation](https://docs.henrikdev.xyz/) describes the current process.

## 2. Add the key to your sheet's script

1. Open your copy of the leaderboard sheet and choose **Extensions → Apps Script**.
2. In the left menu, click the gear icon (**Project Settings**).
3. Scroll to **Script properties** and click **Add script property** (or **Edit script properties**).
4. Enter these, one row per key:

   | Property | Value |
   |---|---|
   | `API_KEY1` | your first token, for example `HDEV-xxxxxxxx-...` |
   | `API_KEY2` | your second token (optional) |

   The script accepts any property named `API_KEY`, `API_KEY1`, `API_KEY2`, `API_KEY_3`, and so on, so you can
   add as many keys as you like. You can also put several keys in one property, separated by commas.
5. Click **Save script properties**.

Optional properties: `REGION` (default `ap`, other values include `eu`, `na`, `kr`) and `SHEET_NAME` (default is
the first tab).

## 3. Run it once

1. In the Apps Script editor, paste in [`Code.gs`](../Code.gs) and save.
2. Pick `refreshValorantData` in the function dropdown and click **Run**.
3. Google asks you to **Review permissions**. Allow them: the script needs to call the HenrikDev API and,
   if you use auto-refresh, to create a trigger.
4. Go back to the sheet. If the rank icons show `#REF!`, click **Allow access** on the yellow banner once.
5. Optional: **Valorant → Enable auto-refresh on open** makes the sheet refresh whenever it is opened.

## Troubleshooting

| What you see in the rank cell | What it means |
|---|---|
| `API key not found.` (in the execution log) | No property named `API_KEY...` exists, or it was saved with a typo. |
| `API Error: 401` / `Unauthorized` | The key is wrong, still pending, or was regenerated or deleted. Paste the current token. |
| `API Error: 404 (Account not found)` | The Riot ID name or tag is wrong. Check spelling and the number after `#`. |
| `API Error: 429` (rate limit) | Too many requests per minute. The script already retries; add a second key or refresh less often. A row that fails this way keeps its previous rank. |
| `Missing name or tag` | The row has a tag but no name, or a name but no tag. |
| `#REF!` in the icon column | Click **Allow access** on the yellow banner (Google asks before `IMAGE()` loads external pictures). |

## If a key leaks

If a key was ever shown in a sheet, a screenshot or a commit, open the dashboard, click **Regenerate** (or
**Delete** and request a new one), and update the Script property. The old token stops working immediately.
