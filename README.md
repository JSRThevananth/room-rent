# Roommate Rent & Expense Splitter

A small Google Apps Script web app for 4 (or more) roommates to log their
monthly essentials spending and see what each person owes toward rent.

## How the rent math works

Each month, every roommate's rent = `base rent + (average spend this period - their own spend this period)`.
People who spent more on groceries/essentials pay less rent; people who spent
less pay more. The rents always sum to `roommate count x base rent`, so the
household's total cost is unchanged — this just redistributes it fairly.

**Timing:** expenses logged during a calendar month (e.g. September) are used
to calculate the *following* month's rent (October), and that figure is shown
on the home page starting the 20th of that following month (configurable).
Before the 20th, the app just shows the expense-entry form.

**Locking:** the first time a given month's rent is calculated, the numbers
are written to the `RentHistory` sheet and locked in. If someone edits an old
expense row afterwards, the already-shown rent figure won't silently change —
you'd need to clear that row from `RentHistory` and reload to force a
recalculation.

## One-time setup

You can start this from either a new standalone Apps Script project
(script.google.com -> New project) or from inside a Google Sheet
(Extensions -> Apps Script) — the script works either way and sets up its
own spreadsheet automatically the first time you run it.

1. Open or create your Apps Script project.
2. Delete the placeholder `Code.gs` content and paste in the `Code.gs` from
   this folder.
3. In the Apps Script editor, click the **+** next to "Files" -> **HTML**,
   name it `Index`, and paste in `Index.html`. Repeat for `Stylesheet` and
   `JavaScript`.
4. Click the gear icon (Project Settings) and check **"Show appsscript.json
   manifest file in editor"**. Open that file and replace its contents with
   `appsscript.json` from this folder. If you're not in the Toronto timezone,
   change `timeZone` to yours (e.g. `"Asia/Colombo"`).
5. In the function dropdown at the top, select `setupSpreadsheet`, click
   **Run**, and approve the permissions prompt.
   - If you opened this from inside a Sheet, it sets up that same Sheet.
   - If you started a standalone project, it creates a brand-new Sheet
     called "Roommate Rent Tracker" in your Drive and remembers it (it won't
     create a second one if you run `setupSpreadsheet` again).
   Either way this creates the `Roommates`, `Settings` and `RentHistory`
   tabs, seeded with Roy, Bavan, Kana and Prashan, a $530 base rent, and PIN
   `4655`.
6. Open that Sheet and check the `Settings` tab — adjust the rent amount, the
   "Rent Display Day" (default 20), or the `PIN` row if needed. The PIN is
   whatever's in that cell at the time someone unlocks the app — change it
   any time, no redeploy needed.
7. Back in Apps Script: **Deploy -> New deployment -> Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
8. Click **Deploy**, copy the web app URL, and share it with your roommates.

## PIN gate

The app shows nothing — not even the roommate names — until the right PIN is
entered. It's a single shared PIN for everyone (not per-person accounts), so
treat it like a door code: good enough to keep randoms who find the link from
poking around, not real per-user security. Whoever unlocks it stays unlocked
in that browser tab (via sessionStorage) until the tab is closed.

## Day-to-day use

- Open the link, enter the PIN, tap your name, add each item + amount, tap **+ Add item**
  for more, then **Submit expenses**. You can submit multiple times through
  the month — it all adds up.
- To add or remove a roommate, just edit the `Roommates` tab in the Sheet
  directly (one name per row) — the app always reads the current list.
- Every month's raw entries land in their own Sheet tab (e.g. `2026-09`), so
  you can always open the Sheet to review or correct anything by hand.

## A heads-up on the example in the original request

The worked example used to spec this out (`245.56 + 145.67 + 156.78 +
345.67 = 960.01`) doesn't actually add up — the real total is **893.68**,
giving an average of 223.42 (not 240.00) and a different rent for everyone.
The app always computes from whatever is actually entered, so this doesn't
affect it — just flagging it so the numbers you see from a real test don't
look "wrong" against that example.
