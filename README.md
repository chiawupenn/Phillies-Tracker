# Phillies Tracker

Calendar pages that show which Phillies season-ticket games are still available.
Each page reads a Google Sheet through a small Apps Script web app.

| Season | Page | Data |
|---|---|---|
| 2026 | `index.html` | Phillies Season Ticket 2026 (Google Sheet) |
| 2027 | `2027/index.html` | Phillies Season Ticket 2027 (Google Sheet), feed script in `2027/Code.gs` |

Green dates are available, gray dates are sold. Visitors add games to a cart and
download the list as an image to send to the seller.

## 2027 setup

The 2027 page works before the feed is connected: it shows all 81 home games from
the MLB schedule as available, with times TBD. To show live availability:

1. Open the **Phillies Season Ticket 2027** sheet in Google Drive (Phillies folder).
2. **Extensions > Apps Script**, replace the contents of `Code.gs` with
   [`2027/Code.gs`](2027/Code.gs), and save.
3. Run `setupCheckboxes` once from the editor to turn the Ticket Forwarded,
   Ticket Platform Sale, Self Sell and Attended Game columns into checkboxes
   (or select those cells and use **Insert > Checkbox**). Existing ticks are kept.
4. **Deploy > New deployment > Web app**, with *Execute as: Me* and
   *Who has access: Anyone*. Approve the permissions prompt.
5. Copy the web app URL (it ends in `/exec`) into `APPS_SCRIPT_URL` near the top of
   the script in `2027/index.html`, then commit.

A game shows as sold once **Ticket Forwarded** is checked, the same column the
sheet counts as Total Sold Games. When MLB posts game times, type them into
**Game Time** and the page picks them up on its next refresh (every 5 minutes).

The feed publishes only the date, time, opponent, special event and sold status.
Buyer names and prices are not sent to the page.
