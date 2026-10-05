# Sallee Local Dispatch Board

The local-fleet twin of the OTR dispatch board: drivers as columns, dates as rows, two weeks at a time, jobs in the grid.
Single self-contained `index.html` (no build step), hosted on GitHub Pages, data in SharePoint, Microsoft 365 sign-in.

## What's different from the OTR board
- **Drivers** are grouped Local / Local-OTR / Agent / Part-time (not KY / FL / NY).
- **Jobs** have a **Customer** line, **In** (report time) and **Out** (clock-out time, or **WD** = when done), free-text origin/destination (autocomplete from past jobs) and local job types: Drop, Race & Return, Swap, Breeding trip, Night breeding trip, Shuttle, Wait & Return, Other.
- **Import Excel** reads the KY Local chart, **one year (tab) at a time**. Tick "Replace" so re-importing a tab never duplicates.
- Drive-time turnaround checks are **off** (they only know track-to-track times).
- Sign-in page: a van pulls in, then Microsoft sign-in. After sign-in, a spoken "Welcome, <first name>" (once per person per day, after a click on *Enter* because browsers block autoplay audio).
- Put a photo of a real van on the sign-in page: set `VAN_PHOTO_URL` near the top of the "LOCAL BOARD ADDITIONS" script section.

## One-time setup
1. **Lists** (Graph can't create lists on this tenant): on the *Sallee Dispatch Board* SharePoint site choose New > List > **From Excel** and upload `setup/Local-Drivers.xlsx` (name it exactly **Local Drivers**) and `setup/Local-Jobs.xlsx` (name it exactly **Local Jobs**). Delete the sample row in each. Optional: change `Notes` to *multiple lines of text*.
2. **Permissions**: for each of the two lists: List settings > Permissions for this list > *Stop inheriting permissions*, then give **Edit** to the local dispatchers (Manu, astanalonis@salleehorsevans.com, ldolan@salleehorsevans.com) and leave everyone else at **Read**. (Site Members stay read-only on the OTR lists.)
3. **Sign-in**: in the existing *Sallee Dispatch Board* Entra app registration, add `https://canumanu.github.io/SALLEE-LOCAL-DISPATCH-BOARD/` as a Single-page-application redirect URI. No new consent needed.
4. Create a GitHub repo named `SALLEE-LOCAL-DISPATCH-BOARD`, upload `index.html` (and this README) to the root, then Settings > Pages > Deploy from branch > main / root.
5. Sign in, click **Import Excel**, pick the file, pick the tab, **Import**.

## Changing the job fields
The field list is deliberately easy to change; each field lives in these places in `index.html`:
`jobToFields` / `fieldsToJob` (SharePoint column mapping), the job modal HTML + `openJobModal` + the `jobSaveBtn` handler, `buildJobCard` (what shows on the card), and the parser `parseLocalChart` (what the Excel import fills). New SharePoint columns must also be added to the **Local Jobs** list.

## Files
- `index.html` the whole app
- `parse_local_chart.js` the Excel parser (also embedded in `index.html`; kept here for testing)
- `setup/` list templates for SharePoint
