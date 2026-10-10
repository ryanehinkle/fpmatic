# FPmatic

Topgolf Naperville event floor-plan generator. Static GitHub Pages app, styled to match [ARMatic](https://github.com/ryanehinkle/armatic).

## Usage

Open https://ryanehinkle.github.io/fpmatic/.

**Main workflow:** Select the two Salesforce CSV exports, **Event Staffing** and **List of Catering Menus - Specific Date**, then click **Build floor plan**. The Event Staffing report is the authority for event records, dates, overall event times, bay ranges, guest counts, tentative/definite status, room products, and golf-only times. Catering Menus is matched on the full Opportunity Name to determine food packages and event colors. Duplicate product rows are grouped into one event; bookings without menu rows remain included. The parser flags events without assigned bays for review instead of estimating.

There is also a small **Use BEO PDF instead** option, preserved as an alternative.

All data stays in the browser. Imported report content is not transmitted to GitHub or any server. The selected report filenames and event records are persisted in browser local storage; CSV contents are not committed to the public repository.

### Floor-plan rules

- 102 bays: first floor 101–134, second 201–234, third 301–334.
- An event is **PM if it starts at or after 4 PM, or ends at or after 7 PM**. All other events are AM.
- Exact bay assignments come only from the BEO; missing assignments are **never invented**.
- Blocks are three rows tall by default. Where another event uses an overlapping bay range in the same half of the day, the blocks are two rows tall.
- Event labels: event name; overall event time preceded by booked room(s) if any; separate "Golf ..." line if gameplay times differ; guest count; package name.
- Definite events use pale fill colors; tentative bookings (including "Agreement Sent") use stronger fills. Field Trip definite is lavender \`#D9D2E9\`.
- Yellow-numbered bays are couch bays; white bays are high-tops. Neutral header bays preserve the template's unfilled header groups.

### Excel export and master

FPmatic ships a **format-preserved copy of the exact original blank floor-plan workbook** containing all seven individually formatted daily sheets and all 11 original tabs. Every daily floor-plan cell, merge, row height, column width, landmark, border, and Excel style matches the uploaded master. Only individual employee names in the staff reference lists were redacted before shipping a public copy.

The exporter uses local JSZip and surgically updates **only occupied event cells, their new merge ranges, and newly appended event styles** in the original XLSX ZIP. It never rebuilds the original worksheets using the spreadsheet conversion library.

To retain the private employee reference names in the export, choose **Excel master** and select your complete original `.xlsx` file. Everything stays in your browser and the original is never overwritten.

## Regression checks

The GitHub Pages workflow runs the original-master Excel preservation tests, floor stacking/border checks, and the CSV report join tests before deploying.

## Development / hosting

No build step. GitHub Pages deployment runs automatically via `.github/workflows/pages.yml` after Pages is enabled. The repository owner must first open **Settings → Pages → Build and deployment → Source → GitHub Actions**, save, then rerun the deployment workflow in **Actions**. The workflow cannot create a Pages site if the repository's token lacks Pages administration permission. Dependencies served by CDNs:
- pdf.js 4.10.38
- xlsx-js-style 1.2.0

## Notes

Review the extracted event categories, golf-only times, and multi-space events before use: BEO formatting and descriptions can vary. This is a floor-plan scheduling aid, not a live Salesforce/BEO integration.
