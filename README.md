# FPmatic

Topgolf Naperville event floor-plan generator. Static GitHub Pages app, styled to match [ARMatic](https://github.com/ryanehinkle/armatic).

## Usage

Open https://ryanehinkle.github.io/fpmatic/ and upload one or more combined Manager Catering BEO PDF packets. The PDFs are processed **locally in the browser** using PDF.js; no backend, credentials, or PDF upload service is involved.

FPmatic extracts events, dates, times, guest counts, booking status, bay assignments, room/space information, and main food/package labels. It automatically plots the bay ranges for each of seven daily worksheets (Monday–Sunday), separated into floors and AM/PM sections. Click any event to correct parsed details. Events with missing/invalid bays remain in the event list for review; bay/time overlaps are flagged.

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

## Development / hosting

No build step. GitHub Pages deployment runs automatically via `.github/workflows/pages.yml` after Pages is enabled. The repository owner must first open **Settings → Pages → Build and deployment → Source → GitHub Actions**, save, then rerun the deployment workflow in **Actions**. The workflow cannot create a Pages site if the repository's token lacks Pages administration permission. Dependencies served by CDNs:
- pdf.js 4.10.38
- xlsx-js-style 1.2.0

## Notes

Review the extracted event categories, golf-only times, and multi-space events before use: BEO formatting and descriptions can vary. This is a floor-plan scheduling aid, not a live Salesforce/BEO integration.
