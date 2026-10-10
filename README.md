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

The app includes a built-in event-focused workbook layout based on the uploaded blank floor-plan template's grid coordinates, floor numbering, and venue landmarks. For a **full-fidelity export using the complete original master**, select **Excel master** and choose the blank .xlsx file before exporting. That file is read only in your browser; the original is never overwritten. The precise master is not stored in this public repository.

Browser session event edits are saved in local storage; imported PDFs and original XLSX files themselves are not saved or transmitted. Re-upload the XLSX master after reloading if you need complete template fidelity.

## Development / hosting

No build step. GitHub Pages from the repository's **main** branch, root folder. Dependencies served by CDNs:
- pdf.js 4.10.38
- xlsx-js-style 1.2.0

## Notes

Review the extracted event categories, golf-only times, and multi-space events before use: BEO formatting and descriptions can vary. This is a floor-plan scheduling aid, not a live Salesforce/BEO integration.
