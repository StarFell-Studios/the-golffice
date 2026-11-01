# Content review

A client review page: each piece of content gets Approve, Maybe or No plus a note. Votes save to a Google Sheet through a small Apps Script.

- Live: https://starfell-studios.github.io/the-golffice/review/
- Reviewer links: add `?r=michael`, `?r=kerry` or `?r=mike`.

## New batch
1. Put the new files in `media/`.
2. Edit `content.json`: change `batch` (for example `2026-12-v1`), the title and intro, and the `items` list.
3. Commit. Votes for each batch are kept separately in the Sheet.

## Vote storage (one time)
1. Make a Google Sheet, open Extensions, then Apps Script, paste `apps-script/Code.gs`, run `setup` once and authorize.
2. Deploy, New deployment, type Web app, Execute as Me, Who has access Anyone. Copy the web app URL.
3. Paste that URL into `apiUrl` in `config.js` and commit.
