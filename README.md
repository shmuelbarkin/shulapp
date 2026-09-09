# Shul App — Manage My Shul

Google Apps Scripts and a static website for tools that help manage a Shul/synagogue.

> **This project is archived.** Its successor is **[ShulBase](https://shulbase.com)**, where
> members, the luach and zmanim, messaging, pledges and billing are built as a real
> application rather than spreadsheets and scripts. The material here still works and
> remains free to use, but it is no longer maintained.

## Layout

| Path | What it is |
|---|---|
| `scripts/` | The Google Apps Script source for the sheets |
| `website/` | [managemyshul.com](https://managemyshul.com) — a static site, no build step |

## Scripts

### Accounting app

Features of this sheet:

- Billing automation
- Membership management
- Simcha list and messaging

Scripts for this sheet are in `scripts/Accounting`.

### Messaging app

Features of this sheet:

- Send emails and SMS
- Add attachments
- Schedule messages

Scripts for this sheet are in `scripts/Messaging`.

## Website

Plain HTML and CSS in `website/`, with no dependencies.

`website/assets/` holds the sources: `style.css`, `banner.js` (dismisses the
archive notice) and `nav.js` (replaces `<main>` on a link click instead of
loading a new document). `build.sh` inlines all three into every page, so each
page is a complete, self-contained document. Edit an asset, then run the script:

```bash
./build.sh
cd website && python3 -m http.server 8000
```

## Usage

The simplest way to use these scripts is to make a copy of the Google Sheets that have
the scripts already added to them. Head over to <https://managemyshul.com/> for links to
the sheets, more info and setup instructions.
