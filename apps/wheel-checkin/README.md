# Wheel Check-in

A single-file emotion journal built as a Claude Artifact. It uses the Junto
Institute Emotion & Feeling Wheel (six core emotions, each with secondary and
specific feelings) and captures, per entry:

- one or more feelings picked from the wheel
- an intensity rating from 1 (faint) to 5 (intense)
- the situation
- commentary on why the feeling might be there

Entries are tagged as a midday check-in, an evening check-in, or an anytime
note. The two cards at the top of the page show whether today's midday and
evening check-ins are logged, due, or still ahead. Check-in times are
configurable and kept in the browser.

## Storage

Published on claude.ai with the `db` capability, entries are saved in the
artifact's own document store under the `entries` collection, so they follow
the owner across devices and can be read back by Claude for summaries. Opened
as a plain file, the page falls back to browser storage and says so.

## Files

- `wheel-checkin.html` is the page body published to the artifact. The
  artifact runtime wraps it in the document skeleton, so it carries no
  `<html>` or `<body>` tags of its own.

Live artifact: https://claude.ai/artifact/2JWRwcKRovbpUR6zQr8Gh6
