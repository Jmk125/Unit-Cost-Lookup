# Unit Cost Lookup

Standalone Windows helper for macro-free Excel estimates.

## Setup

1. Paste `server_route.js` into the Raspberry Pi `server.js` before `app.listen(...)`.
2. Run:
   `node --check server.js`
   `sudo systemctl restart unit-costs`
3. Test `http://YOUR-PI-IP:3077/api/unit-cost-lookup`.
4. Edit `config.json` and enter the working Pi URL.
5. Install and run:
   `python -m pip install -r requirements.txt`
   `python unit_cost_lookup.py`

## Workflow

- Ctrl+Alt+U opens the app.
- Search and/or filter by division.
- Select a unit cost and edit the comment if desired.
- Populate Excel: click the button, select a cell in Excel, then press Ctrl+Alt+Enter.
- Populate Current Cell writes immediately to Excel's active cell.
- Copy Cost copies a clean numeric value.
- The Excel workbook can remain `.xlsx`.

## Build EXE

Run `build_exe.bat`. The EXE appears in `dist`.
Copy `config.json` beside the EXE.

## Startup

Press Win+R, enter `shell:startup`, and place a shortcut to the EXE there.


## Version 2 additions

- Dark square retro-inspired interface.
- Refresh button reloads live server data.
- Ctrl+Alt+U now toggles the window both open and hidden.
- Fixed the Windows hotkey listener by explicitly importing `ctypes.wintypes`.
- Open in Browser button.

For Open in Browser, configure this in `config.json`:

```json
"unit_detail_url_template": "http://YOUR-PI-IP:3077/unit-costs/{id}"
```

Replace the URL pattern with the actual individual-unit URL used by your web
application. Keep `{id}` as the placeholder for the selected database ID.


## Version 3 — View Details

Replace the older lookup route in `server.js` with the complete contents of
`server_route.js`. It now includes:

- `GET /api/unit-cost-lookup`
- `GET /api/unit-cost-lookup/:id`

The **View Details** button opens a live detail window containing:

- Unit-cost metadata and markups
- Comments
- Calculation scratch
- Material lines
- Labor lines
- Publication history

No browser URL template is required.


## Version 4 — Publication and stale-material indicator

The main table now includes a **Last Published** column.

A trailing `!` means at least one material referenced by the latest publication
snapshot has a `materials.date_updated` value older than one year.

Example:

```text
2025-04-18  !
```

Replace the prior lookup routes in `server.js` with the new contents of
`server_route.js`, then restart the service.


## Version 5 — Stale flag on individual materials

The Materials tab in View Details now includes a **Price Updated** column.

A material updated more than one year ago displays:

```text
2024-05-10  !
```

The detail API now joins each unit-cost material line to the live `materials`
record and returns `date_updated` plus `is_stale`.
