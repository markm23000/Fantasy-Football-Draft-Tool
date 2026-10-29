import json
import re
import threading
import unicodedata
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

HOST = "127.0.0.1"
PORT = 8765
VERSION = "2.4-ticker-reconcile-fixed"

try:
    import pythoncom
    import win32com.client
except Exception:
    print("\nMissing pywin32. Run Install Bridge.bat once, then start again.\n")
    raise

lock = threading.Lock()
state = {
    "app": None,
    "wb": None,
    "log_ws": None,
    "rank_ws": None,
    "existing_by_pick": {},
    "existing_players": set(),
    "player_to_pick": {},
    "page_stats": {},
}


def norm(value):
    s = unicodedata.normalize("NFKD", str(value or ""))
    s = s.replace("’", "'").replace("‘", "'").replace("`", "'")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().casefold()


def _rows(value):
    if value is None:
        return []
    if isinstance(value, tuple):
        return value
    return ((value,),)


def _last_data_row(ws):
    # xlUp = -4162. Check all four Draft Log data columns because gaps are
    # intentional in v2.4 (row number is tied to pick number).
    rows = []
    for col in range(1, 5):
        try:
            rows.append(ws.Cells(ws.Rows.Count, col).End(-4162).Row)
        except Exception:
            rows.append(1)
    return max(rows or [1])


def connect_excel(force=False):
    if not force and state["wb"] is not None:
        try:
            _ = state["wb"].Name
            return state["app"], state["wb"], state["log_ws"], state["rank_ws"]
        except Exception:
            pass

    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception as exc:
        raise RuntimeError("Excel is not open. Open your draft workbook first.") from exc

    chosen = None
    for wb in app.Workbooks:
        try:
            names = {ws.Name for ws in wb.Worksheets}
            if "Draft Log" in names and "Rankings Master" in names:
                chosen = wb
                break
        except Exception:
            continue
    if chosen is None:
        raise RuntimeError('No open workbook with both "Draft Log" and "Rankings Master" sheets was found.')

    log_ws = chosen.Worksheets("Draft Log")
    rank_ws = chosen.Worksheets("Rankings Master")
    state.update(app=app, wb=chosen, log_ws=log_ws, rank_ws=rank_ws)

    # Fix any rows written by v2.3 before building the cache. After this, every
    # numbered pick has one deterministic home: Excel row = pick + 1.
    canonicalize_draft_log()
    rebuild_excel_cache()
    return app, chosen, log_ws, rank_ws


def canonicalize_draft_log():
    ws = state["log_ws"]
    if ws is None:
        return

    last = _last_data_row(ws)
    if last < 2:
        return

    vals = ws.Range(f"A2:D{last}").Value
    records = {}
    valid_rows = 0
    needs_rewrite = False

    for idx, row in enumerate(_rows(vals), start=2):
        if not row:
            continue
        pick_raw = row[0] if len(row) > 0 else None
        player_raw = row[1] if len(row) > 1 else None
        drafter_raw = row[2] if len(row) > 2 else None
        stamp_raw = row[3] if len(row) > 3 else None
        player = str(player_raw or "").strip()
        if not player:
            continue
        try:
            pick = int(float(pick_raw))
        except Exception:
            needs_rewrite = True
            continue
        if pick <= 0 or pick > 500:
            needs_rewrite = True
            continue

        valid_rows += 1
        if idx != pick + 1 or pick in records:
            needs_rewrite = True
        # If duplicates exist, the later row wins because it is most likely the
        # newest correction from an older bridge version.
        records[pick] = (
            player,
            str(drafter_raw or "").strip(),
            stamp_raw,
        )

    if not records:
        return
    if len(records) != valid_rows:
        needs_rewrite = True
    if not needs_rewrite:
        return

    clear_end = max(last, max(records) + 1)
    ws.Range(f"A2:D{clear_end}").ClearContents()
    for pick in sorted(records):
        player, drafter, stamp = records[pick]
        row = pick + 1
        if stamp in (None, ""):
            stamp = datetime.now().strftime("%I:%M:%S %p")
        ws.Range(f"A{row}:D{row}").Value = ((pick, player, drafter, stamp),)

    print(f"Reconciled Draft Log into pick-number rows ({len(records)} existing picks).", flush=True)


def rebuild_excel_cache():
    ws = state["log_ws"]
    if ws is None:
        return

    last = _last_data_row(ws)
    by_pick = {}
    players = set()
    player_to_pick = {}

    if last >= 2:
        vals = ws.Range(f"A2:C{last}").Value
        for idx, row in enumerate(_rows(vals), start=2):
            if not row:
                continue
            pick_raw = row[0] if len(row) > 0 else None
            player_raw = row[1] if len(row) > 1 else None
            drafter_raw = row[2] if len(row) > 2 else None
            player = str(player_raw or "").strip()
            if not player:
                continue
            try:
                pick = int(float(pick_raw))
            except Exception:
                continue
            if pick <= 0 or pick > 500:
                continue
            nk = norm(player)
            drafter = str(drafter_raw or "").strip()
            by_pick[pick] = (idx, player, drafter)
            players.add(nk)
            player_to_pick[nk] = pick

    state["existing_by_pick"] = by_pick
    state["existing_players"] = players
    state["player_to_pick"] = player_to_pick


def set_status(text=None, last_pick=None):
    try:
        ws = state["log_ws"]
        if ws is None:
            return
        if text is not None:
            ws.Range("J2").Value = text
        if last_pick is not None:
            ws.Range("J3").Value = last_pick
    except Exception:
        pass


def _clean_pick(item):
    if not isinstance(item, dict):
        return None
    player = str(item.get("player") or item.get("name") or "").strip()
    if len(player) < 2 or len(player) > 80:
        return None
    try:
        pick = int(float(item.get("pick") or item.get("overall")))
    except Exception:
        return None
    if pick <= 0 or pick > 500:
        return None
    return {
        "pick": pick,
        "player": player,
        "round": item.get("round"),
        "round_pick": item.get("round_pick"),
        "source": str(item.get("source") or "ticker"),
        "drafter": str(item.get("drafter") or "").strip(),
    }


def upsert_picks(items):
    # Pick number is authoritative. If the same numbered pick appears more than
    # once in one POST, the newest version wins (useful when team text fills in).
    by_pick = {}
    for item in items or []:
        p = _clean_pick(item)
        if p:
            by_pick[p["pick"]] = p
    cleaned = [by_pick[k] for k in sorted(by_pick)]
    if not cleaned:
        return {"ok": True, "added": 0, "updated": 0, "ignored": 0}

    with lock:
        app, wb, ws, rank_ws = connect_excel()
        added = 0
        updated = 0
        ignored = 0

        for p in cleaned:
            pick = p["pick"]
            player = p["player"]
            drafter = p.get("drafter", "")
            nk = norm(player)
            row = pick + 1
            existing = state["existing_by_pick"].get(pick)

            if existing:
                _, old_player, old_drafter = existing
                same_player = norm(old_player) == nk
                team_improved = bool(drafter) and norm(drafter) != norm(old_drafter)

                if same_player and not team_improved:
                    ignored += 1
                    continue

                # The Picks ticker is authoritative for this pick number. Correct
                # player and/or fill in a later-arriving Drafted By team.
                stamp = datetime.now().strftime("%I:%M:%S %p")
                ws.Range(f"A{row}:D{row}").Value = ((pick, player, drafter or old_drafter, stamp),)

                if not same_player:
                    state["existing_players"].discard(norm(old_player))
                    state["player_to_pick"].pop(norm(old_player), None)
                state["existing_players"].add(nk)
                state["player_to_pick"][nk] = pick
                state["existing_by_pick"][pick] = (row, player, drafter or old_drafter)
                updated += 1
                print(f"Corrected #{pick}: {player} — {(drafter or old_drafter) or 'team unavailable'}", flush=True)
                continue

            # If the same player somehow exists under a different pick from an old
            # bad capture, move him to the ticker-authoritative numbered pick.
            other_pick = state["player_to_pick"].get(nk)
            if other_pick and other_pick != pick:
                old_row = other_pick + 1
                ws.Range(f"A{old_row}:D{old_row}").ClearContents()
                state["existing_by_pick"].pop(other_pick, None)
                state["player_to_pick"].pop(nk, None)

            stamp = datetime.now().strftime("%I:%M:%S %p")
            ws.Range(f"A{row}:D{row}").Value = ((pick, player, drafter, stamp),)
            state["existing_by_pick"][pick] = (row, player, drafter)
            state["existing_players"].add(nk)
            state["player_to_pick"][nk] = pick
            added += 1
            print(f"Captured #{pick}: {player} — {drafter or 'team unavailable'}", flush=True)

        if added or updated:
            last = max(cleaned, key=lambda x: x["pick"])
            set_status("Connected ✓ — PICKS TICKER v2.4", f'#{last["pick"]} {last["player"]}')
            # Automatic Excel calculation updates T:V by itself. Only calculate the
            # small status/ranking range if the workbook is in Manual mode.
            try:
                if int(app.Calculation) == -4135:  # xlCalculationManual
                    rank_ws.Range("T4:V203").Calculate()
            except Exception:
                pass

        return {
            "ok": True,
            "added": added,
            "updated": updated,
            "ignored": ignored,
            "latest": max((p["pick"] for p in cleaned), default=0),
        }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return

    def _json(self, body, code=200):
        raw = json.dumps(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_OPTIONS(self):
        self._json({}, 204)

    def do_GET(self):
        if self.path != "/status":
            self._json({"ok": False, "error": "not found"}, 404)
            return
        try:
            _, wb, _, _ = connect_excel()
            set_status("Connected ✓ — PICKS TICKER v2.4")
            self._json({
                "ok": True,
                "version": VERSION,
                "workbook": wb.Name,
                "drafted_rows": len(state["existing_by_pick"]),
                "page": state["page_stats"],
            })
        except Exception as exc:
            self._json({"ok": False, "version": VERSION, "error": str(exc)}, 503)

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length) or b"{}")
            if self.path == "/picks":
                self._json(upsert_picks(payload.get("picks", [])))
            elif self.path == "/stats":
                state["page_stats"] = payload if isinstance(payload, dict) else {}
                self._json({"ok": True})
            elif self.path == "/refresh":
                connect_excel(force=True)
                self._json({"ok": True, "message": "Excel cache refreshed + Draft Log reconciled."})
            elif self.path == "/test":
                pick = payload.get("pick", 999)
                player = payload.get("player", "Ja'Marr Chase")
                self._json(upsert_picks([{"pick": pick, "player": player, "source": "test"}]))
            else:
                self._json({"ok": False, "error": "not found"}, 404)
        except Exception as exc:
            try:
                connect_excel(force=True)
            except Exception:
                pass
            self._json({"ok": False, "error": str(exc)}, 500)


def main():
    print("ESPN → Excel Draft Bridge v2.4 FIXED — RECONCILING PICKS TICKER + TEAM")
    print("---------------------------------------------------------------")
    print(f"Listening locally on http://{HOST}:{PORT}")
    print("Source of truth: ESPN Picks sidebar cards (including transition/hidden card text).")
    print("Pick number is authoritative. Excel row = pick + 1, so out-of-order arrivals cannot scramble the log.")
    print("No WebSocket parsing. No player-pool scanning. Drafted By comes from the Picks ticker card.")
    print("Keep this window open during your draft.\n")
    try:
        _, wb, _, _ = connect_excel(force=True)
        set_status("Connected ✓ — PICKS TICKER v2.4")
        print(f"Excel workbook found: {wb.Name}")
        print(f"Existing Draft Log picks: {len(state['existing_by_pick'])}")
    except Exception as exc:
        print(f"Excel check: {exc}")
        print("Open your workbook, then use Check Connection in the Chrome extension.")
    print()
    server = HTTPServer((HOST, PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
