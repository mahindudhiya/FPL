"""
FPL Weekly Advisor
==================
Run once a week (after the previous gameweek finishes, before the next deadline):

    python fpl_weekly.py YOUR_TEAM_ID

Your team ID is the number in the URL of your Points page on the FPL site:
fantasy.premierleague.com/entry/1234567/event/9  ->  team ID is 1234567

It pulls live data from FPL's public API, scores every player on form,
underlying stats, fixtures, minutes and fitness, then writes ONE Excel file:
    FPL_Weekly_Report_GW<next>.xlsx

Needs: Python 3.9+ and openpyxl  (pip install openpyxl)
"""

import json
import os
import sys
import time
import unicodedata
import urllib.request
from bisect import bisect_left
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ============================ SETTINGS (edit these) ============================
TEAM_ID = None            # or put your ID here, e.g. 1234567, and run with no argument
LOOKAHEAD = 5             # how many upcoming gameweeks the fixture score looks at
MIN_MINUTES = 270         # players below this are still scored, but not used to set the benchmarks

# How much each factor counts in the 0-100 score (should add up to 1.0)
WEIGHTS = {
    "form": 0.30,         # recent FPL points (last ~30 days) blended with season points per game
    "underlying": 0.25,   # attackers: xG+xA per 90. Defenders/GKs: expected goals conceded (and xGI for DEF)
    "fixtures": 0.25,     # official FPL fixture difficulty over the next LOOKAHEAD gameweeks
    "ep_next": 0.20,      # FPL's own expected points for the next gameweek
}

SELL_BELOW = 35           # score under this = SELL
WATCH_BELOW = 50          # score under this = WATCH
MIN_GAIN = 15             # a transfer must beat your player by this many points, otherwise roll it

# Transfers you've already made this week. They only appear in the API after the deadline,
# so list them here as (player out, player in), using names as shown on the FPL site.
PENDING_TRANSFERS = [
    # ("João Pedro", "N.Jackson"),
]
# ==============================================================================

API = "https://fantasy.premierleague.com/api/"
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fpl_cache")
POS = {1: "GK", 2: "DEF", 3: "MID", 4: "FWD"}


def get(path, offline=False):
    """Fetch JSON from the FPL API, keeping a local copy in fpl_cache/."""
    os.makedirs(CACHE, exist_ok=True)
    fname = os.path.join(CACHE, path.strip("/").replace("/", "_") + ".json")
    if offline:
        with open(fname, encoding="utf-8") as fh:
            return json.load(fh)
    req = urllib.request.Request(API + path, headers={"User-Agent": "Mozilla/5.0 (fpl-weekly-advisor)"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                data = json.loads(r.read().decode("utf-8"))
            with open(fname, "w", encoding="utf-8") as fh:
                json.dump(data, fh)
            return data
        except Exception as exc:  # FPL is often slow right after a gameweek
            if attempt == 2:
                raise SystemExit(f"Could not reach the FPL API ({exc}). "
                                 "If the game is updating, try again in an hour.")
            time.sleep(3)


def norm(name):
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().lower()
    return s.replace(".", "").replace(" ", "")


def f(x, default=0.0):
    try:
        return float(x)
    except (TypeError, ValueError):
        return default


def percentile(sorted_vals, v):
    if not sorted_vals:
        return 50.0
    return 100.0 * bisect_left(sorted_vals, v) / len(sorted_vals)


def availability(e):
    chance = e.get("chance_of_playing_next_round")
    if chance is not None:
        return chance / 100.0
    status = e.get("status", "a")
    if status == "a":
        return 1.0
    if status == "d":
        return 0.5
    return 0.0  # injured, suspended, unavailable, on loan


# ------------------------------------------------------------------------------
def analyse(team_id, offline=False):
    boot = get("bootstrap-static/", offline)
    fixtures = get("fixtures/", offline)
    teams = {t["id"]: t for t in boot["teams"]}
    events = boot["events"]

    nxt = next((ev for ev in events if ev.get("is_next")), None)
    cur = next((ev for ev in events if ev.get("is_current")), None)
    if nxt is None:
        raise SystemExit("No upcoming gameweek found (season over?).")
    next_gw = nxt["id"]
    gws = [g for g in range(next_gw, next_gw + LOOKAHEAD) if g <= 38]
    deadline = datetime.fromisoformat(nxt["deadline_time"].replace("Z", "+00:00")).astimezone()

    # ---- fixtures: games played, upcoming opponents, fixture ease ----
    played = defaultdict(int)
    upcoming = defaultdict(lambda: defaultdict(list))  # team -> gw -> [(text, fdr)]
    for fx in fixtures:
        h, a = fx["team_h"], fx["team_a"]
        if fx.get("finished"):
            played[h] += 1
            played[a] += 1
        if fx.get("event") in gws:
            upcoming[h][fx["event"]].append((f"{teams[a]['short_name']} (H)", fx["team_h_difficulty"]))
            upcoming[a][fx["event"]].append((f"{teams[h]['short_name']} (A)", fx["team_a_difficulty"]))

    raw_ease = {}
    for tid in teams:
        total = sum((6 - fdr) for g in gws for _, fdr in upcoming[tid][g])  # blanks add 0, doubles add twice
        raw_ease[tid] = total / max(len(gws), 1)
    lo, hi = min(raw_ease.values()), max(raw_ease.values())
    ease = {t: (100.0 * (v - lo) / (hi - lo) if hi > lo else 50.0) for t, v in raw_ease.items()}

    # ---- player metrics ----
    players = {}
    for e in boot["elements"]:
        if e["element_type"] not in POS:
            continue
        games = max(played[e["team"]], 1)
        starts = e.get("starts")
        share = (starts / games) if starts is not None else e["minutes"] / (90.0 * games)
        players[e["id"]] = {
            "id": e["id"], "name": e["web_name"], "pos": e["element_type"], "team": e["team"],
            "club": teams[e["team"]]["short_name"], "price": e["now_cost"] / 10.0,
            "form": f(e.get("form")), "ppg": f(e.get("points_per_game")), "ep": f(e.get("ep_next")),
            "xgi90": f(e.get("expected_goal_involvements_per_90")),
            "xgc90": f(e.get("expected_goals_conceded_per_90")),
            "minutes": e["minutes"], "share": min(share, 1.0),
            "avail": availability(e), "news": e.get("news", ""), "owned_pct": f(e.get("selected_by_percent")),
            "fix": ease[e["team"]],
        }

    # benchmarks per position (only regular players set the bar)
    bench = defaultdict(lambda: defaultdict(list))
    for p in players.values():
        if p["minutes"] >= MIN_MINUTES:
            b = bench[p["pos"]]
            b["form"].append(0.6 * p["form"] + 0.4 * p["ppg"])
            b["xgi"].append(p["xgi90"])
            b["cs"].append(-p["xgc90"])
            b["ep"].append(p["ep"])
    for b in bench.values():
        for k in b:
            b[k].sort()

    for p in players.values():
        b = bench[p["pos"]]
        p["form_pct"] = percentile(b["form"], 0.6 * p["form"] + 0.4 * p["ppg"])
        xgi_pct, cs_pct = percentile(b["xgi"], p["xgi90"]), percentile(b["cs"], -p["xgc90"])
        p["under_pct"] = cs_pct if p["pos"] == 1 else (0.5 * cs_pct + 0.5 * xgi_pct if p["pos"] == 2 else xgi_pct)
        p["ep_pct"] = percentile(b["ep"], p["ep"])
        raw = (WEIGHTS["form"] * p["form_pct"] + WEIGHTS["underlying"] * p["under_pct"]
               + WEIGHTS["fixtures"] * p["fix"] + WEIGHTS["ep_next"] * p["ep_pct"])
        p["mins_factor"] = max(0.2, min(1.0, p["share"] / 0.8))
        p["score"] = round(raw * p["avail"] * p["mins_factor"], 1)

    # ---- your squad ----
    ref_gw = cur["id"] if cur else next_gw - 1
    entry = get(f"entry/{team_id}/", offline)
    picks_data = get(f"entry/{team_id}/event/{ref_gw}/picks/", offline)
    squad = [pk["element"] for pk in picks_data["picks"]]
    bank = picks_data.get("entry_history", {}).get("bank", 0) / 10.0

    by_name = defaultdict(list)
    for p in players.values():
        by_name[norm(p["name"])].append(p["id"])
    warnings = []
    for out_name, in_name in PENDING_TRANSFERS:
        out_ids = [i for i in by_name.get(norm(out_name), []) if i in squad]
        in_ids = by_name.get(norm(in_name), [])
        if not out_ids or len(in_ids) != 1:
            warnings.append(f"Could not apply pending transfer {out_name} -> {in_name}; check the spelling.")
            continue
        bank += players[out_ids[0]]["price"] - players[in_ids[0]]["price"]
        squad[squad.index(out_ids[0])] = in_ids[0]

    club_count = defaultdict(int)
    for pid in squad:
        club_count[players[pid]["team"]] += 1

    # verdicts
    for pid in squad:
        p = players[pid]
        reasons = []
        if p["avail"] < 1:
            reasons.append(f"{int(p['avail'] * 100)}% to play" + (f": {p['news']}" if p["news"] else ""))
        if p["mins_factor"] < 0.75:
            reasons.append("not starting regularly")
        if p["fix"] < 35:
            reasons.append("tough fixtures")
        if p["form_pct"] < 35:
            reasons.append("poor form")
        if p["under_pct"] < 35:
            reasons.append("weak underlying stats")
        if p["avail"] <= 0.25 or p["score"] < SELL_BELOW:
            p["verdict"] = "SELL"
        elif p["avail"] < 1 or p["score"] < WATCH_BELOW:
            p["verdict"] = "WATCH"
        else:
            p["verdict"] = "HOLD"
        p["reason"] = "; ".join(reasons) if reasons else "good form, minutes and fixtures"

    # transfer suggestions: best like-for-like upgrade for each squad player
    options = []
    owned = set(squad)
    for pid in squad:
        out = players[pid]
        budget = out["price"] + bank
        for c in players.values():
            if c["id"] in owned or c["pos"] != out["pos"] or c["price"] > budget + 1e-9:
                continue
            if c["avail"] < 0.75 or c["mins_factor"] < 0.75:
                continue
            if club_count[c["team"]] - (1 if c["team"] == out["team"] else 0) >= 3:
                continue
            if c["score"] > out["score"]:
                options.append((round(c["score"] - out["score"], 1), out, c))
    options.sort(key=lambda m: -m[0])
    moves, used_out, used_in = [], set(), set()
    for gain, out, c in options:  # best move per seller, and never the same target twice
        if out["id"] in used_out or c["id"] in used_in:
            continue
        moves.append((gain, out, c))
        used_out.add(out["id"])
        used_in.add(c["id"])
        if len(moves) == 5:
            break

    # captain: highest expected points x availability among your outfield and GK players
    captains = sorted((players[i] for i in squad), key=lambda p: -(p["ep"] * p["avail"]))[:3]

    # targets: top 6 per position you don't own
    targets = {}
    for pos in POS:
        pool = [p for p in players.values() if p["pos"] == pos and p["id"] not in owned
                and p["avail"] >= 0.75 and p["mins_factor"] >= 0.75]
        targets[pos] = sorted(pool, key=lambda p: -p["score"])[:6]

    return {
        "entry": entry, "next_gw": next_gw, "gws": gws, "deadline": deadline, "bank": bank,
        "squad": sorted((players[i] for i in squad), key=lambda p: (p["pos"], -p["score"])),
        "moves": moves[:5], "captains": captains, "targets": targets,
        "teams": teams, "upcoming": upcoming, "ease": ease, "warnings": warnings,
    }


# ------------------------------------------------------------------------------
FDR_FILL = {1: ("1E7B34", "FFFFFF"), 2: ("A9D18E", "000000"), 3: ("E7E6E6", "000000"),
            4: ("F4B183", "000000"), 5: ("C00000", "FFFFFF")}
VERDICT_FILL = {"SELL": "F8CBAD", "WATCH": "FFE699", "HOLD": "C6EFCE"}


def write_report(r, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly Report"
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    navy = PatternFill("solid", fgColor="1F2A44")
    band = PatternFill("solid", fgColor="D9E1F2")
    centre = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def font(**k):
        return Font(name="Arial", **k)

    row = [1]

    def section(title):
        row[0] += 1
        for c in range(1, 16):
            ws.cell(row=row[0], column=c).fill = band
        ws.cell(row=row[0], column=1, value=title).font = font(bold=True, size=11)
        row[0] += 1

    def header(cols):
        for j, h in enumerate(cols, 1):
            c = ws.cell(row=row[0], column=j, value=h)
            c.font, c.fill, c.alignment, c.border = font(bold=True, color="FFFFFF", size=9), navy, centre, border
        row[0] += 1

    def line(values, bold_cols=()):
        for j, v in enumerate(values, 1):
            c = ws.cell(row=row[0], column=j, value=v)
            c.font = font(size=10, bold=j in bold_cols)
            c.border = border
            c.alignment = left if isinstance(v, str) and len(v) > 12 else centre
        row[0] += 1
        return row[0] - 1

    def fixture_cells(rw, start_col, team_id):
        for k, g in enumerate(r["gws"]):
            fx = r["upcoming"][team_id][g]
            c = ws.cell(row=rw, column=start_col + k)
            c.border, c.alignment = border, centre
            if not fx:
                c.value, c.fill, c.font = "BLANK", PatternFill("solid", fgColor="000000"), font(size=9, color="FFFFFF", bold=True)
                continue
            c.value = " + ".join(t for t, _ in fx)
            avg = round(sum(d for _, d in fx) / len(fx))
            bg, fg = FDR_FILL.get(avg, FDR_FILL[3])
            c.fill, c.font = PatternFill("solid", fgColor=bg), font(size=9, color=fg, bold=len(fx) > 1)

    e = r["entry"]
    ws["A1"] = f"FPL Weekly Report: GW{r['next_gw']}  |  {e.get('name', '')}"
    ws["A1"].font = font(bold=True, size=14)
    row[0] = 2
    ws.cell(row=2, column=1, value=(f"Deadline: {r['deadline']:%a %b %d, %I:%M %p} (your local time)   |   "
                                    f"Bank: £{r['bank']:.1f}m   |   Generated {datetime.now():%b %d %Y, %I:%M %p}")).font = font(size=10)
    for w in r["warnings"]:
        row[0] += 1
        ws.cell(row=row[0], column=1, value="WARNING: " + w).font = font(size=10, bold=True, color="C00000")

    # 1. headline moves
    section("1. RECOMMENDED TRANSFERS (best like-for-like upgrade, within your budget and the 3-per-club limit)")
    if not r["moves"] or r["moves"][0][0] < MIN_GAIN:
        line([f"ROLL YOUR TRANSFER: no upgrade beats your current player by {MIN_GAIN}+ points this week."], (1,))
    header(["Gain", "Sell", "Club", "Price", "Score", "Buy", "Club", "Price", "Score", "Verdict on seller"])
    for gain, out, inn in r["moves"]:
        rw = line([gain, out["name"], out["club"], out["price"], out["score"],
                   inn["name"], inn["club"], inn["price"], inn["score"], out["verdict"]], (1, 2, 6))
        if gain >= MIN_GAIN:
            ws.cell(row=rw, column=1).fill = PatternFill("solid", fgColor="C6EFCE")

    # 2. captain
    section("2. CAPTAIN OPTIONS (FPL expected points for next GW, adjusted for fitness)")
    header(["Rank", "Player", "Club", "Exp. pts", "Next fixture"])
    for i, p in enumerate(r["captains"], 1):
        nf = r["upcoming"][p["team"]][r["gws"][0]]
        line([i, p["name"], p["club"], round(p["ep"] * p["avail"], 1), " + ".join(t for t, _ in nf) or "BLANK"], (2,))

    # 3. squad
    section("3. MY SQUAD")
    cols = ["Pos", "Player", "Club", "Price", "Fit %", "Form", "PPG", "xGI/90", "Starts %", "Score", "Verdict", "Reason"]
    gw_cols = [f"GW{g}" for g in r["gws"]]
    header(cols + gw_cols)
    for p in r["squad"]:
        rw = line([POS[p["pos"]], p["name"], p["club"], p["price"], int(p["avail"] * 100), p["form"], p["ppg"],
                   round(p["xgi90"], 2), int(p["share"] * 100), p["score"], p["verdict"], p["reason"]], (2, 10, 11))
        ws.cell(row=rw, column=11).fill = PatternFill("solid", fgColor=VERDICT_FILL[p["verdict"]])
        fixture_cells(rw, len(cols) + 1, p["team"])

    # 4. targets
    section("4. TOP TARGETS BY POSITION (players you don't own, fit and starting)")
    header(["Pos", "Player", "Club", "Price", "Owned %", "Form", "PPG", "xGI/90", "Starts %", "Score", "", ""] + gw_cols)
    for pos in POS:
        for p in r["targets"][pos]:
            rw = line([POS[pos], p["name"], p["club"], p["price"], p["owned_pct"], p["form"], p["ppg"],
                       round(p["xgi90"], 2), int(p["share"] * 100), p["score"], "", ""], (2, 10))
            fixture_cells(rw, len(cols) + 1, p["team"])

    # 5. fixture ticker
    section(f"5. FIXTURE TICKER: all 20 teams, easiest run first (official FPL difficulty, next {len(r['gws'])} GWs)")
    header(["Rank", "Team", "Ease (0-100)"] + [""] * (len(cols) - 3) + gw_cols)
    ranked = sorted(r["teams"], key=lambda t: -r["ease"][t])
    for i, t in enumerate(ranked, 1):
        rw = line([i, r["teams"][t]["name"], round(r["ease"][t])] + [""] * (len(cols) - 3), (2,))
        fixture_cells(rw, len(cols) + 1, t)

    # 6. method
    section("6. HOW THE SCORE WORKS")
    notes = [
        f"Score (0-100) = {int(WEIGHTS['form']*100)}% form + {int(WEIGHTS['underlying']*100)}% underlying stats + "
        f"{int(WEIGHTS['fixtures']*100)}% fixtures + {int(WEIGHTS['ep_next']*100)}% FPL expected points, each ranked against "
        "players in the same position, then multiplied by chance of playing and how regularly the player starts.",
        f"Verdicts: SELL below {SELL_BELOW} or 25% or less to play. WATCH below {WATCH_BELOW} or any fitness flag. HOLD otherwise.",
        "Prices are current prices. Your real selling price can be lower if a player rose after you bought him, so check the transfer screen.",
        "Data: Fantasy Premier League public API (fantasy.premierleague.com/api). Weights and thresholds are at the top of fpl_weekly.py.",
        "This is a guide, not a rule: always check press conferences before the deadline.",
    ]
    for n in notes:
        ws.cell(row=row[0], column=1, value=n).font = font(size=9, italic=True)
        ws.merge_cells(start_row=row[0], start_column=1, end_row=row[0], end_column=len(cols) + len(gw_cols))
        ws.cell(row=row[0], column=1).alignment = left
        ws.row_dimensions[row[0]].height = 26
        row[0] += 1

    widths = [7, 15, 7, 8, 8, 7, 7, 8, 9, 8, 9, 34] + [13] * len(gw_cols)
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = w
    ws.freeze_panes = "C4"
    wb.save(path)


# ------------------------------------------------------------------------------
if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    offline = "--offline" in sys.argv     # re-run using the last downloaded data
    team_id = int(args[0]) if args else TEAM_ID
    if not team_id:
        raise SystemExit("Usage: python fpl_weekly.py YOUR_TEAM_ID")

    result = analyse(team_id, offline)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"FPL_Weekly_Report_GW{result['next_gw']}.xlsx")
    write_report(result, out)

    print(f"\nGW{result['next_gw']} | deadline {result['deadline']:%a %b %d %I:%M %p} | bank £{result['bank']:.1f}m")
    for w in result["warnings"]:
        print("WARNING:", w)
    if result["moves"] and result["moves"][0][0] >= MIN_GAIN:
        g, o, i = result["moves"][0]
        print(f"Top move: {o['name']} -> {i['name']}  (+{g})")
    else:
        print("Top move: roll the transfer")
    print("Captain:", result["captains"][0]["name"])
    for p in result["squad"]:
        if p["verdict"] != "HOLD":
            print(f"  {p['verdict']:5} {p['name']}: {p['reason']}")
    print(f"\nReport saved: {out}")
