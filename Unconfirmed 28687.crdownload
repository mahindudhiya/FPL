"""
FPL Weekly Advisor: web version.
Pulls live data from the FPL API every time you open or refresh the page.
Scoring logic lives in fpl_weekly.py (same folder).

Bookmark your link with your team ID so it loads straight away:
    https://YOUR-APP.streamlit.app/?team=1234567
"""

import os
import tempfile
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import fpl_weekly as fw

fw.CACHE = os.path.join(tempfile.gettempdir(), "fpl_cache")
LOCAL_TZ = ZoneInfo("America/Toronto")

FDR_CSS = {
    0: "background-color:#000000;color:#FFFFFF;font-weight:bold",
    1: "background-color:#1E7B34;color:#FFFFFF",
    2: "background-color:#A9D18E;color:#000000",
    3: "background-color:#E7E6E6;color:#000000",
    4: "background-color:#F4B183;color:#000000",
    5: "background-color:#C00000;color:#FFFFFF",
}
VERDICT_CSS = {
    "SELL": "background-color:#F8CBAD;font-weight:bold",
    "WATCH": "background-color:#FFE699;font-weight:bold",
    "HOLD": "background-color:#C6EFCE;font-weight:bold",
}


# ----------------------------- data helpers -----------------------------
def parse_pending(text):
    """'João Pedro > N.Jackson' per line -> [(out, in)]"""
    moves = []
    for line in (text or "").replace(";", "\n").splitlines():
        if ">" in line:
            out_name, in_name = (s.strip() for s in line.split(">", 1))
            if out_name and in_name:
                moves.append((out_name, in_name))
    return moves


@st.cache_resource(ttl=900, show_spinner=False)
def run(team_id, lookahead, weights, pending):
    """Cached for 15 minutes; 'Refresh data now' clears it."""
    fw.LOOKAHEAD = lookahead
    fw.WEIGHTS = dict(weights)
    fw.PENDING_TRANSFERS = list(pending)
    return fw.analyse(team_id)


def fixtures_for(r, team_id):
    texts, fdrs = [], []
    for g in r["gws"]:
        fx = r["upcoming"][team_id][g]
        if not fx:
            texts.append("BLANK")
            fdrs.append(0)
        else:
            texts.append(" + ".join(t for t, _ in fx))
            fdrs.append(round(sum(d for _, d in fx) / len(fx)))
    return texts, fdrs


def player_table(r, players, extra):
    gw_cols = [f"GW{g}" for g in r["gws"]]
    rows, fdr_rows = [], []
    for p in players:
        texts, fdrs = fixtures_for(r, p["team"])
        row = {"Pos": fw.POS[p["pos"]], "Player": p["name"], "Club": p["club"], "Price": p["price"],
               "Fit %": int(p["avail"] * 100), "Form": p["form"], "xGI/90": round(p["xgi90"], 2),
               "Starts %": int(p["share"] * 100), "Score": p["score"]}
        row.update(extra(p))
        row.update(dict(zip(gw_cols, texts)))
        rows.append(row)
        fdr_rows.append(dict(zip(gw_cols, fdrs)))
    return pd.DataFrame(rows), pd.DataFrame(fdr_rows), gw_cols


def styled(df, fdr_df, gw_cols, verdict_col=None):
    def css(_):
        out = pd.DataFrame("", index=df.index, columns=df.columns)
        for c in gw_cols:
            out[c] = fdr_df[c].map(FDR_CSS)
        if verdict_col:
            out[verdict_col] = df[verdict_col].map(VERDICT_CSS).fillna("")
        return out
    return df.style.apply(css, axis=None).format(precision=1)


def ticker_table(r):
    gw_cols = [f"GW{g}" for g in r["gws"]]
    rows, fdr_rows = [], []
    for i, t in enumerate(sorted(r["teams"], key=lambda t: -r["ease"][t]), 1):
        texts, fdrs = fixtures_for(r, t)
        rows.append({"Rank": i, "Team": r["teams"][t]["name"], "Ease (0-100)": round(r["ease"][t]),
                     **dict(zip(gw_cols, texts))})
        fdr_rows.append(dict(zip(gw_cols, fdrs)))
    return pd.DataFrame(rows), pd.DataFrame(fdr_rows), gw_cols


def excel_bytes(r):
    path = os.path.join(tempfile.gettempdir(), f"FPL_Weekly_Report_GW{r['next_gw']}.xlsx")
    fw.write_report(r, path)
    with open(path, "rb") as fh:
        return fh.read()


# ----------------------------- page -----------------------------
def main():
    st.set_page_config(page_title="FPL Weekly Advisor", page_icon="⚽", layout="wide")
    qp = st.query_params

    with st.sidebar:
        st.header("Settings")
        team = st.text_input("FPL team ID", value=qp.get("team", ""),
                             help="The number in your Points page URL: .../entry/1234567/event/9")
        pending_text = st.text_area("Transfers made this week (one per line: Out > In)",
                                    value=qp.get("pending", "").replace(";", "\n"),
                                    help="FPL only shows your transfers after the deadline. Example: João Pedro > N.Jackson")
        lookahead = st.slider("Fixture lookahead (gameweeks)", 3, 8, 5)
        with st.expander("Score weights"):
            w_form = st.slider("Form", 0, 100, 30)
            w_under = st.slider("Underlying stats (xG/xA, xGC)", 0, 100, 25)
            w_fix = st.slider("Fixtures", 0, 100, 25)
            w_ep = st.slider("FPL expected points", 0, 100, 20)
        if st.button("Refresh data now", use_container_width=True):
            run.clear()
        st.caption("Data refreshes automatically every 15 minutes.")

    if not team.strip().isdigit():
        st.title("FPL Weekly Advisor")
        st.info("Enter your FPL team ID in the sidebar to start. Tip: bookmark this page with ?team=YOURID on the end.")
        return

    total = (w_form + w_under + w_fix + w_ep) or 1
    weights = (("form", w_form / total), ("underlying", w_under / total),
               ("fixtures", w_fix / total), ("ep_next", w_ep / total))

    try:
        with st.spinner("Pulling live FPL data..."):
            r = run(int(team), lookahead, weights, tuple(parse_pending(pending_text)))
    except SystemExit as exc:
        st.error(str(exc))
        return
    except Exception as exc:
        st.error(f"Couldn't read data for team {team}. Check the ID, or try again if FPL is updating. ({exc})")
        return

    deadline = r["deadline"].astimezone(LOCAL_TZ)
    st.title(f"GW{r['next_gw']} Advisor: {r['entry'].get('name', '')}")
    for w in r["warnings"]:
        st.warning(w)

    top = r["moves"][0] if r["moves"] else None
    make_move = bool(top and top[0] >= fw.MIN_GAIN)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Deadline (Toronto)", f"{deadline:%a %b %d, %I:%M %p}")
    c2.metric("Bank", f"£{r['bank']:.1f}m")
    c3.metric("Top move", f"{top[1]['name']} → {top[2]['name']}" if make_move else "Roll transfer",
              f"+{top[0]}" if make_move else None)
    c4.metric("Captain", r["captains"][0]["name"] if r["captains"] else "-")

    # 1. transfers
    st.subheader("1. Recommended transfers")
    if make_move:
        st.success(f"Best move: **{top[1]['name']} → {top[2]['name']}** (+{top[0]} score).")
    else:
        st.info(f"Roll your transfer: no upgrade beats your current player by {fw.MIN_GAIN}+ points.")
    if r["moves"]:
        st.dataframe(pd.DataFrame([{
            "Gain": g, "Sell": o["name"], "Sell club": o["club"], "Sell price": o["price"], "Sell score": o["score"],
            "Buy": i["name"], "Buy club": i["club"], "Buy price": i["price"], "Buy score": i["score"],
            "Seller verdict": o["verdict"]} for g, o, i in r["moves"]]).style.format(precision=1),
            hide_index=True, use_container_width=True)
    st.caption("Uses current prices. Your selling price can be lower if a player rose after you bought him.")

    # 2. captain
    st.subheader("2. Captain options")
    st.dataframe(pd.DataFrame([{
        "Rank": n, "Player": p["name"], "Club": p["club"], "Expected pts": round(p["ep"] * p["avail"], 1),
        "Next fixture": fixtures_for(r, p["team"])[0][0]} for n, p in enumerate(r["captains"], 1)]),
        hide_index=True, use_container_width=True)

    # 3. squad
    st.subheader("3. My squad")
    df, fdr, gw_cols = player_table(r, r["squad"], lambda p: {"Verdict": p["verdict"], "Reason": p["reason"]})
    st.dataframe(styled(df, fdr, gw_cols, "Verdict"), hide_index=True, use_container_width=True)

    # 4. targets
    st.subheader("4. Top targets (fit, starting, not in your team)")
    picks = [p for pos in fw.POS for p in r["targets"][pos]]
    df, fdr, gw_cols = player_table(r, picks, lambda p: {"Owned %": p["owned_pct"]})
    positions = st.multiselect("Positions", ["GK", "DEF", "MID", "FWD"], default=["DEF", "MID", "FWD"])
    max_price = st.slider("Max price (£m)", 4.0, 16.0, 16.0, 0.5)
    mask = df["Pos"].isin(positions) & (df["Price"] <= max_price)
    st.dataframe(styled(df[mask].reset_index(drop=True), fdr[mask].reset_index(drop=True), gw_cols),
                 hide_index=True, use_container_width=True)

    # 5. fixtures
    st.subheader(f"5. Fixture ticker (next {len(r['gws'])} gameweeks, easiest first)")
    df, fdr, gw_cols = ticker_table(r)
    st.dataframe(styled(df, fdr, gw_cols), hide_index=True, use_container_width=True, height=740)

    st.download_button("Download this as Excel", excel_bytes(r),
                       file_name=f"FPL_Weekly_Report_GW{r['next_gw']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with st.expander("How the score works"):
        st.markdown(
            "Each player gets a **0 to 100 score**: form (recent points and points per game), underlying stats "
            "(xG + xA per 90 for attackers, expected goals conceded for defenders and keepers), official FPL fixture "
            "difficulty over the lookahead window, and FPL's expected points for next week. Each part is ranked "
            "against players in the same position, weighted by the sliders, then multiplied by chance of playing "
            "and how regularly the player starts.\n\n"
            f"**SELL** below {fw.SELL_BELOW} or 25% or less to play. **WATCH** below {fw.WATCH_BELOW} or any "
            "fitness flag. **HOLD** otherwise. Data: Fantasy Premier League public API. "
            "It can't read press conferences, so check team news before the deadline.")


if __name__ == "__main__":
    main()
