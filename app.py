"""
Automation Tool web app: upload a spreadsheet of companies, get back their latest turnover
from Companies House.

Run locally:   streamlit run app.py
Deploy:        see README.md
"""

import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

import turnover
from inputs import NONE, guess, guess_header_row, read_upload
from run_queue import QUEUE

HERE = Path(__file__).parent
SECONDS_PER_COMPANY = 5
DEFAULT_KEYWORDS = (HERE / "default_keywords.txt").read_text(encoding="utf-8").strip() \
    if (HERE / "default_keywords.txt").exists() else ""
# Status colours on the dark page: (background, text). The Excel file keeps its light fills.
STATUS_COLOURS = {"FOUND": ("#0f2f26", "#34d399"), "VERIFY": ("#33290d", "#fbbf24"),
                  "NOT DISCLOSED": ("#241b3d", "#c4b5fd"), "CHECK PDF": ("#36200f", "#fb923c"),
                  "NOT FOUND": ("#3a1520", "#fb7185"), "ERROR": ("#1f2430", "#94a3b8")}

st.set_page_config(page_title="Automation Tool", page_icon="📊", layout="centered")
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
  html, body, .stApp {font-family: 'Inter', sans-serif;}
  .stApp h1, .stApp h2, .stApp h3, .stApp p, .stApp label, .stApp input, .stApp textarea,
  .stApp button p, .stApp li {font-family: 'Inter', sans-serif;}
  .stApp {background:
      radial-gradient(ellipse 70% 55% at 18% -5%, rgba(124, 58, 237, .38), transparent 60%),
      radial-gradient(ellipse 45% 40% at 95% 35%, rgba(168, 85, 247, .16), transparent 60%),
      radial-gradient(ellipse 60% 40% at 50% 100%, rgba(91, 33, 182, .14), transparent 60%),
      #07040d;}
  [data-testid="stHeader"] {background: transparent;}
  [data-testid="stSidebar"] {background: rgba(13, 8, 23, .94); border-right: 1px solid rgba(255,255,255,.06);}
  .block-container {padding-top: 3.2rem; max-width: 1120px;}

  /* hero */
  .hero {margin: .5rem 0 2.2rem 0;}
  .eyebrow {color: #a78bfa; font-size: .78rem; font-weight: 600; letter-spacing: .32em;
            text-transform: uppercase; margin-bottom: 1rem;}
  .hero h1 {font-size: 3.4rem; line-height: 1.04; font-weight: 800; letter-spacing: -.035em; margin: 0;
            padding: 0; color: #fff;}
  .grad {background: linear-gradient(90deg, #f5f3ff 0%, #d8b4fe 45%, #c084fc 100%);
         -webkit-background-clip: text; background-clip: text; color: transparent;}
  .hero p.lede {color: #b8afc9; font-size: 1.08rem; line-height: 1.65; margin: 1.3rem 0 1.4rem 0;
                max-width: 34rem;}
  .trust {display: flex; flex-wrap: wrap; gap: 1.4rem; color: #c9c1d9; font-size: .86rem;}
  .trust span {display: inline-flex; align-items: center; gap: .45rem;}
  .trust svg {width: 17px; height: 17px; stroke: #a78bfa; fill: none; stroke-width: 2;}
  @media (max-width: 860px) {.hero h1 {font-size: 2.5rem;}}

  /* steps */
  .step {display: flex; align-items: center; gap: .7rem; margin: 2rem 0 .7rem 0;}
  .step .n {background: linear-gradient(135deg, #7c3aed, #c084fc); color: #fff; border-radius: 50%;
            width: 1.8rem; height: 1.8rem; display: inline-flex; align-items: center; justify-content: center;
            font-weight: 700; font-size: .9rem; flex: none; box-shadow: 0 0 18px rgba(168,85,247,.45);}
  .step .t {font-size: 1.25rem; font-weight: 700; color: #f5f3ff; letter-spacing: -.01em;}

  /* inputs and buttons */
  [data-testid="stFileUploaderDropzone"] {background: rgba(255,255,255,.03);
      border: 1px dashed rgba(167,139,250,.38); border-radius: 16px;}
  .stTextArea textarea {background: rgba(255,255,255,.03) !important; border-radius: 14px !important;}
  .stButton button[kind="primary"], [data-testid="stDownloadButton"] button[kind="primary"] {
      background: linear-gradient(90deg, #6d28d9 0%, #a855f7 100%); border: none; border-radius: 12px;
      font-weight: 600; padding: .7rem 1rem; box-shadow: 0 10px 34px rgba(139, 92, 246, .38);}
  .stButton button[kind="primary"]:hover, [data-testid="stDownloadButton"] button[kind="primary"]:hover {
      filter: brightness(1.1); box-shadow: 0 12px 40px rgba(168, 85, 247, .5);}
  .stButton button[kind="secondary"], [data-testid="stDownloadButton"] button[kind="secondary"] {
      background: rgba(255,255,255,.03); border: 1px solid rgba(255,255,255,.14); border-radius: 12px;}
  [data-testid="stDataFrame"] {border: 1px solid rgba(255,255,255,.07); border-radius: 14px; overflow: hidden;}
  [data-testid="stProgress"] > div > div > div > div {background: linear-gradient(90deg, #6d28d9, #c084fc);}

  /* results and legend */
  .cards {display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: .7rem;
          margin: .8rem 0 1.1rem 0;}
  .card {border-radius: 14px; padding: .85rem 1rem; border: 1px solid rgba(255,255,255,.07);}
  .card .v {font-size: 1.9rem; font-weight: 800; line-height: 1.1;}
  .card .l {font-size: .72rem; font-weight: 600; letter-spacing: .12em; text-transform: uppercase; opacity: .9;}
  .legend div {margin-bottom: .7rem; font-size: .83rem; line-height: 1.4; color: #a69cba;}
  .legend b {display: inline-block; padding: .08rem .5rem; border-radius: 6px; font-size: .7rem;
             letter-spacing: .06em; margin-bottom: .2rem;}
</style>
""", unsafe_allow_html=True)


def step(n, title):
    st.markdown(f'<div class="step"><span class="n">{n}</span><span class="t">{title}</span></div>',
                unsafe_allow_html=True)


def chip(status):
    bg, fg = STATUS_COLOURS[status]
    return f'<b style="background:{bg};color:{fg}">{status}</b>'


ICONS = {  # small outline icons for the trust row
    "doc": '<svg viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/>'
           '<path d="M14 3v5h5M9 13h6M9 17h6"/></svg>',
    "check": '<svg viewBox="0 0 24 24"><path d="M12 3l7 3v6c0 4.5-3 7.7-7 9-4-1.3-7-4.5-7-9V6z"/>'
             '<path d="M9 12l2 2 4-4"/></svg>',
    "scan": '<svg viewBox="0 0 24 24"><path d="M4 8V5a1 1 0 0 1 1-1h3M16 4h3a1 1 0 0 1 1 1v3M20 16v3a1 1 0 0 1-1 '
            '1h-3M8 20H5a1 1 0 0 1-1-1v-3M7 12h10"/></svg>',
    "search": '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>',
}

def hero():
    st.markdown(f"""
<div class="hero">
  <div>
    <div class="eyebrow">Companies House accounts</div>
    <h1><span class="grad">Automation</span> Tool</h1>
    <p class="lede">Upload a list of UK companies. Get back each one's turnover, cash, debtors and people
      with significant control from Companies House, and which of your keywords their accounts mention.</p>
    <div class="trust">
      <span>{ICONS["doc"]}Companies House data</span>
      <span>{ICONS["check"]}Exact tagged figures</span>
      <span>{ICONS["scan"]}Scanned PDFs read by OCR</span>
      <span>{ICONS["search"]}Keyword search</span>
    </div>
  </div>
</div>""", unsafe_allow_html=True)


def preview(ws, header_row, rows=5):
    heads = [str(c.value) if c.value is not None else f"Column {c.column_letter}" for c in ws[header_row]]
    data = [[c.value for c in ws[r]] for r in range(header_row + 1, min(ws.max_row, header_row + rows) + 1)]
    width = len(heads)
    return pd.DataFrame([(d + [None] * width)[:width] for d in data], columns=heads).astype(str).replace("None", "")


TABLE_ORDER = ["Company", "Status", "Turnover", "Cash", "Debtors", "PSC", "Matched", "Keywords", "Notes"]


def results_frame(rows):
    df = pd.DataFrame(rows, columns=TABLE_ORDER)
    return df.style.apply(lambda col: [f"background-color:{STATUS_COLOURS.get(v, ('#161024', '#fff'))[0]};"
                                       f"color:{STATUS_COLOURS.get(v, ('#161024', '#fff'))[1]};font-weight:600"
                                       for v in col], subset=["Status"])


TABLE_COLUMNS = {"Company": st.column_config.TextColumn(width="medium"),
                 "Status": st.column_config.TextColumn(width="medium"),
                 "Turnover": st.column_config.TextColumn(width="medium"),
                 "Cash": st.column_config.TextColumn(width="small"),
                 "Debtors": st.column_config.TextColumn(width="small"),
                 "PSC": st.column_config.TextColumn(width="medium", help="Persons with significant control"),
                 "Matched": st.column_config.NumberColumn(width="small", help="Different keywords found"),
                 "Keywords": st.column_config.TextColumn(width="large"),
                 "Notes": st.column_config.TextColumn(width="large")}


def money(res, value):
    if value is None:
        return ""
    symbol = {"GBP": "£", "EUR": "€", "USD": "$", "": "£"}.get(res.currency, "")
    return f"{symbol}{value:,.0f}" + ("" if symbol else f" {res.currency}")


def table_row(name, res):
    return {"Company": res.company_name or name, "Status": res.status, "Turnover": money(res, res.turnover),
            "Cash": money(res, res.cash), "Debtors": money(res, res.debtors), "PSC": turnover.psc_cell(res) or "",
            "Matched": turnover.keyword_count(res), "Keywords": turnover.keyword_cell(res) or "", "Notes": res.note}


def ordinal(n):
    return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def mentions_keywords(row):
    return bool(row["Keywords"]) and row["Keywords"] != "none found"


# ---------------------------------------------------------------- page

hero()

with st.sidebar:
    st.header("Settings")
    read_scanned = st.toggle("Read scanned PDFs", value=True,
                             help="Reads scanned PDFs with OCR to find the turnover. When off these rows are marked CHECK PDF")
    ocr_keywords = st.toggle("Search scanned PDFs for keywords", value=True, disabled=not read_scanned,
                             help="Reads every page of scanned PDFs to find keywords. Takes about 3 seconds per page")
    max_pages = st.slider("Max pages per PDF (0 = no limit)", 0, 500, 0, step=10, disabled=not read_scanned,
                          help="Most pages read from each PDF. Set to 0 to read every page")
    time_limit = st.number_input("Time limit per company (minutes)", min_value=0, max_value=240, value=0, step=1,
                                 disabled=not read_scanned,
                                 help="Most minutes spent reading each company's accounts. Set to 0 for no limit")
    if read_scanned and not turnover.ocr_available():
        st.warning("The scanned-PDF reader isn't working on this server, so scanned PDFs will be marked "
                   "CHECK PDF. Everything else works. It needs Python 3.12 and the packages.txt file; "
                   "the app's logs give the reason.")
    st.divider()
    st.subheader("What the results mean")
    st.markdown('<div class="legend">' + "".join(
        f"<div>{chip(s)}<br>{turnover.MEANINGS[s]}</div>" for s in turnover.STATUSES) + "</div>",
        unsafe_allow_html=True)

step(1, "Upload your spreadsheet and keywords")
up_col, kw_col = st.columns(2, gap="medium")
with up_col:
    upload = st.file_uploader("Companies (Excel or CSV, one per row)", type=["xlsx", "xlsm", "xls", "csv"])
    st.caption("Needs a column of company names, company numbers, or both. Numbers give exact "
               "matches. Nothing you upload is stored.")
    sample = HERE / "sample_companies.xlsx"
    if sample.exists() and not upload:
        st.download_button("Download a sample file", sample.read_bytes(), "sample_companies.xlsx",
                           width="stretch")
with kw_col:
    kw_text = st.text_area("Keywords to look for in the accounts (optional)", value=DEFAULT_KEYWORDS,
                           height=190, help="One keyword or phrase per line. Capital letters don't matter")
    keywords = turnover.parse_keywords(kw_text)
    st.caption(f"{len(keywords)} keyword{'s' if len(keywords) != 1 else ''}. Clear the box to skip the search."
               if keywords else "No keywords: only turnover will be looked up.")
if not upload:
    st.stop()

data = upload.getvalue()
try:
    wb, values = read_upload(data, upload.name)
except Exception as e:
    st.error(f"Couldn't open that file. Is it a valid Excel or CSV file? ({type(e).__name__}: {e})")
    st.stop()

sheets = wb.sheetnames
sheet = st.selectbox("Sheet", sheets) if len(sheets) > 1 else sheets[0]
ws_values = values[sheet]
if ws_values.max_row < 1 or not any(c.value not in (None, "") for row in ws_values.iter_rows(max_row=10) for c in row):
    st.error("That sheet looks empty.")
    st.stop()

step(2, "Check the columns")
header_row = st.number_input("Headers are in row", 1, max(ws_values.max_row, 1),
                             guess_header_row(ws_values), key=f"{upload.name}-{sheet}-hdr")
headers = [str(c.value).strip() for c in ws_values[header_row] if c.value not in (None, "")]
if not headers:
    st.error(f"Row {header_row} is empty. Pick the row that has your column headings.")
    st.stop()
st.dataframe(preview(ws_values, header_row), hide_index=True, width="stretch")

options = [NONE] + headers
num_guess = guess(headers, ["company no", "company number", "crn", "reg no", "registration", "number", "no."])
name_guess = guess(headers, ["company name", "name", "client", "entity", "company"], skip=num_guess)
left, right = st.columns(2)
name_col = left.selectbox("Company name column", options, index=options.index(name_guess),
                          key=f"{upload.name}-{sheet}-{header_row}-name")
number_col = right.selectbox("Company number column", options, index=options.index(num_guess),
                             key=f"{upload.name}-{sheet}-{header_row}-num")
picked = {k: v for k, v in (("name", name_col), ("number", number_col)) if v != NONE}
picked_cols = [turnover.find_col(ws_values, h, header_row) for h in picked.values()]
n_rows = sum(1 for r in range(header_row + 1, ws_values.max_row + 1)
             if any(ws_values.cell(row=r, column=c).value not in (None, "") for c in picked_cols))

step(3, "Run it")
if not picked:
    st.warning("Pick a company name column, a company number column, or both.")
    st.stop()
minutes = max(1, round(n_rows * SECONDS_PER_COMPANY / 60))
slow_keywords = bool(keywords) and read_scanned and ocr_keywords
st.caption(f"{n_rows} companies{f', searching for {len(keywords)} keywords' if keywords else ''}. "
           f"Expect about {minutes} minute{'s' if minutes != 1 else ''}"
           + ((f", plus up to {time_limit} minute{'s' if time_limit != 1 else ''} for each large company with a "
               "scanned annual report (the time limit you set)." if time_limit
               else ", plus 5-15 minutes for each large company with a scanned annual report, because every page "
               "is read to search for the keywords (turn this off, or set a time limit, in the sidebar)."
               if slow_keywords
               else ", plus 1-4 minutes for each large company with a scanned annual report.")
              if read_scanned else ".")
           + " Keep this tab open while it runs.")

running, waiting = QUEUE.status()
if running >= QUEUE.max_active:
    st.caption(f"The app is busy: {running} runs in progress" + (f", {waiting} waiting" if waiting else "")
               + ". You can still start; your run will wait its turn.")

if st.button("Find turnover", type="primary", width="stretch", disabled=n_rows == 0):
    ticket = QUEUE.join()
    queue_note = st.empty()

    def show_place(place, active):
        queue_note.info(f"⏳ You're {ordinal(place)} in the queue. {active} run{'s are' if active != 1 else ' is'} "
                        "in progress; yours starts automatically. Keep this tab open.")

    try:
        QUEUE.wait_turn(ticket, on_wait=show_place)
    except BaseException:
        QUEUE.leave(ticket)                          # left the page while waiting
        raise
    queue_note.empty()
    _wb, values = read_upload(data, upload.name)
    ws_values = values[sheet]
    cols = {k: turnover.find_col(ws_values, picked.get(k), header_row) for k in ("name", "number")}
    out_wb, out_ws = turnover.new_output()
    out_name = re.sub(r"\.(xlsx|xlsm|xls|csv)$", "", upload.name, flags=re.I) + "_results.xlsx"
    st.session_state.pop("result", None)

    bar = st.progress(0.0, text="Starting...")
    activity = st.empty()
    partial = st.empty()
    table = st.empty()
    seen = []

    def save(done, final=False):
        if final:
            turnover.add_summary(out_wb, counts_so_far(), upload.name, keywords)
        buf = io.BytesIO()
        out_wb.save(buf)
        st.session_state["result"] = {"bytes": buf.getvalue(), "name": out_name, "rows": list(seen),
                                      "done": done, "total": n_rows, "final": final}

    def counts_so_far():
        c = {}
        for r in seen:
            c[r["Status"]] = c.get(r["Status"], 0) + 1
            if mentions_keywords(r):
                c["_with_keywords"] = c.get("_with_keywords", 0) + 1
        return c

    def on_row(done, total, name, res):
        bar.progress(min(done / max(total, 1), 1.0), text=f"Checked {done} of {total}")
        seen.append(table_row(name, res))
        table.dataframe(results_frame(seen[-12:]), hide_index=True, width="stretch",
                        column_config=TABLE_COLUMNS)
        activity.empty()
        if done % 5 == 0 and done < total:
            save(done)                                   # partial results survive an interruption
            # Downloading these doesn't rerun the app, so the run carries on.
            partial.download_button(f"⬇  Download the {done} done so far (Excel)", st.session_state["result"]["bytes"],
                                    file_name=out_name.replace("_results.xlsx", f"_first_{done}.xlsx"),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    on_click="ignore", width="stretch", key=f"partial-{done}")

    try:
        turnover.process_sheet(ws_values, cols, out_ws, time_limit=time_limit, max_pages=max_pages or None,
                               on_row=on_row, header_row=header_row, read_scanned=read_scanned,
                               keywords=keywords or None, ocr_keywords=ocr_keywords,
                               on_detail=lambda msg: activity.caption(f"⏳ {msg}"))
    finally:
        QUEUE.leave(ticket)                          # finished, stopped or left: next in line goes
    save(len(seen), final=True)
    bar.empty()
    partial.empty()
    table.empty()

res = st.session_state.get("result")
if res:
    counts = {}
    for r in res["rows"]:
        counts[r["Status"]] = counts.get(r["Status"], 0) + 1
    if not res["final"]:
        st.warning(f"The last run stopped after {res['done']} of {res['total']} companies. "
                   "You can download what was finished, or run it again.")
    cards = "".join(
        f'<div class="card" style="background:{STATUS_COLOURS[s][0]};color:{STATUS_COLOURS[s][1]}">'
        f'<div class="v">{counts.get(s, 0)}</div><div class="l">{s.title()}</div></div>'
        for s in turnover.STATUSES if counts.get(s) or s in ("FOUND", "NOT DISCLOSED"))
    st.markdown(f'<div class="cards">{cards}</div>', unsafe_allow_html=True)
    st.download_button("⬇  Download results (Excel)", res["bytes"], file_name=res["name"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", width="stretch")
    st.dataframe(results_frame(res["rows"]), hide_index=True, width="stretch",
                 height=min(38 + 35 * len(res["rows"]), 420), column_config=TABLE_COLUMNS)
    with_kw = sum(1 for r in res["rows"] if mentions_keywords(r))
    st.caption("The Excel file has one row per company: turnover, prior year, cash, debtors, persons with "
               "significant control, a LinkedIn search link, keywords, and links to each PDF and company page, "
               "plus a Summary sheet."
               + (f" {with_kw} of {len(res['rows'])} companies mention at least one keyword."
                  if any(r["Keywords"] for r in res["rows"]) else ""))
