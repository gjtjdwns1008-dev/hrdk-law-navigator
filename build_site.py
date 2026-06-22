# -*- coding: utf-8 -*-
"""
build_site.py — 자격증 법령 네비게이터 (통합 1파일 빌더)
================================================================
두 화면을 '한 개의 index.html'로 굽는다. 상단 탭을 누르면 페이지 이동 없이
보이는 화면만 바뀐다(SPA식).
  · 탭1 「법령 제개정에 따른 자격증 활용도 모니터링」 — 법령 카드(월별 기간선택 + 검색)
  · 탭2 「자격증별 채용시장 우대사항 모니터링」      — 자격증 카드(빈도순 배지 + 2차 팝업)
서버 불필요(GitHub Actions → Pages).

[실행]
  운영(기본):
    monitor → GCP_SA_JSON,  GOOGLE_SHEET_ID, SOURCE_WORKSHEET(기본 "연관 높은 법령")
    radar   → RADAR_SA_JSON(없으면 GCP_SA_JSON 폴백), RADAR_SHEET_ID, RADAR_WORKSHEET(기본 "우대사항_대장")
  미리보기:
    monitor → LOCAL_XLSX,        LOCAL_SHEET
    radar   → LOCAL_XLSX_RADAR,  LOCAL_SHEET_RADAR
  옵션: M_MAX(기본 5000), R_MAX(기본 9999), OUT_DIR(기본 dist)
"""
import os, re, json, html, hashlib, datetime
from collections import defaultdict, Counter
from urllib.parse import quote

OUT_DIR = os.environ.get("OUT_DIR", "dist")
M_MAX = int(os.environ.get("M_MAX", "5000"))
R_MAX = int(os.environ.get("R_MAX", "9999"))

MCOL = {"law":"법령명","ministry":"소관부처","date":"시행일자","kind":"개정유형",
        "summary1":"활용도 분석 상세","summary2":"주요 제·개정내용",
        "certs":"법령 관련 국가기술자격 종목","article":"근거조문","link":"조문별 다이렉트 링크"}
RCOL = {"law":"법령명","article":"조문","pref":"우대분류","certs":"해당 자격종목",
        "t1type":"Track1_취급유형","t1risk":"Track1_위험도","t2":"Track2_효용코드",
        "sjb":"중처법대상","note":"비고"}
PREF_ORDER = ["의무고용","직무권한부여","인사우대","시험면제","기타"]
PREF_COLOR = {"의무고용":"#C0492F","직무권한부여":"#1F6FB2","인사우대":"#0F6E56","시험면제":"#5B4BB0","기타":"#8A8F98"}

TRACK1_TYPE = {
 "A":["신분형성형","자격 취득이 행정청 면허로 이어져 평생 직업·신분을 부여하는 유형."],
 "B":["영업요건형","사업 등록·허가·지정 시 자격자 보유가 의무인 유형."],
 "C":["직역독점형","특정 직무(선임·배치·서명·확인)를 자격자만 수행할 수 있는 유형."],
 "D":["인사가산형","공무원·근로자의 채용·승진·평정·보수에 부가로 우대되는 유형."],
 "E":["검정연계형","다른 자격·시험의 응시자격·시험면제와 연계되는 유형."]}
TRACK1_RISK = {
 "N":["무관","자격이 직역 진입 조건이 아니라 부가우대만 주는 경우."],
 "L":["저위험","자격이 진입 조건이나 학력·경력·유사자격으로 우회 가능."],
 "M":["중위험","법령이 인정하는 복수 자격 중 하나로 대체 가능(우회로 존재)."],
 "H":["고위험","자격과 경력을 동시에 요구해 경력 선행 조건이 되는 경우."],
 "C":["임계위험","단일 자격만 인정되어 대체 경로가 없는 경우."]}
TRACK2 = {
 "Ⅰ-1":["면허전환형","Ⅰ 직업창출형","자격 취득이 행정청 면허로 이어져 평생 직업·신분을 부여."],
 "Ⅰ-2":["개업창업형","Ⅰ 직업창출형","자격자 본인이 단독으로 직무를 수행·서명할 수 있어 1인 사업이 가능."],
 "Ⅱ-1":["등록필수형","Ⅱ 취업관문형","사업체 등록·허가 시 자격자를 일정 인원 이상 보유해야 하는 유형."],
 "Ⅱ-2":["지정인력형","Ⅱ 취업관문형","국가 지정·위탁·대행 기관(검사·검정·인증·진단 등)의 인력 요건."],
 "Ⅱ-3":["전속배치형","Ⅱ 취업관문형","사업장에 단일 자격자만 선임 가능(대체 불가). 매우 드문 유형."],
 "Ⅱ-4":["선택배치형","Ⅱ 취업관문형","법령이 인정하는 복수 자격 중 택일하여 선임하는 유형."],
 "Ⅱ-5":["현장배치형","Ⅱ 취업관문형","공사·사업장 규모·종별에 따라 자격자를 배치하는 유형."],
 "Ⅲ-1":["부가우대(시험면제)","Ⅲ 부가우대형","다른 자격·면허·임용시험에서 시험과목을 면제받는 유형."],
 "Ⅲ-2":["부가우대(인사)","Ⅲ 부가우대형","채용·보수·평정·승진 등 인사에서 우대받는 유형."],
 "Ⅲ-3":["부가우대(위촉·자문)","Ⅲ 부가우대형","위원회·심의위원·시험위원 등 자문성 위촉 자격."],
 "Ⅳ-0":["제외","분류 외","중복·삭제·이관 등 분류 대상에서 제외된 조항."]}


# ───────── 공통 유틸 ─────────
def _sheet_key(v):
    m = re.search(r"/d/([A-Za-z0-9_-]+)", str(v or "")); return m.group(1) if m else str(v or "").strip()
def _client(raw): 
    import gspread; return gspread.service_account_from_dict(json.loads(raw.strip(), strict=False))
def digits(v): return re.sub(r"\D", "", str(v or ""))[:8]
def fmt_date(v):
    d = digits(v); return f"{d[:4]}.{d[4:6]}.{d[6:]}" if len(d) == 8 else str(v or "")
def esc(v): return html.escape(str(v or "").strip())
def law_url_name(name): return f"https://www.law.go.kr/법령/{quote(str(name or '').strip())}"
def tok(v): return str(v or "").split(" ")[0].strip()  # "B (영업요건형)" -> "B"


# ───────── 로드 ─────────
def load_monitor():
    lx = os.environ.get("LOCAL_XLSX", "").strip()
    ws = os.environ.get("SOURCE_WORKSHEET", "연관 높은 법령")
    if lx:
        import pandas as pd
        sh = os.environ.get("LOCAL_SHEET", ws)
        return pd.read_excel(lx, sheet_name=sh).fillna("").to_dict("records")
    gc = _client(os.environ["GCP_SA_JSON"])
    return gc.open_by_key(_sheet_key(os.environ["GOOGLE_SHEET_ID"])).worksheet(ws).get_all_records()

def _nospace(s): return re.sub(r"\s+", "", str(s or ""))
def _detail_map(rows):
    """국가기술자격 관련법령 탭 → {법령명(공백제거): 상세 분석 결과} 매핑."""
    m = {}
    for r in rows:
        ln = _nospace(r.get("법령명"))
        d = str(r.get("상세 분석 결과") or "").strip()
        if ln and d and ln not in m:
            m[ln] = d
    return m

def load_radar():
    """(우대사항_대장 행들, 상세분석 맵) 반환. 대장=뼈대, 관련법령=상세분석 보강."""
    ws  = os.environ.get("RADAR_WORKSHEET", "우대사항_대장")
    dws = os.environ.get("RADAR_DETAIL_WORKSHEET", "국가기술자격 관련법령")
    lx = os.environ.get("LOCAL_XLSX_RADAR", "").strip()
    if lx:
        import pandas as pd
        sh = os.environ.get("LOCAL_SHEET_RADAR", ws)
        led = pd.read_excel(lx, sheet_name=sh).fillna("").to_dict("records")
        try:
            det = pd.read_excel(lx, sheet_name=dws).fillna("").to_dict("records")
        except Exception:
            det = []
        return led, _detail_map(det)
    raw = os.environ.get("RADAR_SA_JSON", "").strip() or os.environ["GCP_SA_JSON"].strip()
    ss = _client(raw).open_by_key(_sheet_key(os.environ["RADAR_SHEET_ID"]))
    led = ss.worksheet(ws).get_all_records()
    try:
        det = ss.worksheet(dws).get_all_records()
    except Exception:
        det = []
    return led, _detail_map(det)


# ───────── monitor 데이터/카드 ─────────
def m_fields(row):
    certs = [c.strip() for c in re.split(r"[,/·\n]", str(row.get(MCOL["certs"]) or "")) if c.strip()]
    arts = [a.strip() for a in re.split(r"[,\n;·]", str(row.get(MCOL["article"]) or "")) if a.strip()]
    mn = str(row.get(MCOL["ministry"]) or "").strip(); kd = str(row.get(MCOL["kind"]) or "").strip()
    dt = fmt_date(row.get(MCOL["date"]))
    return {"law":str(row.get(MCOL["law"]) or "").strip(), "month":digits(row.get(MCOL["date"]))[:6],
            "meta":" · ".join(x for x in [mn,dt,kd] if x), "certs":certs,
            "summary_main":str(row.get(MCOL["summary2"]) or "").strip(),
            "summary_use":str(row.get(MCOL["summary1"]) or "").strip(),
            "articles":arts, "url":(str(row.get(MCOL["link"]) or "").strip() if str(row.get(MCOL["link"]) or "").startswith("http") else law_url_name(row.get(MCOL["law"])))}

def m_card(d, i):
    shown = d["certs"][:4]; extra = len(d["certs"]) - len(shown)
    chips = "".join(f'<span class="chip">{esc(c)}</span>' for c in shown) + (f'<span class="chip chip-more">+{extra}</span>' if extra>0 else "")
    summ = esc(d["summary_use"] or d["summary_main"] or "요약 준비 중입니다.")
    return f"""
    <article class="card" data-i="{i}" data-month="{d['month']}">
      <div class="card-head">{esc(d['meta'])}</div>
      <h3 class="card-title"><button type="button" class="title-btn">{esc(d['law'])}</button></h3>
      <div class="chips">{chips}</div>
      <p class="summary">{summ}</p>
      <div class="card-foot"><button type="button" class="detail-link">분석 상세 보기 →</button>
        <a class="ext" href="{esc(d['url'])}" target="_blank" rel="noopener">법제처 원문</a></div>
    </article>"""


# ───────── radar 데이터/카드 ─────────
def r_pref_idx(p): return PREF_ORDER.index(p) if p in PREF_ORDER else len(PREF_ORDER)

def r_build(rows, detail_map):
    entries = []                 # 고유 우대조항(법령·조문 단위)
    cert_map = defaultdict(list) # 자격증 -> entries 인덱스 참조
    for r in rows:
        certs = [c.strip() for c in re.split(r"[,/·\n]", str(r.get(RCOL["certs"]) or "")) if c.strip()]
        if not certs: continue
        law = str(r.get(RCOL["law"]) or "").strip()
        sjb = str(r.get(RCOL["sjb"]) or "").strip() not in ("","비대상","해당없음")
        e = {"law":law, "a":str(r.get(RCOL["article"]) or "").strip(),
             "p":str(r.get(RCOL["pref"]) or "").strip() or "기타",
             "t1":tok(r.get(RCOL["t1type"])), "t1r":tok(r.get(RCOL["t1risk"])), "t2":tok(r.get(RCOL["t2"])),
             "s":1 if sjb else 0}
        det = detail_map.get(_nospace(law), "")   # 관련법령 탭의 상세 분석 결과
        if det: e["d"] = det
        ei = len(entries); entries.append(e)
        for c in certs: cert_map[c].append(ei)
    items = sorted(cert_map.items(), key=lambda kv: len({entries[ei]["law"] for ei in kv[1]}), reverse=True)[:R_MAX]
    certs_out = []
    for cert, idxs in items:
        prefs = [p for p,_ in Counter(entries[ei]["p"] for ei in idxs).most_common()]
        idxs_sorted = sorted(idxs, key=lambda ei:(r_pref_idx(entries[ei]["p"]), entries[ei]["law"]))
        certs_out.append({"cert":cert, "prefs":prefs,
                          "law_count":len({entries[ei]["law"] for ei in idxs}),
                          "sjb":any(entries[ei]["s"] for ei in idxs), "idx":idxs_sorted})
    return certs_out, entries, len(cert_map)

def r_card(d, i):
    badges = "".join(f'<span class="pf" style="--c:{PREF_COLOR.get(p,"#8A8F98")}">{esc(p)}</span>' for p in d["prefs"])
    sjb = '<span class="sjb-badge">⚠ 중대재해처벌법 관련</span>' if d["sjb"] else ""
    return f"""
    <article class="card rcard" data-i="{i}">
      <h3 class="cert"><button type="button" class="title-btn">{esc(d['cert'])}</button></h3>
      <div class="pfs">{badges}</div>
      <div class="card-foot">
        <div class="foot-meta"><span class="lc">우대 법령 {d['law_count']}개</span>{sjb}</div>
        <div class="foot-action"><button type="button" class="detail-link">우대 근거 상세 →</button></div>
      </div>
    </article>"""


# ───────── 조립 ─────────
def build():
    # monitor
    mrows = [r for r in load_monitor() if len(digits(r.get(MCOL["date"]))) == 8]
    mrows.sort(key=lambda r: digits(r.get(MCOL["date"])), reverse=True)
    mrows = mrows[:M_MAX]
    mdata = [m_fields(r) for r in mrows]
    months = sorted({d["month"] for d in mdata if d["month"]})
    def_to = months[-1] if months else ""; def_from = months[-2] if len(months)>=2 else def_to
    m_total_certs = len({c for d in mdata for c in d["certs"]})
    m_opts = "".join(f'<option value="{m}">{m[:4]}.{m[4:6]}</option>' for m in reversed(months))
    m_cards = "\n".join(m_card(d,i) for i,d in enumerate(mdata)) or '<p class="empty">표시할 법령이 없습니다.</p>'

    # radar
    rrows, rdetail = load_radar()
    rcerts, rentries, r_total = r_build(rrows, rdetail)
    r_cards = "\n".join(r_card(d,i) for i,d in enumerate(rcerts)) or '<p class="empty">자료가 없습니다.</p>'

    out = PAGE
    repl = {
      "@@M_OPTS@@":m_opts, "@@M_DEF_FROM@@":def_from, "@@M_DEF_TO@@":def_to,
      "@@M_TOTAL_CERTS@@":str(m_total_certs), "@@M_CARDS@@":m_cards,
      "@@R_CARDS@@":r_cards, "@@R_TOTAL@@":str(r_total),
      "@@BUILT_AT@@":datetime.datetime.now().strftime("%Y.%m.%d"),
      "@@MLAWS@@":json.dumps(mdata, ensure_ascii=False).replace("</","<\\/"),
      "@@RCERTS@@":json.dumps(rcerts, ensure_ascii=False).replace("</","<\\/"),
      "@@RENTRIES@@":json.dumps(rentries, ensure_ascii=False).replace("</","<\\/"),
      "@@T1TYPE@@":json.dumps(TRACK1_TYPE, ensure_ascii=False),
      "@@T1RISK@@":json.dumps(TRACK1_RISK, ensure_ascii=False),
      "@@T2@@":json.dumps(TRACK2, ensure_ascii=False),
      "@@PFC@@":json.dumps(PREF_COLOR, ensure_ascii=False),
    }
    for k,v in repl.items(): out = out.replace(k,v)
    return out, len(mdata), len(rcerts), r_total

def main():
    out, nm, nr, total = build()
    os.makedirs(OUT_DIR, exist_ok=True)
    p = os.path.join(OUT_DIR, "index.html")
    open(p,"w",encoding="utf-8").write(out)
    print(f"✅ 생성: {p}  (법령 {nm}건 / 자격증 {nr}개[전체 {total}])")


PAGE = r"""<!DOCTYPE html>
<html lang="ko"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>자격증 법령 네비게이터 · HRDK</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
  :root{--navy:#1F3864;--ink:#16243F;--body:#33394A;--muted:#6B7280;--line:#E4E7EC;--bg:#F6F8FB;--surface:#fff;--accent:#0F6E56;}
  *{box-sizing:border-box;} html{scroll-behavior:smooth;}
  body{margin:0;background:var(--bg);color:var(--body);font-family:"Pretendard",-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;font-size:16px;line-height:1.65;-webkit-font-smoothing:antialiased;}
  a{color:inherit;} .wrap{max-width:1080px;margin:0 auto;padding:0 20px;}
  [hidden]{display:none !important;}
  .gov-bar{background:var(--navy);color:#fff;font-size:13px;} .gov-bar .wrap{padding:7px 20px;display:flex;gap:8px;align-items:center;}
  .gov-bar b{font-weight:600;} .gov-bar span{opacity:.8;}
  header.site{background:var(--surface);border-bottom:1px solid var(--line);} header.site .wrap{padding:16px 20px 0;}
  .logo{font-size:19px;font-weight:700;color:var(--ink);} .logo em{color:var(--navy);font-style:normal;}
  .tabs{display:flex;gap:4px;margin-top:14px;flex-wrap:wrap;}
  .tab{padding:11px 16px;font-size:14px;font-weight:600;color:var(--muted);background:none;border:none;border-bottom:3px solid transparent;border-radius:8px 8px 0 0;cursor:pointer;font-family:inherit;}
  .tab:hover{color:var(--ink);background:#F0F3F7;} .tab.active{color:var(--navy);border-bottom-color:var(--navy);}
  .hero{background:var(--surface);border-bottom:1px solid var(--line);} .hero .wrap{padding:32px 20px 26px;}
  .eyebrow{font-size:13px;font-weight:600;color:var(--accent);}
  .hero h1{margin:.4em 0 .1em;font-size:clamp(23px,3.5vw,33px);line-height:1.25;font-weight:700;color:var(--ink);letter-spacing:-.02em;max-width:20em;}
  .hero h1 strong{color:var(--navy);}
  .hero p.lead{margin:8px 0 0;color:var(--muted);font-size:15px;max-width:40em;}
  .note{margin-top:14px;display:inline-block;font-size:12.5px;color:#8A5A00;background:#FBF3E2;border:1px solid #F0E0BC;border-radius:8px;padding:7px 12px;}
  .stats{display:flex;gap:34px;margin-top:22px;flex-wrap:wrap;}
  .stat .n{font-size:28px;font-weight:700;color:var(--navy);font-variant-numeric:tabular-nums;line-height:1;} .stat .l{font-size:13px;color:var(--muted);margin-top:5px;}
  .toolbar{position:sticky;top:0;z-index:5;background:rgba(246,248,251,.93);backdrop-filter:saturate(160%) blur(6px);border-bottom:1px solid var(--line);}
  .toolbar .wrap{padding:14px 20px;display:flex;flex-direction:column;gap:10px;}
  .trow{display:flex;gap:10px;align-items:center;flex-wrap:wrap;}
  .period{font-size:13.5px;color:var(--muted);} .period>span:first-child{margin-right:4px;}
  select{border:1px solid var(--line);border-radius:10px;background:#fff;color:var(--ink);font-size:14px;font-family:inherit;padding:9px 12px;cursor:pointer;outline:none;}
  select:focus{border-color:var(--navy);box-shadow:0 0 0 3px rgba(31,56,100,.12);}
  .count{font-size:13px;color:var(--muted);margin-left:auto;} .count b{color:var(--navy);font-variant-numeric:tabular-nums;}
  .search{position:relative;flex:1;min-width:220px;}
  .search input{width:100%;border:1px solid var(--line);border-radius:11px;padding:12px 14px 12px 42px;font-size:15px;font-family:inherit;background:#fff;color:var(--ink);outline:none;}
  .search input:focus{border-color:var(--navy);box-shadow:0 0 0 3px rgba(31,56,100,.12);}
  .search svg{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--muted);}
  main .wrap{padding:26px 20px 10px;}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:16px;align-items:stretch;}
  .grid.rgrid{grid-template-columns:repeat(auto-fill,minmax(300px,1fr));}
  .card{background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--navy);border-radius:12px;padding:18px 19px;display:flex;flex-direction:column;transition:box-shadow .18s,transform .18s;}
  .card:hover{box-shadow:0 6px 22px -8px rgba(22,36,63,.18);transform:translateY(-2px);}
  .card-head{font-size:12.5px;color:var(--muted);} .card-title{margin:9px 0 0;}
  .title-btn{all:unset;cursor:pointer;font-size:17px;font-weight:600;line-height:1.4;color:var(--ink);letter-spacing:-.01em;}
  .rcard .title-btn{font-weight:700;}
  .title-btn:hover,.title-btn:focus-visible{color:var(--navy);text-decoration:underline;text-underline-offset:3px;}
  .chips{display:flex;flex-wrap:wrap;gap:6px;margin:11px 0 12px;}
  .chip{font-size:12px;background:#EEF2F8;color:#3A4862;border-radius:6px;padding:3px 8px;} .chip-more{background:transparent;color:var(--muted);}
  .summary{margin:0;font-size:14.5px;color:var(--body);line-height:1.62;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;}
  .card-foot{margin-top:auto;padding-top:14px;display:flex;align-items:center;justify-content:space-between;gap:10px;}
  .detail-link{all:unset;cursor:pointer;font-size:13.5px;font-weight:600;color:var(--accent);white-space:nowrap;} .detail-link:hover{text-decoration:underline;text-underline-offset:3px;}
  .card-foot .ext{font-size:12.5px;color:var(--muted);text-decoration:none;} .card-foot .ext:hover{color:var(--navy);text-decoration:underline;}
  .empty,.noresult{color:var(--muted);text-align:center;padding:40px 0;} .noresult{display:none;}
  /* radar 카드 */
  .cert{margin:0;}
  .pfs{display:flex;flex-wrap:wrap;gap:6px;margin:13px 0 4px;}
  .pf{font-size:12px;font-weight:600;color:var(--c);background:color-mix(in srgb,var(--c) 12%,#fff);border:1px solid color-mix(in srgb,var(--c) 30%,#fff);border-radius:999px;padding:3px 10px;}
  .lc{font-size:12.5px;color:var(--muted);} .sjb{color:#C0492F;font-weight:600;}
  .rcard .card-foot{flex-direction:column;align-items:stretch;gap:9px;}
  .foot-meta{display:flex;flex-wrap:wrap;align-items:center;gap:7px;}
  .foot-action{display:flex;justify-content:flex-end;}
  .sjb-badge{white-space:nowrap;font-size:11.5px;font-weight:600;color:#C0492F;background:#FBECEA;border:1px solid #F0D2CC;border-radius:6px;padding:2px 8px;}
  /* 모달 */
  .modal{position:fixed;inset:0;display:none;} .modal.open{display:block;}
  .modal-backdrop{position:absolute;inset:0;background:rgba(16,36,63,.45);}
  .modal-panel{position:relative;max-width:700px;margin:5vh auto;background:#fff;border-radius:16px;max-height:88vh;overflow:auto;padding:30px 30px 28px;box-shadow:0 24px 60px -20px rgba(16,36,63,.55);}
  .modal-close{position:absolute;top:12px;right:14px;border:none;background:transparent;font-size:26px;color:var(--muted);cursor:pointer;}
  #modal{z-index:50;} #modal2{z-index:60;} #modal2 .modal-panel{max-width:640px;}
  .m-title{margin:0 0 4px;font-size:22px;font-weight:700;color:var(--ink);line-height:1.35;padding-right:30px;}
  .m-meta{font-size:13.5px;color:var(--muted);}
  .m-sec{margin-top:22px;} .m-sec h4{margin:0 0 9px;font-size:13px;font-weight:700;color:var(--navy);padding-bottom:7px;border-bottom:1px solid var(--line);}
  .m-sec p{margin:0;font-size:15px;line-height:1.72;color:var(--body);white-space:pre-line;}
  .m-chips{display:flex;flex-wrap:wrap;gap:6px;max-height:170px;overflow:auto;padding:2px;}
  .m-arts{margin:0;padding-left:18px;} .m-arts li{font-size:14px;line-height:1.7;} .m-none{color:var(--muted);font-size:14px;}
  .m-ext{display:inline-block;margin-top:24px;background:var(--navy);color:#fff;text-decoration:none;font-size:14px;font-weight:600;padding:11px 18px;border-radius:9px;}
  .m-cert{margin:0;font-size:23px;font-weight:800;color:var(--ink);padding-right:30px;}
  .m-pfs{display:flex;flex-wrap:wrap;gap:6px;margin-top:14px;}
  .law{padding:11px 0;border-bottom:1px solid #F0F2F5;cursor:pointer;} .law:last-child{border-bottom:none;} .law:hover{background:#FAFBFC;}
  .law-h{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
  .law-name{font-size:15px;font-weight:600;color:var(--ink);} .law:hover .law-name{color:var(--navy);text-decoration:underline;text-underline-offset:3px;}
  .law-m{font-size:12.5px;color:var(--muted);margin-top:3px;}
  .tag-t2{font-size:11px;color:#3A4862;background:#EEF2F8;border:1px solid #DCE3EE;border-radius:5px;padding:1px 7px;}
  .tag-sjb{font-size:11px;color:#fff;background:#C0492F;border-radius:5px;padding:1px 7px;} .law-go{font-size:12px;color:var(--muted);margin-left:auto;}
  .m2-law{font-size:20px;font-weight:800;color:var(--ink);margin:0 30px 2px 0;} .m2-art{font-size:13px;color:var(--muted);}
  .trk{margin-top:14px;border:1px solid var(--line);border-radius:12px;padding:14px 16px;background:#FAFBFD;}
  .trk .k{font-size:12px;font-weight:700;color:var(--navy);} .trk .v{font-size:14.5px;font-weight:700;color:var(--ink);margin-top:3px;}
  .trk .d{font-size:13.5px;color:var(--body);margin-top:4px;line-height:1.6;} .trk .sub{font-size:11.5px;color:var(--muted);}
  .m2-ext{display:inline-block;margin-top:20px;background:var(--navy);color:#fff;text-decoration:none;font-size:13.5px;font-weight:600;padding:10px 16px;border-radius:9px;}
  footer{margin-top:36px;border-top:1px solid var(--line);background:#fff;} footer .wrap{padding:22px 20px;font-size:12.5px;color:var(--muted);line-height:1.7;} footer b{color:var(--body);font-weight:600;}
  @media(max-width:560px){.grid,.grid.rgrid{grid-template-columns:1fr;}.modal-panel{margin:0;border-radius:0;min-height:100vh;}}
</style></head><body>
<div class="gov-bar"><div class="wrap"><b>한국산업인력공단</b><span>· 국가기술자격 AI 법령 모니터링</span></div></div>
<header class="site"><div class="wrap">
  <span class="logo">자격증 <em>법령 네비게이터</em></span>
  <nav class="tabs">
    <button type="button" class="tab active" data-view="monitor">법령 제개정에 따른 자격증 활용도 모니터링</button>
    <button type="button" class="tab" data-view="radar">자격증별 채용시장 우대사항 모니터링</button>
  </nav>
</div></header>

<!-- ===== 화면1: 활용도 모니터링 ===== -->
<section id="view-monitor">
  <div class="hero"><div class="wrap">
    <div class="eyebrow" id="heroPeriod"></div>
    <h1>선택한 기간 동안, 자격증과 관련된<br><strong id="heroN">0</strong>건의 법령이 바뀌었습니다.</h1>
    <p class="lead">매일 새벽 국가법령정보센터를 자동으로 살펴, 국가기술자격과 관련된 제·개정 법령만 골라 정리합니다. 기간을 선택하거나 검색해 보세요.</p>
    <div class="stats">
      <div class="stat"><div class="n" id="statLaws">0</div><div class="l">관련 법령(선택 기간)</div></div>
      <div class="stat"><div class="n">@@M_TOTAL_CERTS@@</div><div class="l">연관 자격종목(전체)</div></div>
    </div>
  </div></div>
  <div class="toolbar"><div class="wrap">
    <div class="trow period"><span>기간</span>
      <select id="mfrom">@@M_OPTS@@</select><span>~</span><select id="mto">@@M_OPTS@@</select>
      <span class="count">표시 중 <b id="cnt">0</b>건</span></div>
    <div class="trow">
      <select id="scope" aria-label="검색 범위"><option value="all">전체검색</option><option value="law">법령명</option><option value="cert">자격명칭</option><option value="detail">상세내용</option></select>
      <div class="search"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
        <input id="qm" type="search" placeholder="법령명, 자격명칭, 상세내용 검색을 통해 관심내용을 확인하세요!" aria-label="검색"></div>
    </div>
  </div></div>
  <main><div class="wrap"><div class="grid" id="grid-m">@@M_CARDS@@</div><p class="noresult" id="nores-m">조건에 맞는 법령이 없습니다.</p></div></main>
</section>

<!-- ===== 화면2: 자격증 우대사항 ===== -->
<section id="view-radar" hidden>
  <div class="hero"><div class="wrap">
    <div class="eyebrow">자격증으로 찾아보기</div>
    <h1>내 자격증, 어떤 법에서 우대받나요?</h1>
    <p class="lead">자격증을 고르면 그 자격으로 우대(의무고용·직무권한·인사우대·시험면제 등)받는 법령과 근거를 한눈에 봅니다.</p>
  </div></div>
  <div class="toolbar"><div class="wrap"><div class="trow">
    <div class="search"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
      <input id="qr" type="search" placeholder="자격증 이름으로 검색 (예: 전기기사)" aria-label="검색"></div>
    <span class="count">자격증 <b id="cntr">0</b>개</span>
  </div></div></div>
  <main><div class="wrap"><div class="grid rgrid" id="grid-r">@@R_CARDS@@</div><p class="noresult" id="nores-r">해당 자격증이 없습니다.</p></div></main>
</section>

<footer><div class="wrap"><b>안내</b> · 이 페이지는 AI가 법령 원문을 분석하고 정리하였습니다. 정확한 법적 효력은 반드시 <a href="https://www.law.go.kr" target="_blank" rel="noopener" style="color:var(--accent)">국가법령정보센터</a> 원문을 확인하세요. 출처: 국가법령정보센터 | 생성일 @@BUILT_AT@@ | 한국산업인력공단 실증(PoC)</div></footer>

<div class="modal" id="modal" aria-hidden="true" role="dialog" aria-modal="true"><div class="modal-backdrop"></div><div class="modal-panel"><button class="modal-close" aria-label="닫기">&times;</button><div id="m-body"></div></div></div>
<div class="modal" id="modal2" aria-hidden="true" role="dialog" aria-modal="true"><div class="modal-backdrop"></div><div class="modal-panel"><button class="modal-close" aria-label="닫기">&times;</button><div id="m2-body"></div></div></div>

<script>
var MLAWS=@@MLAWS@@, RCERTS=@@RCERTS@@, RENTRIES=@@RENTRIES@@, T1TYPE=@@T1TYPE@@, T1RISK=@@T1RISK@@, T2=@@T2@@, PFC=@@PFC@@;
function escq(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];});}
function tok(v){return String(v||'').split(' ')[0].trim();}
function lawUrl(n){return 'https://www.law.go.kr/법령/'+encodeURIComponent(String(n||'').trim());}
function pfBadge(p){return '<span class="pf" style="--c:'+(PFC[p]||'#8A8F98')+'">'+escq(p)+'</span>';}

// ── 탭 전환 ──
var tabs=[].slice.call(document.querySelectorAll('.tab'));
var views={monitor:document.getElementById('view-monitor'),radar:document.getElementById('view-radar')};
tabs.forEach(function(t){t.addEventListener('click',function(){
  tabs.forEach(function(x){x.classList.remove('active');}); t.classList.add('active');
  for(var k in views) views[k].hidden=(k!==t.dataset.view);
  window.scrollTo(0,0);
});});

// ── monitor 검색/기간 ──
var gm=document.getElementById('grid-m'), mcards=[].slice.call(gm.querySelectorAll('.card'));
var qm=document.getElementById('qm'), scope=document.getElementById('scope');
var mfrom=document.getElementById('mfrom'), mto=document.getElementById('mto'), nresM=document.getElementById('nores-m');
mfrom.value="@@M_DEF_FROM@@"; mto.value="@@M_DEF_TO@@";
MLAWS.forEach(function(o){var cs=(o.certs||[]).join(' ');
  o._law=(o.law||'').toLowerCase(); o._cert=cs.toLowerCase();
  o._det=((o.summary_use||'')+' '+(o.summary_main||'')+' '+(o.meta||'')).toLowerCase();
  o._all=((o.law||'')+' '+(o.meta||'')+' '+cs+' '+(o.summary_use||'')+' '+(o.summary_main||'')).toLowerCase();});
function hay(c){var o=MLAWS[+c.dataset.i];var s=scope.value;return s==='law'?o._law:s==='cert'?o._cert:s==='detail'?o._det:o._all;}
function fmtM(m){return m?m.slice(0,4)+'.'+m.slice(4,6):'';}
function filterM(){var term=(qm.value||'').trim().toLowerCase();var a=mfrom.value,b=mto.value;if(a>b){var t=a;a=b;b=t;}var s=0;
  mcards.forEach(function(c){var on=(c.dataset.month>=a&&c.dataset.month<=b)&&(!term||(hay(c)||'').indexOf(term)!==-1);c.style.display=on?'':'none';if(on)s++;});
  document.getElementById('cnt').textContent=s;document.getElementById('heroN').textContent=s;document.getElementById('statLaws').textContent=s;
  document.getElementById('heroPeriod').textContent=fmtM(a)+' ~ '+fmtM(b)+' 기간';nresM.style.display=s?'none':'block';}
[qm,scope,mfrom,mto].forEach(function(el){el.addEventListener('input',filterM);el.addEventListener('change',filterM);});
filterM();

// ── radar 검색 ──
var gr=document.getElementById('grid-r'), rcards=[].slice.call(gr.querySelectorAll('.card'));
var qr=document.getElementById('qr'), nresR=document.getElementById('nores-r');
function filterR(){var t=(qr.value||'').trim().toLowerCase(),s=0;
  rcards.forEach(function(c){var on=!t||RCERTS[+c.dataset.i].cert.toLowerCase().indexOf(t)!==-1;c.style.display=on?'':'none';if(on)s++;});
  document.getElementById('cntr').textContent=s;nresR.style.display=s?'none':'block';}
qr.addEventListener('input',filterR);filterR();

// ── 모달 ──
var modal=document.getElementById('modal'),mb=document.getElementById('m-body');
var modal2=document.getElementById('modal2'),mb2=document.getElementById('m2-body');
function sec(t,inner){return inner?'<div class="m-sec"><h4>'+t+'</h4>'+inner+'</div>':'';}
function openM(modalEl){modalEl.classList.add('open');modalEl.setAttribute('aria-hidden','false');document.body.style.overflow='hidden';}
function closeModal(){modal.classList.remove('open');modal.setAttribute('aria-hidden','true');if(!modal2.classList.contains('open'))document.body.style.overflow='';}
function closeModal2(){modal2.classList.remove('open');modal2.setAttribute('aria-hidden','true');}

// monitor 법령 상세
function openMonitor(i){var d=MLAWS[i];if(!d)return;
  var certs=(d.certs||[]).map(function(c){return '<span class="chip">'+escq(c)+'</span>';}).join('');
  var arts=(d.articles||[]).map(function(a){return '<li>'+escq(a)+'</li>';}).join('');
  mb.innerHTML='<h2 class="m-title">'+escq(d.law)+'</h2><div class="m-meta">'+escq(d.meta)+'</div>'
    +sec('주요 제·개정 내용',d.summary_main?'<p>'+escq(d.summary_main)+'</p>':'')
    +sec('자격증 활용 분석',d.summary_use?'<p>'+escq(d.summary_use)+'</p>':'')
    +sec('관련 자격종목 ('+((d.certs||[]).length)+'개)',certs?'<div class="m-chips">'+certs+'</div>':'<p class="m-none">없음</p>')
    +sec('근거 조문',arts?'<ul class="m-arts">'+arts+'</ul>':'<p class="m-none">표기된 조문 없음</p>')
    +'<a class="m-ext" href="'+escq(d.url)+'" target="_blank" rel="noopener">법제처에서 원문 보기 →</a>';
  openM(modal);}
// radar 자격증 상세(1차)
function openCert(i){var d=RCERTS[i];if(!d)return;
  var pfs=(d.prefs||[]).map(pfBadge).join('');
  var laws=(d.idx||[]).map(function(ei){var l=RENTRIES[ei];var t2n=(T2[l.t2]||[l.t2])[0];var tags='';
    if(l.t2)tags+=' <span class="tag-t2">'+escq(l.t2+' '+t2n)+'</span>';
    if(l.s)tags+=' <span class="tag-sjb">중처법</span>';
    return '<div class="law" data-ei="'+ei+'"><div class="law-h">'+pfBadge(l.p)+'<span class="law-name">'+escq(l.law)+'</span>'+tags+'<span class="law-go">상세 →</span></div><div class="law-m">'+escq(l.a)+'</div></div>';
  }).join('');
  mb.innerHTML='<h2 class="m-cert">'+escq(d.cert)+'</h2>'
    +'<div class="m-pfs">'+pfs+'</div><div class="m-sec"><h4>이 자격증을 우대하는 법령 ('+(d.idx||[]).length+'건)</h4>'+laws+'</div>';
  openM(modal);}
// radar 법령 상세(2차)
function trkBlock(k,code,name,desc,sub){return '<div class="trk"><div class="k">'+k+'</div><div class="v">'+escq(code)+(name?' · '+escq(name):'')+(sub?' <span class="sub">('+escq(sub)+')</span>':'')+'</div>'+(desc?'<div class="d">'+escq(desc)+'</div>':'')+'</div>';}
function openLaw(ei){var l=RENTRIES[ei];if(!l)return;
  var h='<h2 class="m2-law">'+escq(l.law)+'</h2><div class="m2-art">'+escq(l.a)+'</div><div class="m-pfs" style="margin-top:12px;">'+pfBadge(l.p)+'</div>';
  h+='<div class="m-sec"><h4>상세 분석 결과</h4>'+(l.d?'<p style="margin:0;font-size:14.5px;line-height:1.7;">'+escq(l.d)+'</p>':'<p style="margin:0;font-size:14px;color:var(--muted);">상세 분석 결과는 일일 분석(관련법령) 연동 시 표시됩니다.</p>')+'</div>';
  var tt=T1TYPE[l.t1],tr=T1RISK[l.t1r];
  h+='<div class="m-sec"><h4>정책 관점 분류 (Track 1)</h4>';
  if(tt)h+=trkBlock('자격을 다루는 방식 · 취급유형',l.t1,tt[0],tt[1]);
  if(tr)h+=trkBlock('경력이음 위험도',l.t1r,tr[0],tr[1]);
  if(!tt&&!tr)h+='<p class="law-m">분류 정보 없음</p>';h+='</div>';
  var t2=T2[l.t2];h+='<div class="m-sec"><h4>국민 취업정보 관점 분류 (Track 2)</h4>';
  if(t2)h+=trkBlock('노동시장 효용 · 효용코드',l.t2,t2[0],t2[2],t2[1]);else h+='<p class="law-m">분류 정보 없음</p>';h+='</div>';
  h+='<a class="m2-ext" href="'+escq(lawUrl(l.law))+'" target="_blank" rel="noopener">법제처에서 원문 보기 →</a>';
  mb2.innerHTML=h;openM(modal2);}

gm.addEventListener('click',function(e){var t=e.target.closest('.title-btn,.detail-link');if(!t)return;var c=t.closest('.card');if(c)openMonitor(+c.dataset.i);});
gr.addEventListener('click',function(e){var t=e.target.closest('.title-btn,.detail-link');if(!t)return;var c=t.closest('.card');if(c)openCert(+c.dataset.i);});
mb.addEventListener('click',function(e){var law=e.target.closest('.law');if(law)openLaw(+law.dataset.ei);});
modal.addEventListener('click',function(e){if(e.target.classList.contains('modal-backdrop')||e.target.closest('.modal-close'))closeModal();});
modal2.addEventListener('click',function(e){if(e.target.classList.contains('modal-backdrop')||e.target.closest('.modal-close'))closeModal2();});
document.addEventListener('keydown',function(e){if(e.key==='Escape'){if(modal2.classList.contains('open'))closeModal2();else closeModal();}});
</script></body></html>
"""

if __name__ == "__main__":
    main()
