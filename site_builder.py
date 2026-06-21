# -*- coding: utf-8 -*-
"""
site_builder.py — 대국민 법령 네비게이터 정적 사이트 빌더
================================================================

[하는 일]
  구글 시트(모니터링 마스터 DB)에서 최신 분석 결과를 읽어,
  국민이 보기 좋은 공개용 정적 HTML(index.html)을 한 장 찍어낸다.
  서버가 필요 없다 — GitHub Actions 배치가 이 파일을 실행해
  결과 HTML을 GitHub Pages로 올리면 그게 곧 공개 웹사이트가 된다.

[설계 원칙]
  - 외부 레포 의존이 없다. 구글시트만 읽으면 되므로 시트 인증을 자체 포함한다.
    (gspread 하나만 있으면 운영 모드가 돈다.)
  - 기존 배치 코드(main/brain/config)는 건드리지 않는다. 이 파일은 별도 레포의 단독 모듈.
  - 데이터는 "자동으로 채워진다". 사람이 수기로 업로드하지 않는다.

[두 가지 실행 모드]
  1) 운영 모드 (기본): 환경변수로 구글시트에 접속해 읽는다.
       GCP_SA_JSON        : 서비스계정 JSON (코어와 동일한 시크릿)
       GOOGLE_SHEET_ID    : 시트 URL 또는 KEY (코어와 동일)
       SOURCE_WORKSHEET   : 읽을 탭 이름 (기본 "연관 높은 법령")
  2) 미리보기 모드: 로컬 xlsx로 화면만 확인할 때.
       LOCAL_XLSX         : xlsx 경로를 지정하면 시트 대신 이 파일을 읽는다.
       LOCAL_SHEET        : xlsx 안에서 읽을 시트 이름 (기본 SOURCE_WORKSHEET)

[그 외 옵션]
  TARGET_MONTH (YYYYMM) : 특정 달만 보여줄 때. 비우면 데이터의 '가장 최근 달' 자동.
  MAX_CARDS             : 한 페이지 최대 카드 수 (기본 300, 과대 HTML 방지).
  OUT_DIR               : 결과물 폴더 (기본 "dist"). 여기에 index.html이 생성된다.

실행:  python site_builder.py
"""

import os
import re
import json
import html
import datetime
from urllib.parse import quote
from collections import Counter

# ─────────────────────────────────────────────────────────────
# 0. 설정 — 시트 컬럼 이름은 여기서 한 곳으로 관리 (헤더가 바뀌면 여기만 수정)
# ─────────────────────────────────────────────────────────────
COL = {
    "law":      "법령명",
    "ministry": "소관부처",
    "date":     "시행일자",
    "kind":     "개정유형",
    "summary1": "활용도 분석 상세",       # 국민 눈높이 요약 (우선)
    "summary2": "주요 제·개정내용",        # 보조 요약
    "certs":    "법령 관련 국가기술자격 종목",
    "article":  "근거조문",
    "link":     "조문별 다이렉트 링크",     # 있으면 우선 사용
}

SOURCE_WORKSHEET = os.environ.get("SOURCE_WORKSHEET", "연관 높은 법령")
TARGET_MONTH     = os.environ.get("TARGET_MONTH", "").strip()
MAX_CARDS        = int(os.environ.get("MAX_CARDS", "300"))
OUT_DIR          = os.environ.get("OUT_DIR", "dist")

# ─────────────────────────────────────────────────────────────
# 1. 분야 분류 — 이슈브리핑(briefing_maker)의 FIELD_KEYWORDS 로직과 동일 개념.
#    종목명 키워드로 자격 분야를 추정해 카드 색상/필터에 쓴다.
# ─────────────────────────────────────────────────────────────
FIELD_KEYWORDS = {
    "건설·건축": ["건축", "토목", "건설", "조경", "측량", "도시", "구조", "콘크리트"],
    "전기·전자": ["전기", "전자", "통신", "정보통신", "반도체", "제어"],
    "기계·금속": ["기계", "금속", "용접", "주조", "판금", "배관", "냉동", "설비"],
    "화학·환경": ["화학", "환경", "대기", "수질", "폐기물", "위험물", "에너지"],
    "안전·소방": ["안전", "소방", "방재", "산업위생", "가스"],
    "정보·SW":   ["정보처리", "소프트웨어", "데이터", "정보보안", "임베디드", "빅데이터"],
    "보건·식품": ["식품", "조리", "위생", "보건", "의료", "영양"],
    "농림·수산": ["농", "임업", "수산", "축산", "원예", "종자"],
}
FIELD_ORDER = list(FIELD_KEYWORDS.keys()) + ["기타"]

# 분야별 카드 좌측 강조선 색 (차분한 공공 톤)
FIELD_COLOR = {
    "건설·건축": "#B8742A", "전기·전자": "#1F6FB2", "기계·금속": "#5A6472",
    "화학·환경": "#2E8B6B", "안전·소방": "#C0492F", "정보·SW": "#5B4BB0",
    "보건·식품": "#B0476A", "농림·수산": "#6E8B2E", "기타": "#8A8F98",
}

def classify_field(cert_string: str) -> str:
    s = str(cert_string or "")
    for field, kws in FIELD_KEYWORDS.items():
        if any(kw in s for kw in kws):
            return field
    return "기타"


# ─────────────────────────────────────────────────────────────
# 2. 데이터 로드 — 운영(시트) / 미리보기(xlsx) 두 모드 통합
# ─────────────────────────────────────────────────────────────
def _sheet_key(v: str) -> str:
    """전체 URL을 넣어도 시트 KEY만 뽑아낸다(KEY를 그대로 넣어도 OK)."""
    m = re.search(r"/d/([A-Za-z0-9_-]+)", str(v or ""))
    return m.group(1) if m else str(v or "").strip()


def _open_spreadsheet():
    """서비스계정 JSON(문자열) + 시트 KEY로 스프레드시트를 연다. gspread만 의존."""
    import gspread  # 운영 모드에서만 필요 (미리보기 모드는 불필요)
    creds_dict = json.loads(os.environ["GCP_SA_JSON"].strip(), strict=False)
    key = _sheet_key(os.environ["GOOGLE_SHEET_ID"])
    gc = gspread.service_account_from_dict(creds_dict)
    return gc.open_by_key(key)


def load_rows():
    """행 목록(list[dict])을 반환. 두 모드 모두 같은 모양(헤더=키)으로 맞춘다."""
    local_xlsx = os.environ.get("LOCAL_XLSX", "").strip()
    if local_xlsx:
        # ── 미리보기 모드: 로컬 xlsx (pandas/openpyxl 필요) ──
        import pandas as pd
        sheet = os.environ.get("LOCAL_SHEET", SOURCE_WORKSHEET)
        print(f"🖼  미리보기 모드: {local_xlsx} [{sheet}]")
        df = pd.read_excel(local_xlsx, sheet_name=sheet)
        df = df.fillna("")
        return df.to_dict("records")

    # ── 운영 모드: 구글 시트 (자체 인증 — 외부 레포 의존 없음) ──
    ss = _open_spreadsheet()
    print(f"📄 시트 읽기: [{SOURCE_WORKSHEET}]")
    return ss.worksheet(SOURCE_WORKSHEET).get_all_records()


# ─────────────────────────────────────────────────────────────
# 3. 유틸
# ─────────────────────────────────────────────────────────────
def digits(v) -> str:
    """날짜값(문자/Timestamp 무엇이든)에서 숫자만 뽑아 YYYYMMDD 형태로."""
    return re.sub(r"\D", "", str(v or ""))[:8]

def fmt_date(v) -> str:
    d = digits(v)
    if len(d) == 8:
        return f"{d[:4]}.{d[4:6]}.{d[6:]}"
    return str(v or "")

def law_url(row) -> str:
    """법제처 원문 링크. '조문별 다이렉트 링크'가 있으면 우선, 없으면 법령명으로 조합."""
    direct = str(row.get(COL["link"], "") or "").strip()
    if direct.startswith("http"):
        return direct
    name = str(row.get(COL["law"], "") or "").strip()
    return f"https://www.law.go.kr/법령/{quote(name)}"

def cert_chips(cert_string: str, limit=6) -> str:
    items = [c.strip() for c in re.split(r"[,/·\n]", str(cert_string or "")) if c.strip()]
    shown = items[:limit]
    extra = len(items) - len(shown)
    chips = "".join(f'<span class="chip">{html.escape(c)}</span>' for c in shown)
    if extra > 0:
        chips += f'<span class="chip chip-more">+{extra}</span>'
    return chips

def esc(v) -> str:
    return html.escape(str(v or "").strip())


# ─────────────────────────────────────────────────────────────
# 4. 카드 렌더링
# ─────────────────────────────────────────────────────────────
def render_card(row) -> str:
    law      = esc(row.get(COL["law"]))
    ministry = esc(row.get(COL["ministry"]))
    date     = fmt_date(row.get(COL["date"]))
    kind     = esc(row.get(COL["kind"]))
    certs    = row.get(COL["certs"], "")
    field    = classify_field(certs)
    color    = FIELD_COLOR.get(field, FIELD_COLOR["기타"])
    article  = esc(row.get(COL["article"]))
    url      = law_url(row)

    summary = str(row.get(COL["summary1"], "") or "").strip() \
        or str(row.get(COL["summary2"], "") or "").strip()
    summary = esc(summary) or "요약 준비 중입니다."

    meta = " · ".join(x for x in [ministry, date, kind] if x)
    haystack = esc(f"{law} {ministry} {certs} {summary}").lower()

    article_html = f'<div class="article">근거조문 · {article}</div>' if article else ""

    return f"""
    <article class="card" data-field="{esc(field)}" data-q="{haystack}" style="--fc:{color}">
      <div class="card-head">
        <span class="field-tag" style="color:{color}">{esc(field)}</span>
        <span class="meta">{esc(meta)}</span>
      </div>
      <h3 class="card-title"><a href="{esc(url)}" target="_blank" rel="noopener">{law}</a></h3>
      <div class="chips">{cert_chips(certs)}</div>
      <p class="summary">{summary}</p>
      {article_html}
      <a class="more" href="{esc(url)}" target="_blank" rel="noopener">법제처 원문 보기 →</a>
    </article>"""


# ─────────────────────────────────────────────────────────────
# 5. 페이지 조립
# ─────────────────────────────────────────────────────────────
def build_html(rows) -> str:
    # (a) 달 필터: TARGET_MONTH 지정 없으면 데이터의 가장 최근 달
    months = sorted({digits(r.get(COL["date"]))[:6] for r in rows if digits(r.get(COL["date"]))})
    month = TARGET_MONTH or (months[-1] if months else "")
    if month:
        rows = [r for r in rows if digits(r.get(COL["date"])).startswith(month)]
    rows = rows[:MAX_CARDS]

    # (b) 분야 집계 (필터칩 + 통계)
    field_counts = Counter(classify_field(r.get(COL["certs"], "")) for r in rows)
    present_fields = [f for f in FIELD_ORDER if field_counts.get(f)]

    # (c) 통계
    total_laws = len(rows)
    total_certs = len({c.strip() for r in rows
                       for c in re.split(r"[,/·\n]", str(r.get(COL["certs"], "") or "")) if c.strip()})
    month_label = f"{month[:4]}년 {int(month[4:6])}월" if len(month) == 6 else "최근"
    built_at = datetime.datetime.now().strftime("%Y.%m.%d")

    # (d) 필터칩
    chip_html = '<button class="fchip active" data-f="전체">전체</button>'
    for f in present_fields:
        c = FIELD_COLOR.get(f, "#888")
        chip_html += (f'<button class="fchip" data-f="{esc(f)}">'
                      f'<i style="background:{c}"></i>{esc(f)}'
                      f'<b>{field_counts[f]}</b></button>')

    cards = "\n".join(render_card(r) for r in rows) or \
        '<p class="empty">표시할 법령이 없습니다.</p>'

    return PAGE.format(
        month_label=month_label, total_laws=total_laws, total_certs=total_certs,
        field_n=len(present_fields), chips=chip_html, cards=cards, built_at=built_at,
    )


def main():
    rows = load_rows()
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, "index.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(build_html(rows))
    print(f"✅ 생성 완료: {out}  ({len(rows)}행 입력)")


# ─────────────────────────────────────────────────────────────
# 6. HTML 템플릿 (디자인은 한 번만 정의 — 내용은 위에서 자동 주입)
# ─────────────────────────────────────────────────────────────
PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>자격증 법령 네비게이터 · HRDK</title>
<meta name="description" content="내 자격증과 관련해 최근 바뀐 법령을 쉬운 말로. 한국산업인력공단 대국민 법령 모니터링.">
<link rel="preconnect" href="https://cdn.jsdelivr.net">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
  :root {{
    --navy:#1F3864; --ink:#16243F; --body:#33394A; --muted:#6B7280;
    --line:#E4E7EC; --bg:#F6F8FB; --surface:#FFFFFF; --accent:#0F6E56;
  }}
  * {{ box-sizing:border-box; }}
  html {{ scroll-behavior:smooth; }}
  body {{
    margin:0; background:var(--bg); color:var(--body);
    font-family:"Pretendard","Pretendard Variable",-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;
    font-size:16px; line-height:1.65; -webkit-font-smoothing:antialiased;
  }}
  a {{ color:inherit; }}
  .wrap {{ max-width:1080px; margin:0 auto; padding:0 20px; }}

  /* 상단 식별 바 (KRDS 식별자 느낌) */
  .gov-bar {{ background:var(--navy); color:#fff; font-size:13px; }}
  .gov-bar .wrap {{ padding:7px 20px; display:flex; gap:8px; align-items:center; }}
  .gov-bar b {{ font-weight:600; }}
  .gov-bar span {{ opacity:.8; }}

  /* 헤더 */
  header.site {{ background:var(--surface); border-bottom:1px solid var(--line); }}
  header.site .wrap {{ padding:18px 20px; display:flex; align-items:baseline; gap:12px; flex-wrap:wrap; }}
  .logo {{ font-size:19px; font-weight:700; color:var(--ink); letter-spacing:-.01em; }}
  .logo em {{ color:var(--navy); font-style:normal; }}
  .tag {{ font-size:13.5px; color:var(--muted); }}

  /* 히어로 — 데이터 진술형 */
  .hero {{ background:var(--surface); border-bottom:1px solid var(--line); }}
  .hero .wrap {{ padding:40px 20px 30px; }}
  .eyebrow {{ font-size:13px; font-weight:600; color:var(--accent); letter-spacing:.02em; }}
  .hero h1 {{
    margin:.4em 0 .1em; font-size:clamp(26px,4vw,38px); line-height:1.25;
    font-weight:700; color:var(--ink); letter-spacing:-.02em; max-width:18em;
  }}
  .hero h1 strong {{ color:var(--navy); }}
  .hero p.lead {{ margin:8px 0 0; color:var(--muted); font-size:15.5px; max-width:40em; }}
  .stats {{ display:flex; gap:34px; margin-top:26px; flex-wrap:wrap; }}
  .stat .n {{ font-size:30px; font-weight:700; color:var(--navy); font-variant-numeric:tabular-nums; line-height:1; }}
  .stat .l {{ font-size:13px; color:var(--muted); margin-top:5px; }}

  /* 필터 바 (스티키) */
  .toolbar {{ position:sticky; top:0; z-index:5; background:rgba(246,248,251,.92);
    backdrop-filter:saturate(160%) blur(6px); border-bottom:1px solid var(--line); }}
  .toolbar .wrap {{ padding:13px 20px; display:flex; gap:12px; align-items:center; flex-wrap:wrap; }}
  .chips-row {{ display:flex; gap:8px; flex-wrap:wrap; flex:1; min-width:240px; }}
  .fchip {{ display:inline-flex; align-items:center; gap:6px; cursor:pointer;
    border:1px solid var(--line); background:var(--surface); color:var(--body);
    border-radius:999px; padding:6px 13px; font-size:13.5px; font-family:inherit;
    transition:border-color .15s, color .15s, background .15s; }}
  .fchip i {{ width:9px; height:9px; border-radius:50%; display:inline-block; }}
  .fchip b {{ font-weight:600; color:var(--muted); font-variant-numeric:tabular-nums; }}
  .fchip:hover {{ border-color:#C7CDD6; }}
  .fchip.active {{ background:var(--navy); border-color:var(--navy); color:#fff; }}
  .fchip.active b {{ color:rgba(255,255,255,.75); }}
  .search {{ position:relative; }}
  .search input {{ border:1px solid var(--line); border-radius:999px; padding:8px 15px 8px 36px;
    font-size:14px; font-family:inherit; width:220px; background:var(--surface); color:var(--ink);
    outline:none; transition:border-color .15s, box-shadow .15s; }}
  .search input:focus {{ border-color:var(--navy); box-shadow:0 0 0 3px rgba(31,56,100,.12); }}
  .search svg {{ position:absolute; left:13px; top:50%; transform:translateY(-50%); color:var(--muted); }}

  /* 카드 그리드 */
  main .wrap {{ padding:28px 20px 10px; }}
  .grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(330px,1fr)); gap:16px; }}
  .card {{ background:var(--surface); border:1px solid var(--line); border-left:4px solid var(--fc);
    border-radius:12px; padding:18px 19px; display:flex; flex-direction:column;
    transition:box-shadow .18s, transform .18s; }}
  .card:hover {{ box-shadow:0 6px 22px -8px rgba(22,36,63,.18); transform:translateY(-2px); }}
  .card-head {{ display:flex; align-items:center; justify-content:space-between; gap:10px; }}
  .field-tag {{ font-size:12.5px; font-weight:600; }}
  .meta {{ font-size:12.5px; color:var(--muted); text-align:right; }}
  .card-title {{ margin:10px 0 0; font-size:17px; font-weight:600; line-height:1.4; letter-spacing:-.01em; }}
  .card-title a {{ color:var(--ink); text-decoration:none; }}
  .card-title a:hover {{ color:var(--navy); text-decoration:underline; text-underline-offset:3px; }}
  .chips {{ display:flex; flex-wrap:wrap; gap:6px; margin:11px 0 12px; }}
  .chip {{ font-size:12px; background:#EEF2F8; color:#3A4862; border-radius:6px; padding:3px 8px; }}
  .chip-more {{ background:transparent; color:var(--muted); }}
  .summary {{ margin:0; font-size:14.5px; color:var(--body); line-height:1.62;
    display:-webkit-box; -webkit-line-clamp:4; -webkit-box-orient:vertical; overflow:hidden; }}
  .article {{ margin-top:11px; font-size:12.5px; color:var(--muted); border-top:1px dashed var(--line); padding-top:10px; }}
  .more {{ margin-top:13px; align-self:flex-start; font-size:13.5px; font-weight:600;
    color:var(--accent); text-decoration:none; }}
  .more:hover {{ text-decoration:underline; text-underline-offset:3px; }}
  .empty {{ color:var(--muted); padding:40px 0; text-align:center; }}
  .noresult {{ display:none; color:var(--muted); padding:40px 0; text-align:center; }}

  /* 푸터 */
  footer {{ margin-top:36px; border-top:1px solid var(--line); background:var(--surface); }}
  footer .wrap {{ padding:24px 20px; font-size:12.5px; color:var(--muted); line-height:1.7; }}
  footer b {{ color:var(--body); font-weight:600; }}

  @media (max-width:560px) {{
    .grid {{ grid-template-columns:1fr; }}
    .search input {{ width:100%; }}
    .search {{ width:100%; }}
  }}
  @media (prefers-reduced-motion:reduce) {{
    .card {{ transition:none; }} html {{ scroll-behavior:auto; }}
  }}
  /* 로드 페이드 (절제) */
  .card {{ opacity:0; animation:rise .4s ease forwards; }}
  @keyframes rise {{ to {{ opacity:1; transform:translateY(0); }} }}
  @media (prefers-reduced-motion:reduce) {{ .card {{ opacity:1; animation:none; }} }}
</style>
</head>
<body>
  <div class="gov-bar"><div class="wrap"><b>한국산업인력공단</b><span>· 국가기술자격 법령 모니터링</span></div></div>

  <header class="site"><div class="wrap">
    <span class="logo">자격증 <em>법령 네비게이터</em></span>
    <span class="tag">내 자격증과 관련해 바뀐 법, 쉬운 말로</span>
  </div></header>

  <section class="hero"><div class="wrap">
    <div class="eyebrow">{month_label} 업데이트</div>
    <h1>이번 달, 당신의 자격증과 관련해<br><strong>{total_laws}건</strong>의 법령이 바뀌었습니다.</h1>
    <p class="lead">매일 새벽 국가법령정보센터를 자동으로 살펴, 국가기술자격과 관련된 제·개정 법령만 골라 쉬운 말로 정리합니다.</p>
    <div class="stats">
      <div class="stat"><div class="n">{total_laws}</div><div class="l">관련 법령</div></div>
      <div class="stat"><div class="n">{total_certs}</div><div class="l">연관 자격종목</div></div>
      <div class="stat"><div class="n">{field_n}</div><div class="l">자격 분야</div></div>
    </div>
  </div></section>

  <div class="toolbar"><div class="wrap">
    <div class="chips-row">{chips}</div>
    <div class="search">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
      <input id="q" type="search" placeholder="법령·자격·내용 검색" aria-label="검색">
    </div>
  </div></div>

  <main><div class="wrap">
    <div class="grid" id="grid">{cards}</div>
    <p class="noresult" id="noresult">조건에 맞는 법령이 없습니다.</p>
  </div></main>

  <footer><div class="wrap">
    <b>안내</b> · 이 페이지는 AI가 법령 원문을 국민 눈높이로 요약한 자동 생성물입니다.
    정확한 법적 효력은 반드시 <a href="https://www.law.go.kr" target="_blank" rel="noopener" style="color:var(--accent)">국가법령정보센터</a> 원문을 확인하세요.
    출처: 국가법령정보센터 · 워크넷 &nbsp;|&nbsp; 생성일 {built_at} &nbsp;|&nbsp; 한국산업인력공단 실증(PoC)
  </div></footer>

<script>
  // 분야 필터 + 검색 (클라이언트, 라이브러리 없음)
  var grid = document.getElementById('grid');
  var cards = Array.prototype.slice.call(grid.querySelectorAll('.card'));
  var chips = Array.prototype.slice.call(document.querySelectorAll('.fchip'));
  var q = document.getElementById('q');
  var noresult = document.getElementById('noresult');
  var curField = '전체';

  function apply() {{
    var term = (q.value || '').trim().toLowerCase();
    var shown = 0;
    cards.forEach(function(c) {{
      var okF = (curField === '전체') || (c.dataset.field === curField);
      var okQ = !term || (c.dataset.q.indexOf(term) !== -1);
      var on = okF && okQ;
      c.style.display = on ? '' : 'none';
      if (on) shown++;
    }});
    noresult.style.display = shown ? 'none' : 'block';
  }}
  chips.forEach(function(ch) {{
    ch.addEventListener('click', function() {{
      chips.forEach(function(x) {{ x.classList.remove('active'); }});
      ch.classList.add('active');
      curField = ch.dataset.f;
      apply();
    }});
  }});
  q.addEventListener('input', apply);
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
