# -*- coding: utf-8 -*-
"""
site_builder.py — 대국민 법령 네비게이터 정적 사이트 빌더 (v2: 요약 카드 + 상세 팝업)
================================================================

[하는 일]
  구글 시트(모니터링 마스터 DB)에서 최신 분석 결과를 읽어,
  국민이 보기 좋은 공개용 정적 HTML(index.html)을 한 장 찍어낸다.
  서버가 필요 없다 — GitHub Actions 배치가 이 파일을 실행해
  결과 HTML을 GitHub Pages로 올리면 그게 곧 공개 웹사이트가 된다.

[v2 변경점]
  - 첫 화면 카드는 "요약"만 보여준다 (요약문은 3줄로 줄이고, 근거조문은 카드에서 뺌).
  - 법령명 또는 "분석 상세 보기"를 누르면 같은 페이지에서 팝업(모달)이 떠서
    전체 분석 내용(요약 전문 + 관련 종목 전체 + 근거조문 전체)을 보여준다.
  - 카드 아래 "법제처 원문" 링크는 그대로 외부 이동용으로 유지.

[설계 원칙]
  - 외부 레포 의존이 없다. 구글시트만 읽으면 되므로 시트 인증을 자체 포함한다.
  - 기존 배치 코드(main/brain/config)는 건드리지 않는다. 이 파일은 별도 레포의 단독 모듈.

[두 가지 실행 모드]
  1) 운영 모드 (기본): 환경변수로 구글시트에 접속해 읽는다.
       GCP_SA_JSON        : 서비스계정 JSON 문자열 (코어와 동일한 시크릿)
       GOOGLE_SHEET_ID    : 시트 KEY 또는 URL
       SOURCE_WORKSHEET   : 읽을 탭 이름 (기본 "연관 높은 법령")
  2) 미리보기 모드: 로컬 xlsx로 화면만 확인할 때.
       LOCAL_XLSX         : xlsx 경로 (지정하면 시트 대신 이 파일을 읽음)
       LOCAL_SHEET        : xlsx 안에서 읽을 시트 이름 (기본 SOURCE_WORKSHEET)

[그 외 옵션]
  TARGET_MONTH (YYYYMM) : 특정 달만. 비우면 데이터의 '가장 최근 달' 자동.
  MAX_CARDS             : 한 페이지 최대 카드 수 (기본 300).
  OUT_DIR               : 결과물 폴더 (기본 "dist").

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
    "summary1": "활용도 분석 상세",       # 자격증 활용 분석
    "summary2": "주요 제·개정내용",        # 법령의 주요 변경 내용
    "certs":    "법령 관련 국가기술자격 종목",
    "article":  "근거조문",
    "link":     "조문별 다이렉트 링크",     # 있으면 우선 사용
}

SOURCE_WORKSHEET = os.environ.get("SOURCE_WORKSHEET", "연관 높은 법령")
TARGET_MONTH     = os.environ.get("TARGET_MONTH", "").strip()
MAX_CARDS        = int(os.environ.get("MAX_CARDS", "300"))
OUT_DIR          = os.environ.get("OUT_DIR", "dist")

# ─────────────────────────────────────────────────────────────
# 1. 분야 분류 — 종목명 키워드로 자격 분야를 추정해 카드 색상/필터에 쓴다.
#    (이슈브리핑의 FIELD_KEYWORDS 개념과 동일)
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

FIELD_COLOR = {
    "건설·건축": "#B8742A", "전기·전자": "#1F6FB2", "기계·금속": "#5A6472",
    "화학·환경": "#2E8B6B", "안전·소방": "#C0492F", "정보·SW": "#5B4BB0",
    "보건·식품": "#B0476A", "농림·수산": "#6E8B2E", "기타": "#8A8F98",
}

def classify_field(cert_string: str) -> str:
    """관련 종목 중 가장 많이 매칭되는 분야로 태그한다(동률이면 정의 순서 우선)."""
    s = str(cert_string or "")
    items = re.split(r"[,/·\n]", s)
    score = Counter()
    for field, kws in FIELD_KEYWORDS.items():
        for it in items:
            if any(kw in it for kw in kws):
                score[field] += 1
    if not score:
        return "기타"
    # 최다 득점 분야 (동률이면 FIELD_ORDER 순서가 앞선 것)
    best = max(FIELD_KEYWORDS.keys(), key=lambda f: (score[f], -FIELD_ORDER.index(f)))
    return best if score[best] else "기타"


# ─────────────────────────────────────────────────────────────
# 2. 데이터 로드 — 운영(시트) / 미리보기(xlsx) 두 모드 통합
# ─────────────────────────────────────────────────────────────
def _sheet_key(v: str) -> str:
    """전체 URL을 넣어도 시트 KEY만 뽑아낸다(KEY를 그대로 넣어도 OK)."""
    m = re.search(r"/d/([A-Za-z0-9_-]+)", str(v or ""))
    return m.group(1) if m else str(v or "").strip()


def _open_spreadsheet():
    """서비스계정 JSON(문자열) + 시트 KEY로 스프레드시트를 연다. gspread만 의존."""
    import gspread  # 운영 모드에서만 필요
    creds_dict = json.loads(os.environ["GCP_SA_JSON"].strip(), strict=False)
    key = _sheet_key(os.environ["GOOGLE_SHEET_ID"])
    gc = gspread.service_account_from_dict(creds_dict)
    return gc.open_by_key(key)


def load_rows():
    """행 목록(list[dict])을 반환. 두 모드 모두 같은 모양(헤더=키)으로 맞춘다."""
    local_xlsx = os.environ.get("LOCAL_XLSX", "").strip()
    if local_xlsx:
        import pandas as pd
        sheet = os.environ.get("LOCAL_SHEET", SOURCE_WORKSHEET)
        print(f"🖼  미리보기 모드: {local_xlsx} [{sheet}]")
        df = pd.read_excel(local_xlsx, sheet_name=sheet).fillna("")
        return df.to_dict("records")

    ss = _open_spreadsheet()
    print(f"📄 시트 읽기: [{SOURCE_WORKSHEET}]")
    return ss.worksheet(SOURCE_WORKSHEET).get_all_records()


# ─────────────────────────────────────────────────────────────
# 3. 유틸
# ─────────────────────────────────────────────────────────────
def digits(v) -> str:
    return re.sub(r"\D", "", str(v or ""))[:8]

def fmt_date(v) -> str:
    d = digits(v)
    return f"{d[:4]}.{d[4:6]}.{d[6:]}" if len(d) == 8 else str(v or "")

def law_url(row) -> str:
    direct = str(row.get(COL["link"], "") or "").strip()
    if direct.startswith("http"):
        return direct
    name = str(row.get(COL["law"], "") or "").strip()
    return f"https://www.law.go.kr/법령/{quote(name)}"

def esc(v) -> str:
    return html.escape(str(v or "").strip())


# ─────────────────────────────────────────────────────────────
# 4. 행 → 정제된 필드 묶음 (카드와 팝업이 함께 사용)
# ─────────────────────────────────────────────────────────────
def card_fields(row) -> dict:
    certs = [c.strip() for c in re.split(r"[,/·\n]", str(row.get(COL["certs"]) or "")) if c.strip()]
    articles = [a.strip() for a in re.split(r"[,\n;·]", str(row.get(COL["article"]) or "")) if a.strip()]
    field = classify_field(row.get(COL["certs"], ""))
    ministry = str(row.get(COL["ministry"]) or "").strip()
    date = fmt_date(row.get(COL["date"]))
    kind = str(row.get(COL["kind"]) or "").strip()
    return {
        "law":          str(row.get(COL["law"]) or "").strip(),
        "meta":         " · ".join(x for x in [ministry, date, kind] if x),
        "field":        field,
        "color":        FIELD_COLOR.get(field, FIELD_COLOR["기타"]),
        "certs":        certs,
        "summary_main": str(row.get(COL["summary2"]) or "").strip(),  # 주요 제·개정내용
        "summary_use":  str(row.get(COL["summary1"]) or "").strip(),  # 활용도 분석 상세
        "articles":     articles,
        "url":          law_url(row),
    }


def render_card(d: dict, i: int) -> str:
    """첫 화면 요약 카드. 클릭 시 팝업(openModal)으로 전체를 본다."""
    shown = d["certs"][:4]
    extra = len(d["certs"]) - len(shown)
    chips = "".join(f'<span class="chip">{esc(c)}</span>' for c in shown)
    if extra > 0:
        chips += f'<span class="chip chip-more">+{extra}</span>'

    summary = esc(d["summary_use"] or d["summary_main"] or "요약 준비 중입니다.")
    haystack = esc(f'{d["law"]} {d["meta"]} {" ".join(d["certs"])} {d["summary_use"]} {d["summary_main"]}').lower()

    return f"""
    <article class="card" data-i="{i}" data-field="{esc(d['field'])}" data-q="{haystack}" style="--fc:{d['color']}">
      <div class="card-head">
        <span class="field-tag" style="color:{d['color']}">{esc(d['field'])}</span>
        <span class="meta">{esc(d['meta'])}</span>
      </div>
      <h3 class="card-title"><button type="button" class="title-btn">{esc(d['law'])}</button></h3>
      <div class="chips">{chips}</div>
      <p class="summary">{summary}</p>
      <div class="card-foot">
        <button type="button" class="detail-link">분석 상세 보기 →</button>
        <a class="ext" href="{esc(d['url'])}" target="_blank" rel="noopener">법제처 원문</a>
      </div>
    </article>"""


# ─────────────────────────────────────────────────────────────
# 5. 페이지 조립
# ─────────────────────────────────────────────────────────────
def build_html(rows) -> str:
    months = sorted({digits(r.get(COL["date"]))[:6] for r in rows if digits(r.get(COL["date"]))})
    month = TARGET_MONTH or (months[-1] if months else "")
    if month:
        rows = [r for r in rows if digits(r.get(COL["date"])).startswith(month)]
    rows = rows[:MAX_CARDS]

    data = [card_fields(r) for r in rows]
    field_counts = Counter(d["field"] for d in data)
    present_fields = [f for f in FIELD_ORDER if field_counts.get(f)]

    total_laws = len(data)
    total_certs = len({c for d in data for c in d["certs"]})
    month_label = f"{month[:4]}년 {int(month[4:6])}월" if len(month) == 6 else "최근"
    built_at = datetime.datetime.now().strftime("%Y.%m.%d")

    chip_html = '<button class="fchip active" data-f="전체">전체</button>'
    for f in present_fields:
        c = FIELD_COLOR.get(f, "#888")
        chip_html += (f'<button class="fchip" data-f="{esc(f)}">'
                      f'<i style="background:{c}"></i>{esc(f)}<b>{field_counts[f]}</b></button>')

    cards = "\n".join(render_card(d, i) for i, d in enumerate(data)) or \
        '<p class="empty">표시할 법령이 없습니다.</p>'

    # 팝업이 쓸 전체 데이터를 JS로 주입 (</ 는 스크립트 조기종료 방지로 이스케이프)
    laws_json = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")

    out = PAGE
    for k, v in {
        "@@MONTH_LABEL@@": month_label,
        "@@TOTAL_LAWS@@":  str(total_laws),
        "@@TOTAL_CERTS@@": str(total_certs),
        "@@FIELD_N@@":     str(len(present_fields)),
        "@@CHIPS@@":       chip_html,
        "@@CARDS@@":       cards,
        "@@BUILT_AT@@":    built_at,
        "@@LAWS_JSON@@":   laws_json,
    }.items():
        out = out.replace(k, v)
    return out


def main():
    rows = load_rows()
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, "index.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(build_html(rows))
    print(f"✅ 생성 완료: {out}  ({len(rows)}행 입력)")


# ─────────────────────────────────────────────────────────────
# 6. HTML 템플릿 (디자인은 한 번만 정의 — 내용은 @@토큰@@으로 자동 주입)
# ─────────────────────────────────────────────────────────────
PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>자격증 법령 네비게이터 · HRDK</title>
<meta name="description" content="내 자격증과 관련해 최근 바뀐 법령을 쉬운 말로. 한국산업인력공단 대국민 법령 모니터링.">
<link rel="preconnect" href="https://cdn.jsdelivr.net">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
  :root {
    --navy:#1F3864; --ink:#16243F; --body:#33394A; --muted:#6B7280;
    --line:#E4E7EC; --bg:#F6F8FB; --surface:#FFFFFF; --accent:#0F6E56;
  }
  * { box-sizing:border-box; }
  html { scroll-behavior:smooth; }
  body {
    margin:0; background:var(--bg); color:var(--body);
    font-family:"Pretendard","Pretendard Variable",-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;
    font-size:16px; line-height:1.65; -webkit-font-smoothing:antialiased;
  }
  a { color:inherit; }
  .wrap { max-width:1080px; margin:0 auto; padding:0 20px; }

  .gov-bar { background:var(--navy); color:#fff; font-size:13px; }
  .gov-bar .wrap { padding:7px 20px; display:flex; gap:8px; align-items:center; }
  .gov-bar b { font-weight:600; }
  .gov-bar span { opacity:.8; }

  header.site { background:var(--surface); border-bottom:1px solid var(--line); }
  header.site .wrap { padding:18px 20px; display:flex; align-items:baseline; gap:12px; flex-wrap:wrap; }
  .logo { font-size:19px; font-weight:700; color:var(--ink); letter-spacing:-.01em; }
  .logo em { color:var(--navy); font-style:normal; }
  .tag { font-size:13.5px; color:var(--muted); }

  .hero { background:var(--surface); border-bottom:1px solid var(--line); }
  .hero .wrap { padding:40px 20px 30px; }
  .eyebrow { font-size:13px; font-weight:600; color:var(--accent); letter-spacing:.02em; }
  .hero h1 { margin:.4em 0 .1em; font-size:clamp(26px,4vw,38px); line-height:1.25;
    font-weight:700; color:var(--ink); letter-spacing:-.02em; max-width:18em; }
  .hero h1 strong { color:var(--navy); }
  .hero p.lead { margin:8px 0 0; color:var(--muted); font-size:15.5px; max-width:40em; }
  .stats { display:flex; gap:34px; margin-top:26px; flex-wrap:wrap; }
  .stat .n { font-size:30px; font-weight:700; color:var(--navy); font-variant-numeric:tabular-nums; line-height:1; }
  .stat .l { font-size:13px; color:var(--muted); margin-top:5px; }

  .toolbar { position:sticky; top:0; z-index:5; background:rgba(246,248,251,.92);
    backdrop-filter:saturate(160%) blur(6px); border-bottom:1px solid var(--line); }
  .toolbar .wrap { padding:13px 20px; display:flex; gap:12px; align-items:center; flex-wrap:wrap; }
  .chips-row { display:flex; gap:8px; flex-wrap:wrap; flex:1; min-width:240px; }
  .fchip { display:inline-flex; align-items:center; gap:6px; cursor:pointer;
    border:1px solid var(--line); background:var(--surface); color:var(--body);
    border-radius:999px; padding:6px 13px; font-size:13.5px; font-family:inherit;
    transition:border-color .15s, color .15s, background .15s; }
  .fchip i { width:9px; height:9px; border-radius:50%; display:inline-block; }
  .fchip b { font-weight:600; color:var(--muted); font-variant-numeric:tabular-nums; }
  .fchip:hover { border-color:#C7CDD6; }
  .fchip.active { background:var(--navy); border-color:var(--navy); color:#fff; }
  .fchip.active b { color:rgba(255,255,255,.75); }
  .search { position:relative; }
  .search input { border:1px solid var(--line); border-radius:999px; padding:8px 15px 8px 36px;
    font-size:14px; font-family:inherit; width:220px; background:var(--surface); color:var(--ink);
    outline:none; transition:border-color .15s, box-shadow .15s; }
  .search input:focus { border-color:var(--navy); box-shadow:0 0 0 3px rgba(31,56,100,.12); }
  .search svg { position:absolute; left:13px; top:50%; transform:translateY(-50%); color:var(--muted); }

  main .wrap { padding:28px 20px 10px; }
  .grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(330px,1fr)); gap:16px; align-items:stretch; }
  .card { background:var(--surface); border:1px solid var(--line); border-left:4px solid var(--fc);
    border-radius:12px; padding:18px 19px; display:flex; flex-direction:column;
    transition:box-shadow .18s, transform .18s; }
  .card:hover { box-shadow:0 6px 22px -8px rgba(22,36,63,.18); transform:translateY(-2px); }
  .card-head { display:flex; align-items:center; justify-content:space-between; gap:10px; }
  .field-tag { font-size:12.5px; font-weight:600; }
  .meta { font-size:12.5px; color:var(--muted); text-align:right; }
  .card-title { margin:10px 0 0; }
  .title-btn { all:unset; cursor:pointer; font-size:17px; font-weight:600; line-height:1.4;
    letter-spacing:-.01em; color:var(--ink); }
  .title-btn:hover, .title-btn:focus-visible { color:var(--navy); text-decoration:underline; text-underline-offset:3px; }
  .chips { display:flex; flex-wrap:wrap; gap:6px; margin:11px 0 12px; }
  .chip { font-size:12px; background:#EEF2F8; color:#3A4862; border-radius:6px; padding:3px 8px; }
  .chip-more { background:transparent; color:var(--muted); }
  .summary { margin:0; font-size:14.5px; color:var(--body); line-height:1.62;
    display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden; }
  .card-foot { margin-top:auto; padding-top:14px; display:flex; align-items:center; justify-content:space-between; gap:10px; }
  .detail-link { all:unset; cursor:pointer; font-size:13.5px; font-weight:600; color:var(--accent); }
  .detail-link:hover, .detail-link:focus-visible { text-decoration:underline; text-underline-offset:3px; }
  .card-foot .ext { font-size:12.5px; color:var(--muted); text-decoration:none; }
  .card-foot .ext:hover { color:var(--navy); text-decoration:underline; text-underline-offset:3px; }
  .empty { color:var(--muted); padding:40px 0; text-align:center; }
  .noresult { display:none; color:var(--muted); padding:40px 0; text-align:center; }

  /* 팝업(모달) */
  .modal { position:fixed; inset:0; z-index:50; display:none; }
  .modal.open { display:block; }
  .modal-backdrop { position:absolute; inset:0; background:rgba(16,36,63,.45); }
  .modal-panel { position:relative; max-width:680px; margin:5vh auto; background:var(--surface);
    border-radius:16px; max-height:88vh; overflow:auto; padding:30px 30px 28px;
    box-shadow:0 24px 60px -20px rgba(16,36,63,.55); animation:pop .2s ease; }
  .modal-close { position:absolute; top:12px; right:14px; border:none; background:transparent;
    font-size:26px; line-height:1; color:var(--muted); cursor:pointer; padding:4px 8px; border-radius:8px; }
  .modal-close:hover { background:#F0F2F6; color:var(--ink); }
  .m-field { font-size:13px; font-weight:600; }
  .m-title { margin:6px 0 4px; font-size:22px; font-weight:700; color:var(--ink); line-height:1.35; letter-spacing:-.01em; padding-right:30px; }
  .m-meta { font-size:13.5px; color:var(--muted); }
  .m-sec { margin-top:22px; }
  .m-sec h4 { margin:0 0 8px; font-size:13px; font-weight:700; color:var(--navy);
    letter-spacing:.01em; padding-bottom:7px; border-bottom:1px solid var(--line); }
  .m-sec p { margin:0; font-size:15px; line-height:1.72; color:var(--body); white-space:pre-line; }
  .m-chips { display:flex; flex-wrap:wrap; gap:6px; max-height:170px; overflow:auto; padding:2px; }
  .m-arts { margin:0; padding-left:18px; }
  .m-arts li { font-size:14px; color:var(--body); line-height:1.7; }
  .m-none { color:var(--muted); font-size:14px; }
  .m-ext { display:inline-block; margin-top:24px; background:var(--navy); color:#fff;
    text-decoration:none; font-size:14px; font-weight:600; padding:11px 18px; border-radius:9px; }
  .m-ext:hover { background:#16294e; }

  footer { margin-top:36px; border-top:1px solid var(--line); background:var(--surface); }
  footer .wrap { padding:24px 20px; font-size:12.5px; color:var(--muted); line-height:1.7; }
  footer b { color:var(--body); font-weight:600; }

  @media (max-width:560px) {
    .grid { grid-template-columns:1fr; }
    .search, .search input { width:100%; }
    .modal-panel { margin:0; min-height:100vh; border-radius:0; max-height:100vh; }
  }
  .card { opacity:0; animation:rise .4s ease forwards; }
  @keyframes rise { to { opacity:1; transform:translateY(0); } }
  @keyframes pop { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
  @media (prefers-reduced-motion:reduce) {
    .card, .modal-panel { animation:none; opacity:1; } html { scroll-behavior:auto; }
    .card { transition:none; }
  }
</style>
</head>
<body>
  <div class="gov-bar"><div class="wrap"><b>한국산업인력공단</b><span>· 국가기술자격 법령 모니터링</span></div></div>

  <header class="site"><div class="wrap">
    <span class="logo">자격증 <em>법령 네비게이터</em></span>
    <span class="tag">내 자격증과 관련해 바뀐 법, 쉬운 말로</span>
  </div></header>

  <section class="hero"><div class="wrap">
    <div class="eyebrow">@@MONTH_LABEL@@ 업데이트</div>
    <h1>이번 달, 당신의 자격증과 관련해<br><strong>@@TOTAL_LAWS@@건</strong>의 법령이 바뀌었습니다.</h1>
    <p class="lead">매일 새벽 국가법령정보센터를 자동으로 살펴, 국가기술자격과 관련된 제·개정 법령만 골라 쉬운 말로 정리합니다. 카드를 누르면 전체 분석을 볼 수 있어요.</p>
    <div class="stats">
      <div class="stat"><div class="n">@@TOTAL_LAWS@@</div><div class="l">관련 법령</div></div>
      <div class="stat"><div class="n">@@TOTAL_CERTS@@</div><div class="l">연관 자격종목</div></div>
      <div class="stat"><div class="n">@@FIELD_N@@</div><div class="l">자격 분야</div></div>
    </div>
  </div></section>

  <div class="toolbar"><div class="wrap">
    <div class="chips-row">@@CHIPS@@</div>
    <div class="search">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
      <input id="q" type="search" placeholder="법령·자격·내용 검색" aria-label="검색">
    </div>
  </div></div>

  <main><div class="wrap">
    <div class="grid" id="grid">@@CARDS@@</div>
    <p class="noresult" id="noresult">조건에 맞는 법령이 없습니다.</p>
  </div></main>

  <footer><div class="wrap">
    <b>안내</b> · 이 페이지는 AI가 법령 원문을 국민 눈높이로 요약한 자동 생성물입니다.
    정확한 법적 효력은 반드시 <a href="https://www.law.go.kr" target="_blank" rel="noopener" style="color:var(--accent)">국가법령정보센터</a> 원문을 확인하세요.
    출처: 국가법령정보센터 · 워크넷 &nbsp;|&nbsp; 생성일 @@BUILT_AT@@ &nbsp;|&nbsp; 한국산업인력공단 실증(PoC)
  </div></footer>

  <!-- 팝업(모달) -->
  <div class="modal" id="modal" aria-hidden="true" role="dialog" aria-modal="true">
    <div class="modal-backdrop"></div>
    <div class="modal-panel">
      <button class="modal-close" aria-label="닫기">&times;</button>
      <div id="m-body"></div>
    </div>
  </div>

<script>
  var LAWS = @@LAWS_JSON@@;

  // ── 분야 필터 + 검색 ──
  var grid = document.getElementById('grid');
  var cards = Array.prototype.slice.call(grid.querySelectorAll('.card'));
  var fchips = Array.prototype.slice.call(document.querySelectorAll('.fchip'));
  var q = document.getElementById('q');
  var noresult = document.getElementById('noresult');
  var curField = '전체';
  function applyFilter() {
    var term = (q.value || '').trim().toLowerCase();
    var shown = 0;
    cards.forEach(function(c) {
      var on = (curField === '전체' || c.dataset.field === curField)
            && (!term || c.dataset.q.indexOf(term) !== -1);
      c.style.display = on ? '' : 'none';
      if (on) shown++;
    });
    noresult.style.display = shown ? 'none' : 'block';
  }
  fchips.forEach(function(ch) {
    ch.addEventListener('click', function() {
      fchips.forEach(function(x) { x.classList.remove('active'); });
      ch.classList.add('active');
      curField = ch.dataset.f;
      applyFilter();
    });
  });
  q.addEventListener('input', applyFilter);

  // ── 팝업(모달) ──
  var modal = document.getElementById('modal');
  var mbody = document.getElementById('m-body');
  function escq(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function(c) {
      return ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' })[c];
    });
  }
  function sec(title, inner) {
    return inner ? '<div class="m-sec"><h4>' + title + '</h4>' + inner + '</div>' : '';
  }
  function openModal(i) {
    var d = LAWS[i];
    if (!d) return;
    var certs = (d.certs || []).map(function(c) { return '<span class="chip">' + escq(c) + '</span>'; }).join('');
    var arts  = (d.articles || []).map(function(a) { return '<li>' + escq(a) + '</li>'; }).join('');
    mbody.innerHTML =
        '<div class="m-field" style="color:' + escq(d.color) + '">' + escq(d.field) + '</div>'
      + '<h2 class="m-title">' + escq(d.law) + '</h2>'
      + '<div class="m-meta">' + escq(d.meta) + '</div>'
      + sec('주요 제·개정 내용', d.summary_main ? '<p>' + escq(d.summary_main) + '</p>' : '')
      + sec('자격증 활용 분석', d.summary_use ? '<p>' + escq(d.summary_use) + '</p>' : '')
      + sec('관련 자격종목 (' + ((d.certs || []).length) + '개)', certs ? '<div class="m-chips">' + certs + '</div>' : '<p class="m-none">없음</p>')
      + sec('근거 조문', arts ? '<ul class="m-arts">' + arts + '</ul>' : '<p class="m-none">표기된 조문 없음</p>')
      + '<a class="m-ext" href="' + escq(d.url) + '" target="_blank" rel="noopener">법제처에서 원문 보기 →</a>';
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';
  }
  function closeModal() {
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
  }
  grid.addEventListener('click', function(e) {
    var t = e.target.closest('.title-btn, .detail-link');
    if (!t) return;
    var card = t.closest('.card');
    if (card) openModal(parseInt(card.dataset.i, 10));
  });
  modal.addEventListener('click', function(e) {
    if (e.target.classList.contains('modal-backdrop') || e.target.closest('.modal-close')) closeModal();
  });
  document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closeModal(); });
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
