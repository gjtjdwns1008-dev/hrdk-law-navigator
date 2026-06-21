# 자격증 법령 네비게이터 (hrdk-law-navigator)

내 자격증과 관련해 **최근 바뀐 법령**을, AI가 정리한 쉬운 말로 보여주는 **대국민 공개 웹사이트**입니다.
한국산업인력공단 법령 모니터링 시스템의 데이터를 자동으로 받아 매일 새 화면을 만들어 띄웁니다.

---

## 1. 이게 뭐고, 어떻게 도나요? (한 장 요약)

- **상시 서버가 필요 없습니다.** 내 PC를 켜둘 필요도 없습니다.
- 정해진 시각에 깃허브가 잠깐 컴퓨터를 빌려줘서, 아래 일을 자동으로 합니다.
- 흐름: **구글시트(모니터링 결과) 읽기 → `site_builder.py`가 `index.html` 생성 → GitHub Pages가 웹주소로 띄움.**
- 결과 주소: `https://gjtjdwns1008-dev.github.io/hrdk-law-navigator`

> 데이터는 사람이 손으로 올리는 게 아니라 **자동으로 채워집니다.** 디자인 틀만 코드에 한 번 정해두면, 내용은 매일 시트에서 새로 들어옵니다.

---

## 2. 폴더 구성 (파일은 단 4개)

```
hrdk-law-navigator/
├── site_builder.py                 ← 시트를 읽어 HTML을 만드는 본체
├── requirements.txt                ← 필요한 파이썬 부품 목록
├── .gitignore                      ← 깃허브에 올리지 말 것 목록
└── .github/workflows/build.yml     ← "자동으로 돌려라" 지시서 (이 경로가 중요!)
```

이 레포는 **다른 레포를 전혀 불러오지 않습니다.** 구글시트만 읽으면 되므로 혼자서 완결됩니다.

---

## 3. 처음 한 번만 하는 세팅 (천천히 따라오세요)

### 3-1. 깃허브에 새 레포 만들기
1. 깃허브 로그인 → 오른쪽 위 **`+` → New repository** 클릭.
2. **Repository name**: `hrdk-law-navigator` 입력.
3. **Public**(공개) 선택. (공개해도 시크릿은 안 새어나갑니다. 4-2 참고)
4. 아래 **Create repository** 클릭.

### 3-2. 파일 올리기
> ⚠️ 압축을 풀면 `.github` 폴더가 **숨김 폴더**라 안 보일 수 있습니다. 그래서 큰 파일은 끌어다 올리고, 워크플로우 파일(`build.yml`)만 따로 만드는 게 제일 안전합니다.

**(a) 일반 파일 3개 먼저 올리기**
1. 새로 만든 레포 화면에서 **Add file → Upload files** 클릭.
2. 압축을 푼 폴더 안의 **`site_builder.py`, `requirements.txt`, `.gitignore`, `README.md`** 를 끌어다 놓기.
3. 아래 **Commit changes** 클릭.

**(b) 워크플로우 파일은 직접 만들기 (제일 확실한 방법)**
1. **Add file → Create new file** 클릭.
2. 파일 이름 칸에 정확히 이렇게 입력: `.github/workflows/build.yml`
   - `/`(슬래시)를 치면 폴더가 자동으로 만들어집니다.
3. 압축파일 안 `build.yml` 내용을 **전부 복사해서 붙여넣기**.
4. 아래 **Commit changes** 클릭.

### 3-3. 시크릿 2개 등록 (시트 읽기 열쇠)
1. 레포에서 **Settings → Secrets and variables → Actions** 로 이동.
2. **New repository secret** 을 눌러 아래 2개를 등록.

| 이름 | 넣을 값 |
|------|---------|
| `GCP_SA_JSON` | 서비스계정 JSON **파일 전체 내용**. 로컬 백필 때 쓰는 **`gcp-key.json` 파일을 메모장으로 열어 전부 복사**해서 붙여넣으면 됩니다. |
| `GOOGLE_SHEET_ID` | 모니터링 마스터 시트의 **KEY**. 시트 주소 `.../d/`와 `/edit` 사이의 긴 문자열입니다. (주소를 통째로 넣어도 코드가 알아서 KEY만 뽑습니다.) |

> 이 두 값은 monitor 레포에 넣은 것과 **똑같습니다.** 단, 깃허브는 이미 저장된 시크릿 값을 다시 보여주지 않으니, monitor에서 "복사"는 안 되고 **원본(gcp-key.json, 시트 주소)에서 다시 넣어야** 합니다.
>
> ❗ 그리고 그 서비스계정 이메일이 해당 **구글시트에 '공유'(뷰어 이상)** 되어 있어야 읽을 수 있습니다. (monitor가 잘 돌고 있다면 이미 공유돼 있습니다.)

### 3-4. Pages 켜기 (이걸 빠뜨리면 사이트가 안 떠요)
1. 레포에서 **Settings → Pages** 로 이동.
2. **Source** 를 **`GitHub Actions`** 로 선택.
3. 끝. (저장 버튼 없이 선택만 하면 됩니다.)

---

## 4. 실행하고 결과 보기

1. 레포 상단 **Actions** 탭 클릭.
2. 왼쪽에서 **Build & Deploy navigator site** 선택.
3. 오른쪽 **Run workflow → Run workflow** (초록 버튼) 클릭 → 수동으로 한 번 돌립니다.
4. 1~2분 뒤 **초록색 체크(✓)** 가 뜨면 성공. (빨간 X면 6번 문제해결 참고)
5. 초록불 난 실행을 클릭 → **deploy** 칸의 주소(`...github.io/hrdk-law-navigator`)를 누르면 완성된 사이트가 열립니다.

이후로는 매일 자동으로 갱신되고, 언제든 위 3번처럼 수동으로도 돌릴 수 있습니다.

---

## 5. 자주 바꾸게 될 것들 (site_builder.py 안)

- **읽을 시트 탭**: 파일 위쪽 `SOURCE_WORKSHEET` (기본 `"연관 높은 법령"`).
- **컬럼 이름**: 시트 헤더가 바뀌면 `COL = { ... }` 부분만 고치면 됩니다.
- **색·문구·디자인**: 파일 아래쪽 `PAGE = """..."""` 안의 CSS와 문구.
- **보여줄 달**: 기본은 데이터의 가장 최근 달. 특정 달만 보려면 워크플로우의 `TARGET_MONTH`에 `202606` 같은 값을 넣으면 됩니다.

---

## 6. 문제 해결 (빨간 X가 떴을 때)

- **`KeyError: 'GCP_SA_JSON'`** → 시크릿을 안 넣었거나 이름이 틀림 (3-3 다시).
- **`SpreadsheetNotFound` / 권한 오류** → 서비스계정이 그 시트에 공유 안 됨, 또는 `GOOGLE_SHEET_ID` 값이 틀림.
- **`WorksheetNotFound`** → `SOURCE_WORKSHEET` 이름이 실제 시트 탭 이름과 다름.
- **워크플로우는 초록인데 사이트가 404** → Pages Source를 `GitHub Actions`로 안 바꿈 (3-4).

해결이 안 되면 Actions 실행 로그(빨간 단계)를 캡처해서 물어보세요.

---

## 7. (선택) 내 컴퓨터에서 미리보기

깃허브에 올리기 전에 화면만 미리 보고 싶을 때:

```bash
pip install pandas openpyxl gspread
# 시트 export(xlsx)를 받아두고, 그 경로/시트이름을 넣어 실행
LOCAL_XLSX="./Master_DB.xlsx" LOCAL_SHEET="연관 높은 법령" python site_builder.py
# → dist/index.html 이 생깁니다. 더블클릭해서 브라우저로 열어 확인.
```

---

*이 사이트는 AI가 법령을 국민 눈높이로 요약한 자동 생성물입니다. 정확한 법적 효력은 국가법령정보센터 원문을 확인하세요.*
