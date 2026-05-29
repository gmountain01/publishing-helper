#!/usr/bin/env python3
"""
YES24 베스트셀러 자동 분석 리포트 생성기 v2
- Google Drive 공개 폴더에서 일별 엑셀 파일 감지
- 새 파일만 다운로드 → data/yes24/archive.json 누적
- Claude API로 시장 분석 리포트 생성
- data/reports/yes24_weekly.md 덮어쓰기
"""
import io
import json
import os
import re
import sys
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime, timezone

# ── 설정 ──
DRIVE_FOLDER_ID = "1hGsZv7zT6MmFdq2Ouiwrq4Ee72zg1o4O"

# Windows cp949 인코딩 에러 방지
if sys.stdout.encoding and sys.stdout.encoding.lower().startswith("cp"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
YES24_DIR = os.path.join(SCRIPT_DIR, "..", "data", "yes24")
ARCHIVE_PATH = os.path.join(YES24_DIR, "archive.json")
REPORTS_DIR = os.path.join(SCRIPT_DIR, "..", "data", "reports")
REPORT_PATH = os.path.join(REPORTS_DIR, "yes24_weekly.md")
TODAY = datetime.now(timezone.utc).strftime("%Y-%m-%d")


# ══════════════════════════════════════════════════════
# 1. Google Drive 폴더에서 파일 목록 가져오기
# ══════════════════════════════════════════════════════

def list_drive_files() -> list[dict]:
    """공개 Drive 폴더의 엑셀 파일 목록을 가져온다."""
    print("📂 Google Drive 폴더 스캔...")
    url = f"https://drive.google.com/embeddedfolderview?id={DRIVE_FOLDER_ID}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            html = r.read().decode("utf-8", "replace")
    except Exception as e:
        print(f"  ⚠ 폴더 접근 실패: {e}", file=sys.stderr)
        return []

    # HTML에서 파일 ID와 이름 추출
    ids = re.findall(r'/file/d/([a-zA-Z0-9_-]+)', html)
    names = re.findall(r'class="flip-entry-title">(.*?)<', html)

    files = []
    for fid, fname in zip(ids, names):
        # 날짜 추출: 20260101_yes24... → 2026-01-01
        m = re.match(r"(\d{4})(\d{2})(\d{2})", fname)
        if not m or not fname.endswith(".xlsx"):
            continue
        date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        files.append({"id": fid, "name": fname, "date": date})

    files.sort(key=lambda f: f["date"])
    print(f"   {len(files)}개 엑셀 파일 발견 ({files[0]['date']} ~ {files[-1]['date']})" if files else "   파일 없음")
    return files


# ══════════════════════════════════════════════════════
# 2. 엑셀 파일 다운로드 + 파싱 (순수 Python, 외부 라이브러리 없음)
# ══════════════════════════════════════════════════════

def download_xlsx(file_id: str) -> bytes | None:
    """Drive 파일을 다운로드한다."""
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.read()
    except Exception as e:
        print(f"  ⚠ 다운로드 실패: {e}", file=sys.stderr)
        return None


def parse_xlsx(data: bytes) -> list[dict]:
    """xlsx 바이너리를 순수 Python으로 파싱한다."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return []

    # shared strings 로드
    strings = []
    if "xl/sharedStrings.xml" in zf.namelist():
        ss_xml = zf.read("xl/sharedStrings.xml")
        ss_root = ET.fromstring(ss_xml)
        ns = re.match(r"\{.*\}", ss_root.tag)
        ns = ns.group(0) if ns else ""
        for si in ss_root.findall(f".//{ns}si"):
            texts = []
            for t in si.findall(f".//{ns}t"):
                if t.text:
                    texts.append(t.text)
            strings.append("".join(texts))

    # sheet1 파싱
    sheet_path = "xl/worksheets/sheet1.xml"
    if sheet_path not in zf.namelist():
        return []

    sheet_xml = zf.read(sheet_path)
    sheet_root = ET.fromstring(sheet_xml)
    ns = re.match(r"\{.*\}", sheet_root.tag)
    ns = ns.group(0) if ns else ""

    rows_data = []
    for row_el in sheet_root.findall(f".//{ns}row"):
        cells = {}
        for c_el in row_el.findall(f"{ns}c"):
            ref = c_el.get("r", "")
            col = re.match(r"([A-Z]+)", ref)
            if not col:
                continue
            col = col.group(1)
            t = c_el.get("t", "")
            v_el = c_el.find(f"{ns}v")
            val = v_el.text if v_el is not None else ""

            if t == "s" and val.isdigit():
                idx = int(val)
                val = strings[idx] if idx < len(strings) else val
            cells[col] = val
        if cells:
            rows_data.append(cells)

    if not rows_data:
        return []

    # 헤더 감지 (첫 행)
    header_row = rows_data[0]
    col_map = {}
    for col, val in header_row.items():
        vl = str(val).strip().lower()
        if vl in ("순위", "rank"):
            col_map["rank"] = col
        elif vl in ("상품명", "제목", "도서명", "도서 제목"):
            col_map["title"] = col
        elif vl in ("저자", "작가"):
            col_map["author"] = col
        elif vl in ("출판사",):
            col_map["publisher"] = col
        elif "가격" in vl or "판매가" in vl or "정가" in vl:
            col_map["price"] = col
        elif vl in ("isbn",):
            col_map["isbn"] = col

    # 헤더 없으면 위치 기반 추정
    if "title" not in col_map:
        # 열이 A, B, C... 순서대로 순위, 상품명, ... 일 가능성
        cols = sorted(header_row.keys())
        if len(cols) >= 4:
            col_map = {"rank": cols[0], "title": cols[1], "author": cols[2], "publisher": cols[3]}
            if len(cols) >= 5:
                col_map["price"] = cols[4]
            # 첫 행도 데이터일 수 있음
            start = 0
        else:
            return []
    else:
        start = 1

    items = []
    for row in rows_data[start:]:
        title = str(row.get(col_map.get("title", ""), "")).strip()
        if not title or len(title) < 2:
            continue
        rank_val = str(row.get(col_map.get("rank", ""), "")).strip()
        items.append({
            "rank": int(float(rank_val)) if rank_val.replace(".", "").isdigit() else 0,
            "title": title,
            "author": str(row.get(col_map.get("author", ""), "")).strip(),
            "publisher": str(row.get(col_map.get("publisher", ""), "")).strip(),
            "price": str(row.get(col_map.get("price", ""), "")).strip(),
        })

    return items


# ══════════════════════════════════════════════════════
# 3. 아카이브 관리
# ══════════════════════════════════════════════════════

def load_archive() -> dict:
    if os.path.exists(ARCHIVE_PATH):
        with open(ARCHIVE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"snapshots": {}, "first_date": "", "last_date": "", "total_days": 0}


def save_archive(archive: dict):
    os.makedirs(YES24_DIR, exist_ok=True)
    with open(ARCHIVE_PATH, "w", encoding="utf-8") as f:
        json.dump(archive, f, ensure_ascii=False)

    js_path = os.path.join(YES24_DIR, "archive.js")
    with open(js_path, "w", encoding="utf-8") as f:
        f.write("window._YES24_ARCHIVE = ")
        json.dump(archive, f, ensure_ascii=False)
        f.write(";")


# ══════════════════════════════════════════════════════
# 4. 통계 계산
# ══════════════════════════════════════════════════════

def compute_stats(archive: dict) -> str:
    dates = sorted(archive["snapshots"].keys())
    all_records = []
    for d in dates:
        for item in archive["snapshots"][d]:
            all_records.append({**item, "date": d})

    unique_titles = set(r["title"] for r in all_records)
    date_range = f"{dates[0]} ~ {dates[-1]}"

    # 도서별 등장일수 + 평균 순위
    title_days = defaultdict(set)
    title_ranks = defaultdict(list)
    title_info = {}
    for r in all_records:
        t = r["title"]
        title_days[t].add(r["date"])
        if r["rank"]:
            title_ranks[t].append(r["rank"])
        if t not in title_info:
            title_info[t] = {k: r[k] for k in ("author", "publisher", "price") if k in r}

    top30 = sorted(title_days.items(), key=lambda x: -len(x[1]))[:30]

    # 출판사 점유율
    pub_days = defaultdict(int)
    for t, days in title_days.items():
        pub = title_info[t].get("publisher", "")
        if pub:
            pub_days[pub] += len(days)
    total_days_all = sum(pub_days.values()) or 1
    pub_top = sorted(pub_days.items(), key=lambda x: -x[1])[:15]

    # 카테고리 분류
    KW = {
        "AI/LLM": ["ai", "인공지능", "llm", "gpt", "클로드", "제미나이", "생성형", "딥러닝", "머신러닝"],
        "바이브코딩": ["바이브 코딩", "바이브코딩", "vibe coding"],
        "에이전트/RAG": ["에이전트", "agent", "rag", "랭체인", "langchain", "mcp"],
        "프롬프트/활용": ["프롬프트", "prompt", "챗gpt", "ai 활용", "업무 자동화", "활용법"],
        "데이터분석": ["데이터 분석", "데이터분석", "판다스", "pandas", "엑셀", "통계"],
        "프로그래밍": ["파이썬", "python", "자바", "java", "코딩", "알고리즘", "자료구조"],
        "웹/앱개발": ["웹", "리액트", "react", "flutter", "next.js", "스프링", "spring"],
        "클라우드/인프라": ["클라우드", "aws", "도커", "쿠버네티스", "devops", "azure"],
        "보안": ["보안", "해킹", "정보보안"],
        "이미지/영상AI": ["이미지 생성", "stable diffusion", "미드저니", "comfyui", "영상"],
    }
    cat_titles = defaultdict(set)
    for t in unique_titles:
        tl = t.lower()
        for cat, kws in KW.items():
            if any(k in tl for k in kws):
                cat_titles[cat].add(t)
                break

    # 최근 7일 변동
    recent_change = ""
    if len(dates) >= 14:
        last7 = set(dates[-7:])
        prev7 = set(dates[-14:-7])
        r7 = set()
        p7 = set()
        for d in last7:
            for it in archive["snapshots"].get(d, []):
                r7.add(it["title"])
        for d in prev7:
            for it in archive["snapshots"].get(d, []):
                p7.add(it["title"])
        new_in = r7 - p7
        dropped = p7 - r7
        recent_change = f"\n### 최근 7일 변동\n- 신규 진입: {len(new_in)}권\n- 이탈: {len(dropped)}권\n"
        if new_in:
            recent_change += "- 신규 주요:\n"
            for t in list(new_in)[:10]:
                info = title_info.get(t, {})
                recent_change += f"  - {t} ({info.get('publisher', '?')})\n"

    lines = [
        f"## 기초 통계",
        f"- 분석 기간: {date_range} ({len(dates)}일)",
        f"- 총 스냅샷 레코드: {len(all_records):,}건",
        f"- 고유 도서: {len(unique_titles):,}권",
        f"",
        f"### 출판사 점유율 (등장일수 기준, 상위 15)",
        f"| 순위 | 출판사 | 등장일수 | 점유율 |",
        f"|---:|--------|-------:|------:|",
    ]
    for i, (pub, days) in enumerate(pub_top, 1):
        lines.append(f"| {i} | {pub} | {days}일 | {days/total_days_all*100:.1f}% |")

    lines.append(f"\n### 베스트셀러 TOP 30 (등장일수·평균순위)")
    lines.append(f"| 순위 | 도서명 | 출판사 | 등장일수 | 평균순위 |")
    lines.append(f"|---:|--------|--------|-------:|-------:|")
    for i, (title, days) in enumerate(top30, 1):
        ud = len(days)
        avg = sum(title_ranks[title]) / len(title_ranks[title]) if title_ranks[title] else 0
        pub = title_info[title].get("publisher", "")
        lines.append(f"| {i} | {title} | {pub} | {ud}일 | {avg:.1f} |")

    lines.append(f"\n### 카테고리 분포")
    for cat, titles in sorted(cat_titles.items(), key=lambda x: -len(x[1])):
        lines.append(f"- **{cat}**: {len(titles)}권")

    if recent_change:
        lines.append(recent_change)

    return "\n".join(lines)


# ══════════════════════════════════════════════════════
# 5. Claude API
# ══════════════════════════════════════════════════════

def call_claude(stats: str, dates: list[str]) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("  ⚠ ANTHROPIC_API_KEY 미설정 — 통계만 생성", file=sys.stderr)
        return ""

    date_range = f"{dates[0]} ~ {dates[-1]}"
    prompt = f"""아래는 YES24 IT 베스트셀러의 일별 스냅샷 누적 데이터 통계입니다.
매일 200위까지의 베스트셀러를 수집하여 {len(dates)}일간 누적한 결과입니다.

{stats}

---

이 데이터를 기반으로 IT 도서 시장 분석 리포트를 작성하라:

# YES24 IT 베스트셀러 시장 분석 리포트

> 분석 기간: {date_range} ({len(dates)}일 누적)
> 생성일: {TODAY}
> 데이터: YES24 IT/모바일 일별 베스트셀러 200위 (자동 생성)

## 핵심 인사이트 (5개)
구체적 수치 포함, 출판 기획 관점에서 해석하라.

## 카테고리별 트렌드
상승/하락/안정 판별, 주목 도서와 출판사 언급.

## 출판 기회 (3~5개)
시장 공백·기회를 구체적으로 제안. 각각: 왜 기회인지, 어떤 도서, 타겟 독자.

## 경쟁 동향
주요 출판사별 최근 움직임과 전략.

## 주간 변동 요약
최근 7일 신규 진입/이탈 중심 시장 변화.

---
글쓰기 원칙: AI투 금지, 편집자 브리핑 톤, 수치 자연스럽게 녹이기, 문장 구조 다양화.
"""

    body = json.dumps({
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 4000,
        "messages": [{"role": "user", "content": prompt}]
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={"Content-Type": "application/json", "x-api-key": api_key, "anthropic-version": "2023-06-01"}
    )

    print("🤖 Claude API 분석 중...")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode("utf-8"))["content"][0]["text"]
    except Exception as e:
        print(f"  ⚠ Claude API 실패: {e}", file=sys.stderr)
        return ""


# ══════════════════════════════════════════════════════
# 6. 메인
# ══════════════════════════════════════════════════════

def main():
    os.makedirs(YES24_DIR, exist_ok=True)
    os.makedirs(REPORTS_DIR, exist_ok=True)

    # 1. Drive 폴더 파일 목록
    drive_files = list_drive_files()
    if not drive_files:
        print("❌ Drive 파일 없음 — 종료")
        return

    # 2. 아카이브 로드 + 새 파일 감지
    archive = load_archive()
    existing_dates = set(archive["snapshots"].keys())
    new_files = [f for f in drive_files if f["date"] not in existing_dates]

    if not new_files:
        print(f"⏭ 새 파일 없음 (마지막: {archive.get('last_date', '?')}) — 스킵")
        return

    print(f"🆕 새 파일 {len(new_files)}개 감지")

    # 3. 새 파일 다운로드 + 파싱 + 아카이브 병합
    added = 0
    for f in new_files:
        print(f"  📥 {f['name']}...", end=" ")
        data = download_xlsx(f["id"])
        if not data:
            print("SKIP")
            continue
        items = parse_xlsx(data)
        if not items:
            print("파싱 실패")
            continue
        archive["snapshots"][f["date"]] = items
        added += 1
        print(f"{len(items)}건")

    if added == 0:
        print("❌ 파싱 성공한 파일 없음 — 종료")
        return

    # 메타데이터 갱신
    all_dates = sorted(archive["snapshots"].keys())
    archive["first_date"] = all_dates[0]
    archive["last_date"] = all_dates[-1]
    archive["total_days"] = len(all_dates)

    save_archive(archive)
    print(f"💾 아카이브 저장: {archive['total_days']}일, {sum(len(v) for v in archive['snapshots'].values())}건")

    # 4. 통계 + 리포트
    stats = compute_stats(archive)
    ai_report = call_claude(stats, all_dates)

    if ai_report:
        report = ai_report
    else:
        report = f"# YES24 IT 베스트셀러 시장 분석 리포트\n\n> 분석 기간: {all_dates[0]} ~ {all_dates[-1]} ({len(all_dates)}일)\n> 생성일: {TODAY}\n\n{stats}\n"

    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"✅ 리포트 생성: {REPORT_PATH}")


if __name__ == "__main__":
    main()
