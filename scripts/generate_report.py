#!/usr/bin/env python3
"""
YES24 베스트셀러 자동 분석 리포트 생성기
- Google Sheets에서 YES24 데이터 fetch
- 마지막 리포트 이후 새 데이터가 있으면 분석 실행
- Claude API로 시장 분석 리포트 생성
- data/reports/yes24_weekly.md 덮어쓰기
"""
import csv
import io
import json
import os
import re
import sys
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime, timedelta

# ── 설정 ──
SHEET_URL = "https://script.google.com/macros/s/AKfycbx0PRidfgLM41CLKyM6zmaNkf9_r-a3EZGxU9qicd-a_-i8K0xGGV2XH64geJwQ6k7d/exec"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(SCRIPT_DIR, "..", "data", "reports")
REPORT_PATH = os.path.join(REPORTS_DIR, "yes24_weekly.md")
STATE_PATH = os.path.join(REPORTS_DIR, ".report_state.json")


def fetch_sheet_data() -> list[list[str]]:
    """Google Sheets Apps Script에서 데이터를 가져온다."""
    print("📊 Google Sheets 데이터 fetch...")
    req = urllib.request.Request(SHEET_URL, headers={"User-Agent": "ReportBot/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            text = resp.read().decode("utf-8", errors="replace")
    except Exception as e:
        print(f"  ⚠ fetch 실패: {e}", file=sys.stderr)
        return []

    # Apps Script JSON 응답 파싱
    try:
        data = json.loads(text)
        if isinstance(data, dict) and "rows" in data:
            return data["rows"]
        if isinstance(data, list):
            return data
    except json.JSONDecodeError:
        pass

    # CSV 폴백
    try:
        reader = csv.reader(io.StringIO(text))
        return [row for row in reader]
    except Exception:
        pass

    print("  ⚠ 데이터 파싱 실패", file=sys.stderr)
    return []


def parse_rows(rows: list[list[str]]) -> list[dict]:
    """행 데이터를 파싱하여 구조화된 레코드 목록 반환."""
    if not rows:
        return []

    # 헤더 감지
    KNOWN_PUBS = {"한빛미디어", "길벗", "제이펍", "위키북스", "에이콘출판사", "골든래빗",
                  "이지스퍼블리싱", "영진닷컴", "한빛아카데미", "생능출판사", "인사이트",
                  "커뮤니케이션북스", "앤써북", "성안당", "생능북스", "책만", "비제이퍼블릭"}

    header = rows[0] if rows[0] and any(
        h in str(rows[0]).lower() for h in ["순위", "제목", "상품명", "출판사", "rank"]
    ) else None

    col_map = {}
    if header:
        for i, h in enumerate(header):
            hl = str(h).strip().lower()
            if hl in ("순위", "rank"):
                col_map["rank"] = i
            elif hl in ("상품명", "제목", "도서명", "도서 제목", "책제목"):
                col_map["title"] = i
            elif hl in ("저자", "작가"):
                col_map["author"] = i
            elif hl in ("출판사",):
                col_map["publisher"] = i
            elif "가격" in hl or "정가" in hl:
                col_map["price"] = i
            elif "날짜" in hl or "일자" in hl or "date" in hl:
                col_map["date"] = i
        start = 1
    else:
        # 헤더 없음 → 자동 감지
        start = 0

    records = []
    for row in rows[start:]:
        if not row or len(row) < 3:
            continue
        if header and col_map:
            rec = {
                "rank": str(row[col_map["rank"]]).strip() if "rank" in col_map and col_map["rank"] < len(row) else "",
                "title": str(row[col_map["title"]]).strip() if "title" in col_map and col_map["title"] < len(row) else "",
                "author": str(row[col_map.get("author", -1)]).strip() if "author" in col_map and col_map["author"] < len(row) else "",
                "publisher": str(row[col_map.get("publisher", -1)]).strip() if "publisher" in col_map and col_map["publisher"] < len(row) else "",
                "price": str(row[col_map.get("price", -1)]).strip() if "price" in col_map and col_map["price"] < len(row) else "",
                "date": str(row[col_map.get("date", -1)]).strip() if "date" in col_map and col_map["date"] < len(row) else "",
            }
        else:
            # 위치 기반 추정
            rec = {"rank": str(row[0]).strip(), "title": "", "author": "", "publisher": "", "price": "", "date": ""}
            for i, cell in enumerate(row):
                s = str(cell).strip()
                if not rec["title"] and len(s) > 5 and len(s) < 200:
                    rec["title"] = s
                elif not rec["publisher"] and s in KNOWN_PUBS:
                    rec["publisher"] = s
                elif not rec["price"] and re.match(r"[\d,]+원?$", s):
                    rec["price"] = s
        if rec["title"]:
            records.append(rec)

    return records


def extract_dates(records: list[dict]) -> list[str]:
    """레코드에서 고유 날짜 목록 추출."""
    dates = set()
    for r in records:
        d = r.get("date", "").strip()
        if re.match(r"\d{4}[-/]\d{2}[-/]\d{2}", d):
            dates.add(d[:10].replace("/", "-"))
    return sorted(dates)


def load_state() -> dict:
    """마지막 리포트 상태 로드."""
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_date": "", "last_run": ""}


def save_state(last_date: str):
    """리포트 상태 저장."""
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump({"last_date": last_date, "last_run": datetime.now().strftime("%Y-%m-%d %H:%M")}, f)


def compute_stats(records: list[dict], dates: list[str]) -> str:
    """기초 통계를 문자열로 생성 (Claude 프롬프트에 전달)."""
    total = len(records)
    unique_titles = len(set(r["title"] for r in records))
    date_range = f"{dates[0]} ~ {dates[-1]}" if dates else "불명"

    # 출판사 분포
    pub_cnt = Counter(r["publisher"] for r in records if r["publisher"])
    pub_top = pub_cnt.most_common(15)

    # 제목 빈도 (등장일수 = 인기 지표)
    title_cnt = Counter(r["title"] for r in records if r["title"])
    title_top = title_cnt.most_common(30)

    # 카테고리 키워드 감지
    categories = defaultdict(list)
    KW = {
        "AI/LLM": ["ai", "인공지능", "llm", "gpt", "클로드", "제미나이", "생성형", "딥러닝", "머신러닝", "트랜스포머"],
        "바이브코딩": ["바이브 코딩", "바이브코딩", "vibe coding"],
        "에이전트": ["에이전트", "agent", "rag", "랭체인", "langchain"],
        "프롬프트/활용": ["프롬프트", "prompt", "챗gpt 활용", "ai 활용", "업무 자동화"],
        "데이터분석": ["데이터 분석", "데이터분석", "파이썬 데이터", "판다스", "pandas", "엑셀"],
        "프로그래밍": ["파이썬", "python", "자바", "java", "코딩", "알고리즘", "자료구조"],
        "웹/앱": ["웹", "리액트", "react", "flutter", "앱", "next.js", "스프링"],
        "클라우드/인프라": ["클라우드", "aws", "도커", "쿠버네티스", "kubernetes", "devops"],
        "보안": ["보안", "해킹", "정보보안", "security"],
        "이미지/영상AI": ["이미지 생성", "stable diffusion", "미드저니", "영상 ai", "sora", "comfyui"],
    }
    for r in records:
        t = r["title"].lower()
        for cat, kws in KW.items():
            if any(k in t for k in kws):
                categories[cat].append(r["title"])
                break

    # 최근 7일 vs 이전 비교
    recent = ""
    if len(dates) >= 14:
        last7 = set(dates[-7:])
        prev7 = set(dates[-14:-7])
        r7_titles = set(r["title"] for r in records if r.get("date", "")[:10] in last7)
        p7_titles = set(r["title"] for r in records if r.get("date", "")[:10] in prev7)
        new_entries = r7_titles - p7_titles
        dropped = p7_titles - r7_titles
        recent = f"\n### 최근 7일 변동\n- 신규 진입: {len(new_entries)}권\n- 이탈: {len(dropped)}권\n"
        if new_entries:
            recent += "- 신규 주요: " + ", ".join(list(new_entries)[:10]) + "\n"

    lines = [
        f"## 기초 통계",
        f"- 분석 기간: {date_range} ({len(dates)}일)",
        f"- 총 레코드: {total:,}건",
        f"- 고유 도서: {unique_titles:,}권",
        f"",
        f"### 출판사 상위 15",
    ]
    for i, (pub, cnt) in enumerate(pub_top, 1):
        lines.append(f"| {i} | {pub} | {cnt}건 |")

    lines.append(f"\n### 베스트셀러 TOP 30 (등장일수)")
    for i, (title, cnt) in enumerate(title_top, 1):
        lines.append(f"| {i} | {title} | {cnt}일 |")

    lines.append(f"\n### 카테고리 분포")
    for cat, titles in sorted(categories.items(), key=lambda x: -len(x[1])):
        lines.append(f"- {cat}: {len(set(titles))}권")

    if recent:
        lines.append(recent)

    return "\n".join(lines)


def call_claude(stats: str, dates: list[str]) -> str:
    """Claude API로 시장 분석 리포트 생성."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("  ⚠ ANTHROPIC_API_KEY 미설정 — 통계 리포트만 생성", file=sys.stderr)
        return ""

    date_range = f"{dates[0]} ~ {dates[-1]}" if dates else "불명"
    today = datetime.now().strftime("%Y-%m-%d")

    prompt = f"""아래는 YES24 IT 베스트셀러 데이터의 기초 통계입니다.
이 데이터를 기반으로 IT 도서 시장 분석 리포트를 작성해주세요.

{stats}

---

다음 형식으로 작성하세요:

# YES24 IT 베스트셀러 시장 분석 리포트

> 분석 기간: {date_range}
> 생성일: {today}
> 데이터: YES24 IT/모바일 일별 베스트셀러 (자동 생성)

## 핵심 인사이트 (5개)
각 인사이트는 구체적 수치를 포함하고, 출판 기획 관점에서 의미를 해석하세요.

## 카테고리별 트렌드
상승/하락/안정 트렌드를 판별하고, 각 카테고리에서 주목할 도서와 출판사를 언급하세요.

## 출판 기회 (3~5개)
데이터에서 발견한 시장 공백이나 기회를 구체적으로 제안하세요.
각 기회에 대해: 왜 기회인지, 어떤 도서를 만들면 좋을지, 타겟 독자는 누구인지.

## 경쟁 동향
주요 출판사별 최근 움직임과 전략을 분석하세요.

## 주간 변동 요약
최근 7일 신규 진입/이탈 도서를 중심으로 시장 변화를 설명하세요.

---
글쓰기 원칙:
- AI투 문장 금지 (혁신적인, ~할 수 있습니다, 주목할 만합니다 등)
- 편집자가 동료에게 브리핑하는 톤
- 구체적 수치를 자연스럽게 녹이기
- 문장 구조 다양화 (주어-서술어 단조로움 금지)
"""

    body = json.dumps({
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 4000,
        "messages": [{"role": "user", "content": prompt}]
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01"
        }
    )

    print("🤖 Claude API 분석 중...")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return result["content"][0]["text"]
    except Exception as e:
        print(f"  ⚠ Claude API 실패: {e}", file=sys.stderr)
        return ""


def main():
    os.makedirs(REPORTS_DIR, exist_ok=True)

    # 1. 데이터 fetch
    rows = fetch_sheet_data()
    if not rows:
        print("❌ 데이터 없음 — 종료")
        return

    records = parse_rows(rows)
    if not records:
        print("❌ 파싱된 레코드 없음 — 종료")
        return

    dates = extract_dates(records)
    print(f"📅 데이터 날짜: {len(dates)}일 ({dates[0] if dates else '?'} ~ {dates[-1] if dates else '?'})")

    # 2. 새 데이터 확인
    state = load_state()
    latest_date = dates[-1] if dates else ""

    if state["last_date"] and latest_date <= state["last_date"]:
        print(f"⏭ 새 데이터 없음 (마지막: {state['last_date']}) — 스킵")
        return

    print(f"🆕 새 데이터 감지! {state['last_date'] or '(첫 실행)'} → {latest_date}")

    # 3. 통계 계산
    stats = compute_stats(records, dates)

    # 4. Claude 분석 (API 키 있으면)
    ai_report = call_claude(stats, dates)

    # 5. 리포트 생성
    if ai_report:
        report = ai_report
    else:
        # Claude 없이 통계만
        date_range = f"{dates[0]} ~ {dates[-1]}" if dates else "불명"
        today = datetime.now().strftime("%Y-%m-%d")
        report = f"""# YES24 IT 베스트셀러 시장 분석 리포트

> 분석 기간: {date_range}
> 생성일: {today}
> 데이터: YES24 IT/모바일 일별 베스트셀러 (자동 생성)

{stats}
"""

    # 6. 리포트 메타데이터 (build_reports.py용)
    meta_header = ""
    if not ai_report:
        meta_header = ""  # Claude 리포트에는 이미 헤더 포함

    # 7. 저장
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(report)

    save_state(latest_date)
    print(f"✅ 리포트 생성 완료: {REPORT_PATH}")
    print(f"   최신 데이터 날짜: {latest_date}")


if __name__ == "__main__":
    main()
