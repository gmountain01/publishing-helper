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

TOPIC_KW = {
    "AI/LLM 일반": ["ai", "인공지능", "llm", "gpt", "클로드", "제미나이", "생성형", "챗gpt", "오픈ai", "openai", "claude", "gemini"],
    "바이브코딩/노코드": ["바이브 코딩", "바이브코딩", "vibe coding", "노코드", "로우코드"],
    "AI 에이전트/RAG": ["에이전트", "agent", "rag", "랭체인", "langchain", "langgraph", "mcp", "에이전틱"],
    "프롬프트/활용": ["프롬프트", "prompt", "ai 활용", "업무 자동화", "활용법", "활용 가이드"],
    "이미지/영상 AI": ["이미지 생성", "stable diffusion", "미드저니", "comfyui", "영상 ai", "sora", "캡컷", "영상 편집", "ai 영상", "ai 쇼츠"],
    "데이터분석/사이언스": ["데이터 분석", "데이터분석", "판다스", "pandas", "데이터 사이언스", "통계", "r 프로그래밍"],
    "딥러닝/머신러닝": ["딥러닝", "머신러닝", "deep learning", "machine learning", "텐서플로", "파이토치", "트랜스포머"],
    "파이썬": ["파이썬", "python", "점프 투 파이썬"],
    "웹개발": ["웹", "리액트", "react", "next.js", "스프링", "spring", "html", "css", "자바스크립트", "타입스크립트"],
    "앱개발/모바일": ["앱 개발", "flutter", "swift", "코틀린", "안드로이드", "ios"],
    "컴퓨터과학/기초": ["컴퓨터 개론", "자료구조", "알고리즘", "운영체제", "컴퓨팅", "이산수학", "c언어", "c++", "자바 프로그래밍"],
    "클라우드/DevOps": ["클라우드", "aws", "azure", "도커", "쿠버네티스", "kubernetes", "devops", "terraform"],
    "보안/해킹": ["보안", "해킹", "정보보안", "사이버", "모의침투"],
    "엑셀/오피스": ["엑셀", "excel", "파워포인트", "한글", "오피스", "워드"],
    "게임개발": ["게임 개발", "유니티", "unity", "언리얼", "unreal", "게임 프로그래밍"],
    "비전공자/교양": ["비전공자", "교양", "코딩 입문", "처음 배우는", "쉽게 배우는", "혼자 공부"],
    "자격증/취업": ["자격증", "정보처리", "취업", "코딩 테스트", "코딩테스트"],
    "로봇/IoT/하드웨어": ["로봇", "아두이노", "라즈베리", "iot", "반도체", "하드웨어", "임베디드"],
    "블록체인/Web3": ["블록체인", "web3", "nft", "솔리디티", "이더리움"],
}

def _classify_topic(title: str) -> str:
    """첫 매칭 주제 반환 (단일 분류용)."""
    tl = title.lower()
    for topic, kws in TOPIC_KW.items():
        if any(k in tl for k in kws):
            return topic
    return "기타"

def _classify_topics_multi(title: str) -> list[str]:
    """매칭되는 모든 주제 반환 (복수 분류용)."""
    tl = title.lower()
    matched = []
    for topic, kws in TOPIC_KW.items():
        if any(k in tl for k in kws):
            matched.append(topic)
    return matched if matched else ["기타"]

def compute_stats(archive: dict) -> str:
    dates = sorted(archive["snapshots"].keys())
    all_records = []
    for d in dates:
        for item in archive["snapshots"][d]:
            all_records.append({**item, "date": d})

    unique_titles = set(r["title"] for r in all_records)
    date_range = f"{dates[0]} ~ {dates[-1]}"
    num_days = len(dates)

    # 도서별 등장일수 + 평균 순위
    title_days = defaultdict(set)
    title_ranks = defaultdict(list)
    title_info = {}
    title_first = {}
    for r in all_records:
        t = r["title"]
        title_days[t].add(r["date"])
        if r["rank"]:
            title_ranks[t].append(r["rank"])
        if t not in title_info:
            title_info[t] = {k: r[k] for k in ("author", "publisher", "price") if k in r}
        if t not in title_first or r["date"] < title_first[t]:
            title_first[t] = r["date"]

    # 출판사 점유율
    pub_book_cnt = Counter(title_info[t].get("publisher", "") for t in unique_titles if title_info[t].get("publisher"))
    pub_days = defaultdict(int)
    for t, days in title_days.items():
        pub = title_info[t].get("publisher", "")
        if pub:
            pub_days[pub] += len(days)
    total_days_all = sum(pub_days.values()) or 1
    pub_top20 = sorted(pub_book_cnt.items(), key=lambda x: -x[1])[:20]

    # 가격대 분포
    price_bins = defaultdict(int)
    for t in unique_titles:
        p = title_info[t].get("price", "")
        m = re.search(r"[\d,]+", p.replace(",", ""))
        if m:
            val = int(m.group().replace(",", "")) if m.group().replace(",", "").isdigit() else 0
            if val < 15000: b = "0~15천원"
            elif val < 20000: b = "15~20천원"
            elif val < 25000: b = "20~25천원"
            elif val < 30000: b = "25~30천원"
            elif val < 35000: b = "30~35천원"
            elif val < 40000: b = "35~40천원"
            elif val < 50000: b = "40~50천원"
            else: b = "50천원 이상"
            price_bins[b] += 1

    # 월별 신규 진입
    monthly_new = defaultdict(set)
    seen = set()
    for d in dates:
        month = d[:7]
        for item in archive["snapshots"][d]:
            if item["title"] not in seen:
                seen.add(item["title"])
                monthly_new[month].add(item["title"])

    # ── 주제별 분석 (복수 분류: 한 도서가 여러 주제에 포함) ──
    topic_titles = defaultdict(set)
    for t in unique_titles:
        for topic in _classify_topics_multi(t):
            topic_titles[topic].add(t)

    topic_stats = []
    recent30 = set(dates[-30:]) if len(dates) >= 30 else set(dates)
    prev30 = set(dates[-60:-30]) if len(dates) >= 60 else set()
    for topic, titles in sorted(topic_titles.items(), key=lambda x: -len(x[1])):
        ranks = []
        day_counts = []
        r30 = set()
        p30 = set()
        for t in titles:
            ranks.extend(title_ranks[t])
            day_counts.append(len(title_days[t]))
            for d in title_days[t]:
                if d in recent30: r30.add(t)
                if d in prev30: p30.add(t)
        avg_rank = sum(ranks) / len(ranks) if ranks else 0
        avg_days = sum(day_counts) / len(day_counts) if day_counts else 0
        if prev30:
            trend = "📈 상승" if len(r30) > len(p30) * 1.1 else ("📉 하락" if len(r30) < len(p30) * 0.9 else "➡️ 유지")
        else:
            trend = "➡️ 유지"
        topic_stats.append((topic, len(titles), avg_rank, avg_days, len(r30), trend))

    # 주제별 출판사 점유
    topic_pubs = {}
    for topic, titles in topic_titles.items():
        pc = Counter(title_info[t].get("publisher", "") for t in titles if title_info[t].get("publisher"))
        topic_pubs[topic] = pc.most_common(5)

    # ── 트렌드 분석 ──
    # 급상승 도서 (최근 30일 vs 이전) — 현재 상위권 도서 중 상승폭 큰 순
    surge_books = []
    if len(dates) >= 60:
        r30_set = set(dates[-30:])
        p30_set = set(dates[-60:-30])
        for t in unique_titles:
            r_ranks = [r["rank"] for r in all_records if r["title"] == t and r["date"] in r30_set and r["rank"]]
            p_ranks = [r["rank"] for r in all_records if r["title"] == t and r["date"] in p30_set and r["rank"]]
            if r_ranks and p_ranks:
                r_avg = sum(r_ranks) / len(r_ranks)
                p_avg = sum(p_ranks) / len(p_ranks)
                if p_avg - r_avg >= 5 and r_avg <= 30:  # 현재 30위 이내 + 5순위 이상 상승
                    surge_books.append((round(r_avg), round(p_avg), round(p_avg - r_avg), t, title_info[t].get("publisher", "")))
        surge_books.sort(key=lambda x: -x[2])  # 상승폭 큰 순

    # 장기 스테디셀러 (100일+ 등장)
    steady_threshold = 100 if num_days >= 100 else max(int(num_days * 0.7), 10)
    steady = [(t, len(title_days[t]), sum(title_ranks[t]) / len(title_ranks[t]) if title_ranks[t] else 0, title_info[t].get("publisher", ""))
              for t in unique_titles if len(title_days[t]) >= steady_threshold]
    steady.sort(key=lambda x: -x[1])

    # 신규 트렌드 (최근 60일 첫 등장)
    cutoff_60 = dates[-60] if len(dates) >= 60 else dates[0]
    new_trend = []
    for t in unique_titles:
        if title_first[t] >= cutoff_60:
            best_rank = min(title_ranks[t]) if title_ranks[t] else 999
            nd = len(title_days[t])
            if nd >= 2 and best_rank <= 30:
                new_trend.append((title_first[t], best_rank, nd, t, title_info[t].get("publisher", "")))
    new_trend.sort(key=lambda x: (x[1], -x[2]))

    # ── 한빛미디어 분석 ──
    hanbit_titles = [t for t in unique_titles if "한빛" in title_info[t].get("publisher", "")]
    hanbit_topic = defaultdict(int)
    for t in hanbit_titles:
        for topic in _classify_topics_multi(t):
            hanbit_topic[topic] += 1

    hanbit_books = []
    for t in hanbit_titles:
        avg = sum(title_ranks[t]) / len(title_ranks[t]) if title_ranks[t] else 999
        nd = len(title_days[t])
        topics = ", ".join(_classify_topics_multi(t))
        hanbit_books.append((avg, nd, t, topics))
    hanbit_books.sort(key=lambda x: x[0])

    # ── 리포트 조합 ──
    L = []
    L.append(f"## 1. 기본 통계\n")
    L.append(f"- 분석 대상 파일: {num_days}개")
    L.append(f"- 총 레코드 수: {len(all_records):,}건")
    L.append(f"- 고유 도서 수: {len(unique_titles):,}권\n")

    L.append(f"### 출판사별 도서 수 (상위 20)\n")
    L.append(f"| 순위 | 출판사 | 도서 수 | 점유율(등장일수) |")
    L.append(f"|---:|--------|-------:|--------:|")
    for i, (pub, cnt) in enumerate(pub_top20, 1):
        share = pub_days.get(pub, 0) / total_days_all * 100
        L.append(f"| {i} | {pub} | {cnt}권 | {share:.1f}% |")

    if price_bins:
        L.append(f"\n### 가격대 분포\n")
        L.append(f"| 가격대 | 도서 수 | 비율 |")
        L.append(f"|--------|-------:|-----:|")
        total_priced = sum(price_bins.values()) or 1
        for b in ["0~15천원", "15~20천원", "20~25천원", "25~30천원", "30~35천원", "35~40천원", "40~50천원", "50천원 이상"]:
            if b in price_bins:
                L.append(f"| {b} | {price_bins[b]}권 | {price_bins[b]/total_priced*100:.1f}% |")

    L.append(f"\n### 월별 신규 진입 도서 수\n")
    L.append(f"| 월 | 신규 도서 |")
    L.append(f"|-----|--------:|")
    for m in sorted(monthly_new.keys()):
        L.append(f"| {m} | {len(monthly_new[m])}권 |")

    # ── 2. 주제별 분석 ──
    L.append(f"\n## 2. 주제별 분석\n")
    L.append(f"| 주제 | 도서 수 | 평균순위 | 평균등장일 | 최근30일 도서 | 트렌드 |")
    L.append(f"|------|-------:|--------:|--------:|----------:|-----:|")
    for topic, cnt, avg_r, avg_d, r30_cnt, trend in topic_stats:
        L.append(f"| {topic} | {cnt}권 | {avg_r:.1f} | {avg_d:.0f}일 | {r30_cnt}권 | {trend} |")

    L.append(f"\n### 주제별 주요 출판사 점유\n")
    for topic in sorted(topic_pubs.keys()):
        pubs = topic_pubs[topic]
        if pubs:
            pub_str = ", ".join(f"{p}({c})" for p, c in pubs)
            L.append(f"- **{topic}**: {pub_str}")

    # ── 3. 트렌드 분석 ──
    L.append(f"\n## 3. 트렌드 분석\n")
    if surge_books:
        L.append(f"### 급상승 도서 (최근 30일)\n")
        L.append(f"| 현재 순위 | 이전 순위 | 상승폭 | 도서명 | 출판사 |")
        L.append(f"|--------:|--------:|------:|--------|--------|")
        for r_avg, p_avg, diff, t, pub in surge_books[:15]:
            L.append(f"| {r_avg}위 | {p_avg}위 | +{diff} | {t} | {pub} |")

    if steady:
        L.append(f"\n### 장기 스테디셀러 ({steady_threshold}일 이상 등장)\n")
        L.append(f"| 등장일수 | 평균순위 | 도서명 | 출판사 |")
        L.append(f"|-------:|-------:|--------|--------|")
        for t, nd, avg, pub in steady[:20]:
            L.append(f"| {nd}일 | {avg:.0f}위 | {t} | {pub} |")

    if new_trend:
        L.append(f"\n### 신규 트렌드 (최근 60일 내 첫 등장)\n")
        L.append(f"| 첫등장 | 최고순위 | 등장일 | 도서명 | 출판사 |")
        L.append(f"|-------|-------:|------:|--------|--------|")
        for first, best, nd, t, pub in new_trend[:20]:
            L.append(f"| {first} | {best}위 | {nd}일 | {t} | {pub} |")

    # ── 4. 한빛미디어 분석 ──
    L.append(f"\n## 4. 한빛미디어 분석\n")
    L.append(f"한빛 계열(한빛미디어, 한빛아카데미) 베스트셀러 진입 도서: **{len(hanbit_titles)}권**\n")

    L.append(f"### 한빛미디어 베스트셀러 도서\n")
    L.append(f"| 평균순위 | 등장일 | 도서명 | 주제 |")
    L.append(f"|-------:|------:|--------|------|")
    for avg, nd, t, topic in hanbit_books[:25]:
        L.append(f"| {avg:.0f}위 | {nd}일 | {t} | {topic} |")

    L.append(f"\n### 한빛 주제 분포\n")
    L.append(f"| 주제 | 한빛 도서 수 | 전체 시장 | 점유율 |")
    L.append(f"|------|----------:|--------:|------:|")
    for topic in sorted(topic_titles.keys()):
        hcnt = hanbit_topic.get(topic, 0)
        tcnt = len(topic_titles[topic])
        share = hcnt / tcnt * 100 if tcnt else 0
        L.append(f"| {topic} | {hcnt}권 | {tcnt}권 | {share:.0f}% |")

    # 한빛 공백
    weak = [(topic, len(topic_titles[topic]), hanbit_topic.get(topic, 0))
            for topic in topic_titles if hanbit_topic.get(topic, 0) <= 1 and len(topic_titles[topic]) >= 5]
    if weak:
        L.append(f"\n### 한빛이 약한 영역 (공백)\n")
        L.append(f"| 주제 | 전체 도서 | 한빛 도서 | 주요 경쟁사 |")
        L.append(f"|------|--------:|--------:|------------|")
        for topic, total, hcnt in sorted(weak, key=lambda x: -x[1]):
            competitors = ", ".join(p for p, _ in topic_pubs.get(topic, [])[:3])
            L.append(f"| {topic} | {total}권 | {hcnt}권 | {competitors} |")

    # ── 5. 주제별 경쟁서 상세 (기타 제외, 전체 주제) ──
    L.append(f"\n## 5. 주제별 경쟁서 상세\n")
    top_topics = sorted(((t, ts) for t, ts in topic_titles.items() if t != "기타"), key=lambda x: -len(x[1]))
    for topic, titles in top_topics:
        books = []
        for t in titles:
            avg = sum(title_ranks[t]) / len(title_ranks[t]) if title_ranks[t] else 999
            nd = len(title_days[t])
            books.append((avg, nd, t, title_info[t].get("publisher", "")))
        books.sort(key=lambda x: x[0])
        L.append(f"### {topic} ({len(titles)}권)\n")
        L.append(f"| 평균순위 | 등장일 | 도서명 | 출판사 |")
        L.append(f"|-------:|------:|--------|--------|")
        for avg, nd, t, pub in books[:10]:
            L.append(f"| {avg:.0f}위 | {nd}일 | {t} | {pub} |")
        L.append("")

    return "\n".join(L)


# ══════════════════════════════════════════════════════
# 5. Claude API
# ══════════════════════════════════════════════════════

def call_claude(stats: str, dates: list[str]) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("  ⚠ ANTHROPIC_API_KEY 미설정 — 통계만 생성", file=sys.stderr)
        return ""

    date_range = f"{dates[0]} ~ {dates[-1]}"
    prompt = f"""아래는 YES24 IT 베스트셀러의 일별 스냅샷 누적 데이터(Python 자동 생성 통계)입니다.
매일 200위까지의 베스트셀러를 수집하여 {len(dates)}일간 누적한 결과입니다.

{stats}

---

위 통계 데이터 앞뒤에 붙일 해석 섹션만 작성하라.
통계 테이블은 이미 완성되어 있으므로 다시 쓰지 마라.
아래 형식대로만 작성하라:

## 핵심 인사이트

1. **[인사이트 제목]** [구체적 수치를 포함한 해석. 출판 기획 관점에서 의미를 설명.]
2. ...
3. ...
4. ...
5. ...

## 6. 출판 기획 아이템

### 기획 1: [제목]
[왜 기회인지, 어떤 도서를 만들면 좋을지, 타겟 독자, 예상 경쟁 상황. 3~5문장.]

### 기획 2: [제목]
...

(5~8개 기획 아이템)

## 7. 추천 다음 액션

- [한빛미디어 편집자가 당장 해야 할 구체적 행동 1]
- [구체적 행동 2]
- ...

(5~7개 액션)

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
    ai_sections = call_claude(stats, all_dates)

    # 리포트 조합: 헤더 + 핵심인사이트(AI) + 통계(Python) + 기획아이템(AI)
    header = f"# YES24 IT 베스트셀러 {archive['total_days']}일 분석 리포트\n\n"
    header += f"> 분석 기간: {all_dates[0]} ~ {all_dates[-1]} ({len(all_dates)}일)\n"
    header += f"> 생성일: {TODAY}\n"
    header += f"> 데이터: YES24 IT/모바일 일별 베스트셀러\n\n"

    if ai_sections:
        # AI 결과에서 "핵심 인사이트" 부분과 "6. 출판 기획" 이후 부분을 분리
        insight_part = ""
        planning_part = ""
        lines = ai_sections.split("\n")
        section = ""
        for line in lines:
            if line.startswith("## 핵심 인사이트") or line.startswith("## 핵심"):
                section = "insight"
            elif line.startswith("## 6.") or line.startswith("## 7."):
                section = "planning"
            if section == "insight":
                insight_part += line + "\n"
            elif section == "planning":
                planning_part += line + "\n"

        report = header + insight_part + "\n" + stats + "\n\n" + planning_part
    else:
        report = header + stats + "\n"

    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"✅ 리포트 생성: {REPORT_PATH}")


if __name__ == "__main__":
    main()
