#!/usr/bin/env python3
import sys
if sys.stdout.encoding and sys.stdout.encoding.lower().startswith("cp"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
"""
리포트 빌더
- data/reports/*.md 파일을 스캔
- reports.js 인덱스 생성 (window._REPORTS)
- 각 .md 내용을 JS에 임베드 → 브라우저에서 바로 렌더링

사용법:
  python scripts/build_reports.py                    # 기존 .md → reports.js 빌드
  python scripts/build_reports.py --analyze "경로"   # yes24 분석 실행 후 빌드
"""
import json
import os
import re
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(SCRIPT_DIR, "..", "data", "reports")


def extract_meta(md_text: str, filename: str) -> dict:
    """마크다운에서 제목, 날짜, 요약을 추출한다."""
    lines = md_text.strip().split("\n")

    # 제목: 첫 번째 # 헤딩
    title = filename
    for line in lines[:5]:
        m = re.match(r"^#\s+(.+)", line)
        if m:
            title = m.group(1).strip()
            break

    # 날짜: "생성일:" 우선 → 파일명 YYYYMMDD → 본문 첫 날짜
    date = ""
    for line in lines[:10]:
        gm = re.search(r"생성일[:\s]+(\d{4})-(\d{2})-(\d{2})", line)
        if gm:
            date = gm.group(0).split(":")[-1].strip()[:10]
            break
    if not date:
        dm = re.search(r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})", filename)
        if dm:
            date = f"{dm.group(1)}-{dm.group(2)}-{dm.group(3)}"
        else:
            for line in lines[:10]:
                dm2 = re.search(r"(\d{4})-(\d{2})-(\d{2})", line)
                if dm2:
                    date = dm2.group(0)
                    break

    # 요약: "핵심 인사이트" 또는 첫 문단
    summary = ""
    for i, line in enumerate(lines):
        if "핵심" in line or "요약" in line or "인사이트" in line:
            # 다음 줄들에서 첫 내용 가져오기
            for j in range(i + 1, min(i + 5, len(lines))):
                stripped = lines[j].strip().lstrip("0123456789.-) ")
                if stripped and not stripped.startswith("#"):
                    summary = re.sub(r"\*\*(.+?)\*\*", r"\1", stripped)[:150]
                    break
            break
    if not summary:
        for line in lines[2:10]:
            stripped = line.strip()
            if stripped and not stripped.startswith("#") and not stripped.startswith(">") and len(stripped) > 20:
                summary = stripped[:150]
                break

    return {"title": title, "date": date, "summary": summary}


def build_reports_js():
    """data/reports/*.md → reports.js 빌드."""
    os.makedirs(REPORTS_DIR, exist_ok=True)

    reports = []
    md_files = sorted(
        [f for f in os.listdir(REPORTS_DIR) if f.endswith(".md")],
        reverse=True,  # 최신 먼저
    )

    for fname in md_files:
        fpath = os.path.join(REPORTS_DIR, fname)
        with open(fpath, "r", encoding="utf-8") as f:
            content = f.read()

        meta = extract_meta(content, fname)
        rid = fname.replace(".md", "")

        reports.append({
            "id": rid,
            "filename": fname,
            "title": meta["title"],
            "date": meta["date"],
            "summary": meta["summary"],
            "content": content,
        })

    out = {"built_at": datetime.now().strftime("%Y-%m-%d %H:%M"), "reports": reports}

    js_path = os.path.join(REPORTS_DIR, "reports.js")
    with open(js_path, "w", encoding="utf-8") as f:
        f.write("window._REPORTS = ")
        json.dump(out, f, ensure_ascii=False)
        f.write(";")

    json_path = os.path.join(REPORTS_DIR, "reports.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)

    print(f"✅ reports.js/json 빌드 완료: {len(reports)}개 리포트")
    for r in reports:
        print(f"   [{r['date']}] {r['title']}")


def run_analysis(data_dir: str):
    """yes24 분석 실행 → .md 리포트 생성."""
    # 기존 분석 스크립트 재활용
    analysis_script = os.path.join(SCRIPT_DIR, "..", "_workspace", "yes24_analysis.py")
    if not os.path.exists(analysis_script):
        print("⚠ _workspace/yes24_analysis.py가 없습니다. 먼저 분석을 실행하세요.")
        return

    # 분석 실행
    today = datetime.now().strftime("%Y%m%d")
    out_path = os.path.join(REPORTS_DIR, f"{today}_yes24_analysis.md")

    import subprocess
    result = subprocess.run(
        [sys.executable, analysis_script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
    )

    if result.returncode != 0:
        print(f"⚠ 분석 실패: {result.stderr[:300]}")
        return

    # _workspace의 결과를 reports로 복사
    ws_report = os.path.join(SCRIPT_DIR, "..", "_workspace", f"yes24_analysis_{today}.md")
    if os.path.exists(ws_report):
        import shutil
        shutil.copy2(ws_report, out_path)
        print(f"✅ 리포트 생성: {out_path}")
    else:
        print(f"⚠ 리포트 파일을 찾을 수 없습니다: {ws_report}")


def main():
    args = sys.argv[1:]

    if "--analyze" in args:
        idx = args.index("--analyze")
        data_dir = args[idx + 1] if idx + 1 < len(args) else ""
        if data_dir:
            run_analysis(data_dir)

    build_reports_js()


if __name__ == "__main__":
    main()
