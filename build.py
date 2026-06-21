"""
アイデアカード 更新スクリプト
-------------------------------
使い方:
  1. このフォルダに Excel ファイルを置く（ファイル名は何でもOK、.xlsx）
  2. ターミナルで: python build.py
  または update.bat をダブルクリック
"""
import json, sys, subprocess
from pathlib import Path

HERE = Path(__file__).parent

# ── Excel を探す ──────────────────────────────────────────
xlsx_files = list(HERE.glob("*.xlsx"))
if not xlsx_files:
    sys.exit("[ERROR] .xlsx ファイルが見つかりません。このフォルダに Excel を置いてください。")
if len(xlsx_files) > 1:
    print("複数の .xlsx が見つかりました。最初のものを使います:", xlsx_files[0].name)
xlsx_path = xlsx_files[0]
print(f"[読み込み] {xlsx_path.name}")

# ── Excel → JSON ──────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl

wb = openpyxl.load_workbook(xlsx_path)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))

# ヘッダー行をスキップ（1行目）
cards = []
for row in rows[1:]:
    if row[0] is None:
        continue
    cards.append({
        "no":      row[0],
        "name":    str(row[1]) if row[1] else "",
        "summary": str(row[2]) if row[2] else "",
        "parent":  str(row[3]) if row[3] else "",
        "detail":  str(row[4]) if row[4] else "",
        "other":   str(row[5]) if row[5] else "",
    })

print(f"[OK] {len(cards)} 件のアイデアを読み込みました")

# ── template.html に JSON を埋め込む ──────────────────────
template_path = HERE / "template.html"
if not template_path.exists():
    sys.exit("[ERROR] template.html が見つかりません")

template = template_path.read_text(encoding="utf-8")
cards_json = json.dumps(cards, ensure_ascii=False)
html = template.replace("__CARDS_JSON__", cards_json)

out_path = HERE / "index.html"
out_path.write_text(html, encoding="utf-8")
print(f"[OK] index.html を生成しました ({len(html):,} 文字)")

# ── Git push ──────────────────────────────────────────────
import shutil
if not shutil.which("git"):
    print("[!] git が見つかりません。index.html を手動でコミットしてください。")
    sys.exit(0)

try:
    subprocess.run(["git", "add", "index.html"], cwd=HERE, check=True)
    result = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        cwd=HERE
    )
    if result.returncode == 0:
        print("[info] 変更なし。プッシュをスキップします。")
        sys.exit(0)

    subprocess.run(
        ["git", "commit", "-m", f"Update cards: {len(cards)} ideas from {xlsx_path.name}"],
        cwd=HERE, check=True
    )
    subprocess.run(["git", "push"], cwd=HERE, check=True)
    print("[OK] GitHub Pages にデプロイしました！")
    print("   URL: https://yuto-hub-627.github.io/idea-cards/")
    print("   ※反映まで1〜2分かかります")
except subprocess.CalledProcessError as e:
    print(f"[!] Git エラー: {e}")
    print("   index.html は生成済みです。手動でコミット・プッシュしてください。")
