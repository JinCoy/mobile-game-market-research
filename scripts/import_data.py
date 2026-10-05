"""Read original workbooks without changing them; rebuild the dashboard dataset."""
import argparse
import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parents[1]

def clean(value):
    if value is None:
        return None
    if isinstance(value, (int, float, bool)):
        return value
    if isinstance(value, datetime):
        return value.isoformat()
    value = unicodedata.normalize("NFC", str(value))
    return re.sub(r"[\U0001F000-\U0001FFFF\u2600-\u27BF\uFE0F]", "", value).strip()

def key(name):
    name = unicodedata.normalize("NFKD", name).casefold()
    return re.sub(r"[^a-z0-9가-힣]", "", name)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=ROOT)
    args = parser.parse_args()
    aliases = json.loads((ROOT / "data/aliases.json").read_text())
    config = json.loads((ROOT / "data/import-config.json").read_text())
    hooks = json.loads((ROOT / "data/marketing-hypotheses.json").read_text())
    alias_lookup = {key(name): key(group[0]) for group in aliases for name in group}
    def identity(name):
        return alias_lookup.get(key(name), key(name))

    workbooks = {}
    documents = []
    for path in sorted(args.input_dir.glob("*.xlsx")):
        filename = unicodedata.normalize("NFC", path.name)
        workbook = openpyxl.load_workbook(path, data_only=True)
        sheets = {}
        for sheet in workbook.worksheets:
            rows = [[clean(v) for v in row] for row in sheet.values]
            sheets[sheet.title] = rows
            documents.append({"id": key(filename + sheet.title), "file": filename,
                              "sheet": sheet.title, "rows": rows})
        workbooks[filename] = sheets
    if not workbooks:
        raise ValueError("No XLSX workbooks found")
    market = next(v for k, v in workbooks.items() if "시장조사" in k)
    strategy = next(v for k, v in workbooks.items() if "게임전략" in k)
    overview = {r[0]: r[1] for r in market["0-1 개요"] if r[0] and r[1]}
    observations = []
    games = {}
    def add(name, publisher, genre, subgenre, country, store, date, rank, chart, sheet, row, notes=None, url=None):
        if not isinstance(rank, int) or not 1 <= rank <= 100:
            return
        missing = name == "확인 불가"
        game_id = identity(name) if not missing else f"missing-{country}-{store}-{chart}-{rank}"
        if game_id not in games:
            games[game_id] = {"id": game_id, "name": name, "publisher": publisher,
                              "developer": None, "genre": genre, "subgenre": subgenre,
                              "gameplay": None, "monetization": None, "marketing": None,
                              "artStyle": None, "audience": None, "playerAge": None,
                              "downloads": None, "revenue": None, "analysis": None,
                              "evidence": None, "missing": missing, "marketingHook": None}
        observations.append({"id": f"{country}-{store}-{chart}-{date}-{rank}", "gameId": game_id,
                             "name": name, "publisher": publisher, "genre": genre, "subgenre": subgenre,
                             "country": country, "store": store, "date": date, "year": int(date[:4]),
                             "rank": rank, "chart": chart, "notes": notes,
                             "source": {"file": "시장조사.xlsx", "sheet": sheet, "row": row, "url": url}})
    for store, sheet, url in [("ios", "1-1 App Store Top100", "https://gamedropdaily.com/mobile/"),
                               ("android", "1-2 Google Play Top100", "https://www.appbrain.com/stats/google-play-rankings/top_free/game/us")]:
        date = overview[("App Store" if store == "ios" else "Google Play") + " 기준일"]
        for row, values in enumerate(market[sheet][1:], 2):
            rank, name, publisher, genre, subgenre, notes = values[:6]
            add(name, publisher, genre, subgenre, "미국", store, date, rank, "free", sheet, row, notes, url)
    for sheet in ["2-2 국가별 Top20", "2-7 동남아 분석"]:
        for row, values in enumerate(market[sheet], 1):
            if len(values) >= 8 and isinstance(values[3], int) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(values[2])):
                _, country, date, rank, name, publisher, genre, source = values[:8]
                urls = re.findall(r"https?://\S+", str(source))
                add(name, publisher, genre, None, country, "android", date, rank, "free", sheet, row,
                    values[8] if len(values) > 8 else None, urls[0] if urls else None)
    for row, values in enumerate(market["3-2 매출 Top100"][1:], 2):
        rank, name, publisher, genre, subgenre, notes = values[:6]
        add(name, publisher, genre, subgenre, "미국", "android", config["grossingDate"], rank, "grossing",
            "3-2 매출 Top100", row, notes, "https://www.appbrain.com/stats/google-play-rankings/top_grossing/game/us")
    for sheet in ["1-4 성공 비결 해부", "1-5 성공 비결 확장"]:
        for row, values in enumerate(strategy[sheet][1:], 2):
            if not values[0]:
                continue
            expanded = "확장" in sheet
            name = values[1] if expanded else values[0]
            game = games.get(identity(name)) if isinstance(name, str) else None
            if game:
                offset = 4 if expanded else 1
                game.update({"gameplay": values[offset], "monetization": values[6] if expanded else values[4],
                             "marketing": values[7] if expanded else values[5], "artStyle": values[8] if expanded else values[6],
                             "evidence": values[9] if expanded else values[7],
                             "analysis": {"file": "게임전략_MVP_마케팅.xlsx", "sheet": sheet, "row": row}})
    for values in market["3-3 매출-다운로드 비교"][5:]:
        if values[0] and isinstance(values[3], (int, float)):
            game = games.get(identity(values[0]))
            if game:
                game["downloads"] = {"value": values[3], "kind": "AppBrain 추정 누적 설치", "date": config["installEstimateDate"],
                                     "store": "android", "country": None}
    for name, hook in hooks.items():
        game = games.get(identity(name))
        if game:
            game["marketingHook"] = hook
    dataset = {"schemaVersion": 1, "importedAt": datetime.now(timezone.utc).isoformat(),
               "games": list(games.values()), "observations": observations, "documents": documents,
               "files": [{"name": name, "sheets": len(sheets)} for name, sheets in workbooks.items()]}
    (ROOT / "data/market.json").write_text(json.dumps(dataset, ensure_ascii=False, indent=2))
    print(f"Imported {len(games)} games, {len(observations)} observations, {len(documents)} sheets")

if __name__ == "__main__":
    main()
