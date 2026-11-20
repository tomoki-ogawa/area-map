#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
郵便番号 → 代表座標 の対応表 (data/zip_coords.json) を作るスクリプト。

使うデータ
  1. 日本郵便「郵便番号データ（読み仮名データの促音・拗音を小書きで表記するもの）」
     KEN_ALL.CSV（全国一括, Shift_JIS）
  2. 国土交通省「位置参照情報」大字・町丁目レベル（都道府県ごとの zip, Shift_JIS）

やっていること
  - KEN_ALL の「町域名」と 位置参照情報の「大字町丁目名」を、市区町村コード単位で照合
  - 1つの郵便番号に複数の町丁目が当たる場合は、その町丁目の座標を単純平均
  - 「以下に掲載がない場合」など町域を特定しない郵便番号は、市区町村内の
    全町丁目の平均を代表点にする（精度: 市区町村）
  - 町域名で照合できなかった郵便番号も、市区町村の平均を代表点として出力し、
    tools/report/unmatched.md（と Excel 用の unmatched.csv）に一覧を書き出す

必要なもの: Python 3.8 以上（標準ライブラリのみ。pip install 不要）

使い方（リポジトリのルートで実行）
  python tools/build_zip_coords.py
  python tools/build_zip_coords.py --isj-version 20.0b   # 位置参照情報の版を指定
  python tools/build_zip_coords.py --ken-all path/to/ken_all.zip --isj-dir path/to/dir
  python tools/build_zip_coords.py --no-city-fallback      # 照合できない番号は出力しない

詳しくは README.md の「対応表の作り直し」を参照。
"""

import argparse
import csv
import io
import json
import os
import re
import sys
import unicodedata
import urllib.error
import urllib.request
import zipfile
from collections import OrderedDict, defaultdict
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

KEN_ALL_URL = ("https://www.post.japanpost.jp/service/search/zipcode/download/"
               "kogaki/zip/ken_all.zip")
ISJ_URL = "https://nlftp.mlit.go.jp/isj/dls/data/{ver}/{pref}000-{ver}.zip"

# 対象の都道府県コード: 埼玉県, 千葉県, 東京都, 神奈川県
DEFAULT_PREFS = ["11", "12", "13", "14"]

UA = "Mozilla/5.0 (zip-coords-builder)"


# --------------------------------------------------------------------------
# ダウンロード
# --------------------------------------------------------------------------
def http_get(url, dest):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=120) as res:
        data = res.read()
    if not data.startswith(b"PK"):
        raise RuntimeError("zip ではないデータが返りました: " + url)
    with open(dest, "wb") as f:
        f.write(data)
    return dest


def url_exists(url):
    req = urllib.request.Request(url, method="HEAD", headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            return res.status == 200
    except (urllib.error.HTTPError, urllib.error.URLError):
        return False


def detect_isj_version(pref, start=19, years_ahead=15):
    """公開されている大字・町丁目レベル（xx.0b）のうち最新の版を探す。"""
    latest = None
    for major in range(start, start + years_ahead):
        ver = "%d.0b" % major
        if url_exists(ISJ_URL.format(ver=ver, pref=pref)):
            latest = ver
        elif latest is not None:
            break
    return latest


# --------------------------------------------------------------------------
# 文字列の正規化
# --------------------------------------------------------------------------
KANJI_NUM = {"〇": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
             "六": 6, "七": 7, "八": 8, "九": 9}


def kanji_to_int(s):
    """一〜九十九 程度の漢数字を整数に。"""
    if not s:
        return None
    if all(c in KANJI_NUM for c in s) and "十" not in s:
        n = 0
        for c in s:
            n = n * 10 + KANJI_NUM[c]
        return n
    total, cur = 0, 0
    for c in s:
        if c == "十":
            total += (cur or 1) * 10
            cur = 0
        elif c == "百":
            total += (cur or 1) * 100
            cur = 0
        elif c in KANJI_NUM:
            cur = KANJI_NUM[c]
        else:
            return None
    return total + cur


def norm(s):
    """照合用キー。全角/半角・ヶ/ケ・先頭の「大字」「字」などの揺れを吸収する。"""
    s = unicodedata.normalize("NFKC", s or "")
    s = re.sub(r"\s+", "", s)
    s = s.replace("ヶ", "ケ").replace("ヵ", "カ").replace("ｹ", "ケ")
    s = s.replace("之", "の").replace("ノ", "の")
    s = s.replace("ッ", "ツ").replace("ヱ", "エ").replace("ヰ", "イ")
    s = re.sub(r"^(大字|字)", "", s)
    return s


CHOME_RE = re.compile(r"^(.*?)([〇一二三四五六七八九十百]+)丁目$")


def split_chome(name):
    """'内幸町一丁目' -> ('内幸町', 1) / '大手町' -> ('大手町', None)"""
    m = CHOME_RE.match(name)
    if m and m.group(1):
        n = kanji_to_int(m.group(2))
        if n is not None:
            return m.group(1), n
    return name, None


def parse_chome_list(paren):
    """括弧内が '１～３丁目' '１、２丁目' '３丁目' のような丁目指定なら番号の集合を返す。"""
    s = unicodedata.normalize("NFKC", paren)
    if "丁目" not in s:
        return None
    body = s.replace("丁目", "")
    if not re.fullmatch(r"[\d、,・~〜\-－]+", body):
        return None
    nums = set()
    for a, b in re.findall(r"(\d+)(?:[~〜\-－](\d+))?", body):
        a = int(a)
        b = int(b) if b else a
        if b < a or b - a > 50:
            return None
        nums.update(range(a, b + 1))
    return nums or None


CITY_LEVEL_RE = re.compile(r"(以下に掲載がない場合|の次に番地がくる場合|一円$)")


# --------------------------------------------------------------------------
# データ読み込み
# --------------------------------------------------------------------------
def open_csv_in_zip(path, name_filter=None):
    z = zipfile.ZipFile(path)
    names = [n for n in z.namelist() if n.lower().endswith(".csv")]
    if name_filter:
        names = [n for n in names if name_filter(n)]
    if not names:
        raise RuntimeError("CSV が見つかりません: " + path)
    raw = z.read(names[0])
    return io.StringIO(raw.decode("cp932", errors="replace"), newline="")


def load_isj(paths):
    """
    位置参照情報（大字・町丁目）を読み込み、
      city -> base(正規化) -> [(chome, lat, lng, 元の名前)]
      city -> [(lat, lng)]
    を返す。
    """
    towns = defaultdict(lambda: defaultdict(list))
    city_pts = defaultdict(list)
    city_names = {}
    for p in paths:
        f = open_csv_in_zip(p)
        reader = csv.reader(f)
        header = next(reader)
        idx = {h: i for i, h in enumerate(header)}
        for row in reader:
            if not row:
                continue
            city = row[idx["市区町村コード"]]
            name = row[idx["大字町丁目名"]]
            lat = float(row[idx["緯度"]])
            lng = float(row[idx["経度"]])
            base, chome = split_chome(norm(name))
            towns[city][base].append((chome, lat, lng, name))
            city_pts[city].append((lat, lng))
            city_names[city] = row[idx["都道府県名"]] + row[idx["市区町村名"]]
    return towns, city_pts, city_names


def load_ken_all(path, prefs):
    """
    KEN_ALL を読み、複数行にまたがる町域名を結合して
      [(zip, city_code, pref_name, city_name, town_name)] を返す。
    """
    f = open_csv_in_zip(path)
    reader = csv.reader(f)
    out = []
    pending = None
    for row in reader:
        if len(row) < 9:
            continue
        code, zipc = row[0].strip(), row[2].strip()
        if code[:2] not in prefs:
            continue
        town = row[8]
        if pending is not None:
            if pending[0] == zipc and pending[1] == code:
                pending[4] += town
                if "）" in town:
                    out.append(tuple(pending))
                    pending = None
                continue
            out.append(tuple(pending))
            pending = None
        rec = [zipc, code, row[6], row[7], town]
        if "（" in town and "）" not in town:
            pending = rec
        else:
            out.append(tuple(rec))
    if pending is not None:
        out.append(tuple(pending))
    return out


# --------------------------------------------------------------------------
# 照合
# --------------------------------------------------------------------------
def split_paren(town):
    m = re.match(r"^(.*?)（(.*)）$", town)
    if m:
        return m.group(1), m.group(2)
    m = re.match(r"^(.*?)（(.*)$", town)  # 閉じ括弧が欠けている場合
    if m:
        return m.group(1), m.group(2)
    return town, ""


def match_town(city_towns, town):
    """
    町域名 1 件を照合し、(町丁目の点のリスト, 方法) を返す。照合できなければ ([], None)。
    点は (lat, lng, 町丁目名)。
    """
    base_raw, paren = split_paren(town)
    key = norm(base_raw)
    if not key:
        return [], None

    def pts_of(base, chome_filter=None):
        items = city_towns.get(base, [])
        if chome_filter:
            sel = [it for it in items if it[0] in chome_filter]
            if sel:
                items = sel
        return [(lat, lng, nm) for (_c, lat, lng, nm) in items]

    chome_filter = parse_chome_list(paren) if paren else None

    # 1) 完全一致（丁目違いはまとめて平均）
    if key in city_towns:
        return pts_of(key, chome_filter), "exact"

    # 2) 町域名側に「〇丁目」が付いている場合
    b, c = split_chome(key)
    if c is not None and b in city_towns:
        return pts_of(b, {c}), "exact"

    # 3) 町域名が町丁目名で始まる（ビル名・地区名などが後ろに付く）→ 最長一致
    best = None
    for base in city_towns:
        if len(base) >= 2 and key.startswith(base):
            if best is None or len(base) > len(best):
                best = base
    if best:
        return pts_of(best, chome_filter), "prefix"

    # 4) 町丁目名が町域名で始まる（例: 町域「○○」⇔ 町丁目「○○町」「○○新田」）
    cands = [base for base in city_towns if len(key) >= 2 and base.startswith(key)]
    if cands:
        pts = []
        for base in cands:
            pts.extend(pts_of(base))
        return pts, "extend"

    return [], None


def mean(pts):
    n = len(pts)
    return (sum(p[0] for p in pts) / n, sum(p[1] for p in pts) / n)


def main():
    ap = argparse.ArgumentParser(description="郵便番号→代表座標の対応表を作成")
    ap.add_argument("--prefs", default=",".join(DEFAULT_PREFS),
                    help="対象の都道府県コード（カンマ区切り）既定: 11,12,13,14")
    ap.add_argument("--isj-version", default="auto",
                    help="位置参照情報（大字・町丁目）の版。例: 19.0b。既定: auto（最新を自動検出）")
    ap.add_argument("--ken-all", help="手元の ken_all.zip を使う場合のパス")
    ap.add_argument("--isj-dir", help="手元の位置参照情報 zip（{県コード}000-{版}.zip）があるフォルダ")
    ap.add_argument("--cache", default=os.path.join(ROOT, "tools", ".cache"),
                    help="ダウンロードしたファイルの保存先")
    ap.add_argument("--out", default=os.path.join(ROOT, "data", "zip_coords.json"))
    ap.add_argument("--report", default=os.path.join(ROOT, "tools", "report", "unmatched.csv"))
    ap.add_argument("--no-city-fallback", action="store_true",
                    help="町域で照合できない郵便番号を対応表に入れない（既定は市区町村の代表点で入れる）")
    args = ap.parse_args()

    prefs = [p.strip().zfill(2) for p in args.prefs.split(",") if p.strip()]
    os.makedirs(args.cache, exist_ok=True)

    # ---- KEN_ALL
    if args.ken_all:
        ken_path = args.ken_all
    else:
        ken_path = os.path.join(args.cache, "ken_all.zip")
        print("[download] KEN_ALL:", KEN_ALL_URL)
        http_get(KEN_ALL_URL, ken_path)

    # ---- 位置参照情報
    ver = args.isj_version
    isj_paths = []
    if args.isj_dir:
        for p in prefs:
            found = [f for f in os.listdir(args.isj_dir)
                     if f.startswith(p + "000-") and f.endswith(".zip")]
            if not found:
                sys.exit("位置参照情報の zip が見つかりません: 県コード " + p)
            found.sort()
            isj_paths.append(os.path.join(args.isj_dir, found[-1]))
            ver = found[-1][len(p + "000-"):-4]
    else:
        if ver == "auto":
            print("[detect] 位置参照情報の最新版を確認中...")
            ver = detect_isj_version(prefs[0])
            if not ver:
                sys.exit("位置参照情報の版を検出できませんでした。--isj-version で指定してください。")
            print("[detect] ->", ver)
        for p in prefs:
            url = ISJ_URL.format(ver=ver, pref=p)
            dest = os.path.join(args.cache, "%s000-%s.zip" % (p, ver))
            if not os.path.exists(dest):
                print("[download]", url)
                http_get(url, dest)
            isj_paths.append(dest)

    towns, city_pts, city_names = load_isj(isj_paths)
    ken_rows = load_ken_all(ken_path, set(prefs))
    print("[load] KEN_ALL 行数(結合後): %d / 町丁目数: %d"
          % (len(ken_rows), sum(len(v) for v in city_pts.values())))

    # ---- 郵便番号ごとに集約
    by_zip = OrderedDict()
    for zipc, code, pref_name, city_name, town in ken_rows:
        by_zip.setdefault(zipc, []).append((code, pref_name, city_name, town))

    result = {}
    unmatched = []
    stats = defaultdict(int)
    for zipc, rows in by_zip.items():
        pts = {}
        cities = []
        labels = []
        city_level = []
        failed = []
        for code, pref_name, city_name, town in rows:
            if code not in cities:
                cities.append(code)
            base_label = city_name + ("" if CITY_LEVEL_RE.search(town) else split_paren(town)[0])
            if base_label not in labels:
                labels.append(base_label)
            if CITY_LEVEL_RE.search(town):
                city_level.append(code)
                continue
            found, how = match_town(towns.get(code, {}), town)
            if found:
                stats[how] += 1
                for lat, lng, nm in found:
                    pts[(code, nm)] = (lat, lng)  # 同じ町丁目は1回だけ数える
            else:
                failed.append((code, pref_name, city_name, town))

        label = labels[0] + (" ほか" if len(labels) > 1 else "")
        if pts:
            lat, lng = mean(list(pts.values()))
            result[zipc] = [round(lat, 5), round(lng, 5), label, "t"]
            stats["zip_town"] += 1
            continue

        # 町丁目で決まらなかった → 市区町村の代表点
        cpts = [pt for c in cities for pt in city_pts.get(c, [])]
        if failed:
            for code, pref_name, city_name, town in failed:
                unmatched.append((zipc, code, pref_name, city_name, town))
            stats["zip_unmatched"] += 1
            if args.no_city_fallback or not cpts:
                continue
        elif not cpts:
            stats["zip_no_city"] += 1
            continue
        else:
            stats["zip_city"] += 1
        lat, lng = mean(cpts)
        result[zipc] = [round(lat, 5), round(lng, 5), label, "c"]

    # ---- 出力
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    payload = OrderedDict()
    payload["generated"] = date.today().isoformat()
    payload["prefs"] = prefs
    payload["sources"] = [
        "日本郵便 郵便番号データ (KEN_ALL)",
        "国土交通省 位置参照情報 大字・町丁目レベル %s" % ver,
    ]
    payload["format"] = "zip: [緯度, 経度, 地名, 精度(t=町丁目, c=市区町村の代表点)]"
    payload["zips"] = OrderedDict(sorted(result.items()))
    with open(args.out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        f.write("\n")

    os.makedirs(os.path.dirname(args.report), exist_ok=True)
    with open(args.report, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["郵便番号", "市区町村コード", "都道府県", "市区町村", "町域名(KEN_ALL)",
                    "対応表での扱い"])
        for zipc, code, pref_name, city_name, town in unmatched:
            how = "市区町村の代表点" if zipc in result else "対応表に入れていない"
            w.writerow([zipc, code, pref_name, city_name, town, how])

    # GitHub 上でも読めるように Markdown 版も出す（*.csv は .gitignore 対象のため）
    md_path = os.path.splitext(args.report)[0] + ".md"
    with open(md_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("# 町域名で照合できなかった郵便番号\n\n")
        f.write("- 作成日: %s\n- 位置参照情報: %s\n- 件数: %d\n\n"
                % (date.today().isoformat(), ver, len(unmatched)))
        f.write("位置参照情報（大字・町丁目）に同じ名前の町丁目が見つからなかったもの。"
                "対応表では市区町村の代表点（市区町村内の全町丁目の平均）に置いています。\n\n")
        f.write("| 郵便番号 | 都道府県 | 市区町村 | 町域名 |\n|---|---|---|---|\n")
        for zipc, code, pref_name, city_name, town in unmatched:
            f.write("| %s-%s | %s | %s | %s |\n" % (zipc[:3], zipc[3:], pref_name, city_name, town))

    print("[done] 郵便番号 %d 件を出力: %s" % (len(result), os.path.relpath(args.out, ROOT)))
    print("       町丁目で決定: %d / 市区町村単位(以下に掲載がない場合 等): %d / 照合できず: %d"
          % (stats["zip_town"], stats["zip_city"], stats["zip_unmatched"]))
    print("       照合方法(町域行): 完全一致 %d, 前方一致 %d, 拡張一致 %d"
          % (stats["exact"], stats["prefix"], stats["extend"]))
    print("[report] 照合できなかった町域: %d 行 -> %s / %s"
          % (len(unmatched), os.path.relpath(args.report, ROOT), os.path.relpath(md_path, ROOT)))


if __name__ == "__main__":
    main()
