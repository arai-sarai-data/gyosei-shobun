"""行政処分の横断検索：金融庁「行政処分事例集」（Excel）から、業務改善命令・業務停止命令などの
行政処分を会社名・業態・年で検索でき、処分が多い会社・業態をランキングにする。"""
import io
import re
from collections import Counter, defaultdict

from lib import oc

FSA_PAGE = "https://www.fsa.go.jp/status/s_jirei/kouhyou.html"


def fsa():
    import pandas as pd
    urls = oc.links(FSA_PAGE, r"\.xlsx?$")
    oc.check(urls, "行政処分事例集のExcelが見つかりません")
    rows = []
    for u in urls:
        b = oc.get(u, cache="fsa_" + u.rsplit("/", 1)[-1])
        for name, df in pd.read_excel(io.BytesIO(b), sheet_name=None, header=None).items():
            hdr = None
            for i in range(min(15, len(df))):
                cells = [str(x) for x in df.iloc[i].tolist()]
                if any("処分" in c for c in cells) and any(("名" in c and ("金融機関" in c or "業者" in c or "機関" in c)) for c in cells):
                    hdr = i; break
            if hdr is None:
                continue
            df.columns = [str(c).strip() for c in df.iloc[hdr].tolist()]
            df = df.iloc[hdr + 1:]
            c_name = oc.col(df, "名")
            c_date = oc.col(df, "日", required=False)
            c_kind = oc.col(df, "処分", required=False)
            c_type = oc.col(df, "業態", required=False)
            c_law = oc.col(df, "根拠", required=False) or oc.col(df, "法令", required=False)
            for _, r in df.iterrows():
                nm = str(r[c_name]).strip()
                if not nm or nm == "nan":
                    continue
                raw = r[c_date] if c_date else ""
                d = str(raw)
                if re.fullmatch(r"\d{5}(\.0)?", d):  # Excelの日付シリアル値
                    import datetime as _d
                    d = (_d.date(1899, 12, 30) + _d.timedelta(days=int(float(d)))).isoformat()
                elif hasattr(raw, "isoformat"):
                    d = raw.isoformat()
                y = re.search(r"(19|20)\d\d", d)
                rows.append({"co": nm, "date": d[:10], "year": int(y.group(0)) if y else None,
                             "kind": str(r[c_kind]).strip() if c_kind else "", "type": str(r[c_type]).strip() if c_type else name,
                             "law": str(r[c_law]).strip() if c_law else "", "src": "金融庁"})
        oc.log(u, len(rows))
    return rows


def run():
    rows = fsa()
    oc.check(len(rows) > 1000, "行政処分の件数が少なすぎます(%d)" % len(rows))
    cnt = Counter(r["co"] for r in rows)
    last = defaultdict(str); typ = {}
    for r in rows:
        last[r["co"]] = max(last[r["co"]], r["date"]); typ[r["co"]] = r["type"]
    cos = sorted([{"co": k, "n": v, "type": typ[k], "last": last[k]} for k, v in cnt.items()], key=lambda r: -r["n"])
    types = sorted([{"type": k, "n": v} for k, v in Counter(r["type"] for r in rows).items()], key=lambda r: -r["n"])
    years = sorted([{"year": str(k), "n": v} for k, v in Counter(r["year"] for r in rows if r["year"]).items()], key=lambda r: r["year"])
    kinds = sorted([{"kind": k[:40], "n": v} for k, v in Counter(r["kind"] for r in rows if r["kind"]).items()], key=lambda r: -r["n"])
    meth = ("金融庁が公表している「行政処分事例集」（金融庁・財務局等が出した行政処分の一覧）を集計しました。"
            "事例集の1行（1つの命令）を1件として数えるため、同じ日に業務改善命令と業務停止命令が出ると2件になります。同じ会社でも、合併や商号変更で名前が変わると別々に数えています。処分の重さは種類によって大きく違います。")
    R = [{"id": "co", "tab": "会社", "title": "行政処分を受けた回数が多い金融機関・業者", "rows": cos, "method": meth,
          "columns": [{"key": "co", "label": "会社名"}, {"key": "n", "label": "処分の件数", "type": "num", "bar": True},
                      {"key": "type", "label": "業態"}, {"key": "last", "label": "最新の処分日"}],
          "trend": oc.trend(cos, "type", 50, "社", "社"), "share_label": "co", "share_value": "n"},
         {"id": "type", "tab": "業態", "title": "行政処分が多い業態", "rows": types, "method": meth,
          "columns": [{"key": "type", "label": "業態"}, {"key": "n", "label": "処分の件数", "type": "num", "bar": True}],
          "share_label": "type", "share_value": "n"},
         {"id": "kind", "tab": "処分の種類", "title": "処分の種類（多い順）", "rows": kinds, "method": meth,
          "columns": [{"key": "kind", "label": "処分の種類"}, {"key": "n", "label": "件数", "type": "num", "bar": True}],
          "share_label": "kind", "share_value": "n"},
         {"id": "year", "tab": "年別", "title": "年別の行政処分の件数", "rows": years, "method": meth,
          "columns": [{"key": "year", "label": "年"}, {"key": "n", "label": "件数", "type": "num", "bar": True}],
          "share_label": "year", "share_value": "n"}]
    meta = {"data_period": "金融庁 行政処分事例集（最新版）",
            "stats": [{"label": "収録した行政処分", "value": len(rows), "unit": "件"}, {"label": "処分を受けた会社", "value": len(cnt), "unit": "社"},
                      {"label": "最も多い会社の件数", "value": cos[0]["n"], "unit": "件"}],
            "sources": [{"name": "行政処分事例集", "org": "金融庁", "url": FSA_PAGE, "license": "政府標準利用規約（第2.0版）"}]}
    oc.write(meta, R, search=[{"n": r["co"], "s": r["type"], "a": r["date"], "k": [["処分", r["kind"][:60]], ["根拠", r["law"][:60]]]} for r in rows])


if __name__ == "__main__":
    oc.main(run)
