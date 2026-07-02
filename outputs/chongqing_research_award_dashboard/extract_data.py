import json
from pathlib import Path

import pandas as pd

FILES = [
    (
        "/Users/luosiyao/Library/Containers/com.tencent.xinWeChat/Data/Library/Caches/com.tencent.xinWeChat/2.0b4.0.9/3a082e0dbd57afacfcda605ed43f569f/SaveTemp/386ad4fcd9478e78e64dc1a37d7499d5/研究成果_评审结果_一等奖.xlsx",
        "一等奖",
    ),
    (
        "/Users/luosiyao/Library/Containers/com.tencent.xinWeChat/Data/Library/Caches/com.tencent.xinWeChat/2.0b4.0.9/3a082e0dbd57afacfcda605ed43f569f/SaveTemp/6314ac9722cf640e1c5f87ab85a29f9b/研究成果_评审结果_二等奖.xlsx",
        "二等奖",
    ),
    (
        "/Users/luosiyao/Library/Containers/com.tencent.xinWeChat/Data/Library/Caches/com.tencent.xinWeChat/2.0b4.0.9/3a082e0dbd57afacfcda605ed43f569f/SaveTemp/15bdaa899b317ee3c511170b6da6341c/研究成果_评审结果_三等奖.xlsx",
        "三等奖",
    ),
]


def score_band(score: float) -> str:
    if score >= 94:
        return "94分及以上"
    if score >= 92:
        return "92-94分"
    if score >= 90:
        return "90-92分"
    if score >= 88:
        return "88-90分"
    return "88分以下"


rows = []
for path, award in FILES:
    df = pd.read_excel(path)
    df["奖项"] = award
    rows.append(df)

all_df = pd.concat(rows, ignore_index=True)
award_order = {"一等奖": 1, "二等奖": 2, "三等奖": 3}
all_df["奖项序"] = all_df["奖项"].map(award_order)
all_df["奖项内排名"] = all_df.groupby("奖项")["综合最终得分"].rank(
    method="first", ascending=False
).astype(int)
all_df["大会-小组差"] = (all_df["大会评分"] - all_df["小组评分"]).round(2)
all_df["分数区间"] = all_df["综合最终得分"].map(score_band)
all_df["是否临界"] = all_df.groupby("奖项")["综合最终得分"].transform(
    lambda s: s <= s.quantile(0.2)
)
all_df = all_df.sort_values(
    ["奖项序", "综合最终得分", "大会评分", "小组评分"], ascending=[True, False, False, False]
)

out_rows = all_df[
    [
        "编号",
        "成果名称",
        "奖项",
        "奖项内排名",
        "小组评分",
        "大会评分",
        "大会-小组差",
        "综合最终得分",
        "分数区间",
        "是否临界",
    ]
].to_dict(orient="records")

summary = {
    "total": int(len(all_df)),
    "award_counts": all_df["奖项"].value_counts().reindex(["一等奖", "二等奖", "三等奖"]).fillna(0).astype(int).to_dict(),
    "min_score": float(all_df["综合最终得分"].min()),
    "max_score": float(all_df["综合最终得分"].max()),
    "avg_score": float(all_df["综合最终得分"].mean()),
}

Path("outputs/chongqing_research_award_dashboard/data.json").write_text(
    json.dumps({"rows": out_rows, "summary": summary}, ensure_ascii=False, indent=2),
    encoding="utf-8",
)
