import pickle
import pandas as pd

with open("scratch/multi_year_backtest_results.pkl", "rb") as f:
    res = pickle.load(f)

df = pd.DataFrame(res["raw_opt"]["trades"])

def adj_net(row):
    cd = row["cost_detail"]
    slip = cd.get("slippage", 0.0) if isinstance(cd, dict) else 0.0
    r = str(row["reason"])
    if "STOP_LOSS" in r or "PROFIT_LOCK" in r:
        return row["net"] + slip
    return row["net"]

df["net_sll"] = df.apply(adj_net, axis=1)
df["gross_pct"] = (df["gross"] / (df["entry"] * df["qty"])) * 100.0
df["hour"] = pd.to_datetime(df["entry_time"]).dt.hour

print("--- BY FULL EXIT REASON (TOP 10) ---")
for r, g in df.groupby("reason"):
    if len(g) < 10:
        continue
    wr = (g["net_sll"] > 0).mean() * 100
    gp = g["gross_pct"].mean()
    mfe = g["mfe"].mean()
    mae = g["mae"].mean()
    sg = g["gross"].sum()
    sn = g["net_sll"].sum()
    print(f"{r[:38]:38s} | n={len(g):4d} | WR={wr:5.1f}% | GrossPct={gp:+.3f}% | MFE={mfe:.2f}% | MAE={mae:.2f}% | Gross=Rs.{sg:+9,.0f} | Net=Rs.{sn:+9,.0f}")

print("\n--- BY ENTRY HOUR ---")
for h, g in df.groupby("hour"):
    wr = (g["net_sll"] > 0).mean() * 100
    gp = g["gross_pct"].mean()
    sg = g["gross"].sum()
    sn = g["net_sll"].sum()
    print(f"Hour {h:02d}:00 | n={len(g):4d} | WR={wr:5.1f}% | GrossPct={gp:+.3f}% | Gross=Rs.{sg:+9,.0f} | Net=Rs.{sn:+9,.0f}")

print("\n--- BY STOCK VOLATILITY / SYMBOL (TOP & BOTTOM 5) ---")
sym_grp = df.groupby("symbol").agg(
    n=("net_sll", "count"),
    wr=("net_sll", lambda s: (s > 0).mean() * 100),
    gross_pct=("gross_pct", "mean"),
    gross=("gross", "sum"),
    net=("net_sll", "sum"),
    mfe=("mfe", "mean"),
    mae=("mae", "mean")
).sort_values("net", ascending=False)
print("Top 8 symbols:")
print(sym_grp.head(8).to_string())
print("\nBottom 8 symbols:")
print(sym_grp.tail(8).to_string())
