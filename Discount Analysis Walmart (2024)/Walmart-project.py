# ---------------------------------------------
# 📦 Walmart Supplement Sales Analysis (Canada - 2024)
# 🧑‍💻 Author: Parham Parvizi
# 📅 Date: 2025-04
# 📝 Description: Full data analysis pipeline for supplement sales performance,
#     discount volatility, return impact, and profitability for Walmart Canada
# ---------------------------------------------

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# ---------------------------------------------
# 🧹 Step 1: Load & Prepare the Dataset
# ---------------------------------------------
df = pd.read_csv("Supplement_Sales_Weekly.csv")
df["Date"] = pd.to_datetime(df["Date"])
df = df[(df["Platform"] == "Walmart") & (df["Date"].dt.year == 2024)]
df = df[df["Location"] == "Canada"].reset_index(drop=True)

# ---------------------------------------------
# 📊 Step 2: Monthly & Quarterly Aggregations
# ---------------------------------------------
df["Month"] = df["Date"].dt.strftime("%Y-%m")
df["Quarter"] = df["Date"].dt.to_period("Q")

agg_funcs = {
    "Units Sold": "sum",
    "Revenue": "sum",
    "Units Returned": "sum",
    "Price": "mean",
    "Discount": "mean"
}

monthly_df = df.groupby("Month").agg(agg_funcs).round(2)
quarterly_df = df.groupby("Quarter").agg(agg_funcs).round(2)

# ---------------------------------------------
# 🧪 Step 3: Category-level Aggregation
# ---------------------------------------------
def category_monthly(cat):
    temp = df[df["Category"] == cat].copy()
    return temp.groupby("Month").agg(agg_funcs).round(2)

vitamin = category_monthly("Vitamin")
mineral = category_monthly("Mineral")
protein = category_monthly("Protein")

# ---------------------------------------------
# 📉 Step 4: Discount Volatility Chart (Quarterly)
# ---------------------------------------------
discount_q = df.groupby("Quarter")["Discount"].std().fillna(0) * 100
x_axis = pd.period_range("2024Q1", periods=4, freq="Q").astype(str)

plt.figure(figsize=(8, 5))
plt.plot(x_axis, discount_q, label="Discount Volatility", color="purple", linestyle="--", marker="o")
plt.title("Quarterly Discount Volatility Trends (2024)", fontsize=14)
plt.xlabel("Quarter")
plt.ylabel("Std Dev of Discount (%)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# ---------------------------------------------
# 🔁 Step 5: Discount vs Return Impact
# ---------------------------------------------
df_monthly = df.groupby("Month").agg(agg_funcs).round(2)
df_monthly["Return Impact"] = (df_monthly["Units Returned"] / df_monthly["Units Sold"]) * 100

discount_line = df_monthly["Discount"]
return_line = df_monthly["Return Impact"]

plt.figure(figsize=(8, 5))
plt.plot(df_monthly.index, discount_line * 100, label="Discount (%)", color="blue", linewidth=2)
plt.plot(df_monthly.index, return_line * 5, label="Return Impact (×5)", color="orange", linewidth=2)
plt.title("Movement of Discount and Return Impact", fontsize=14)
plt.xlabel("Month")
plt.ylabel("Percentage %")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# ---------------------------------------------
# 🔍 Step 6: Correlation Heatmap
# ---------------------------------------------
corr_matrix = df_monthly[["Discount", "Return Impact", "Revenue", "Price"]].corr()
plt.figure(figsize=(8, 5))
sns.heatmap(corr_matrix.round(2), annot=True, cmap="mako", vmin=-1, vmax=1)
plt.title("Correlation Between Metrics", fontsize=13)
plt.tight_layout()
plt.show()

# ---------------------------------------------
# 💵 Step 7: Profit Per Unit Calculation
# ---------------------------------------------
profit_per_unit = (df_monthly["Units Sold"] - df_monthly["Units Returned"]) * \
                  (df_monthly["Price"] - df_monthly["Discount"])
profit_df = profit_per_unit.to_frame("Profit per Unit")

# ---------------------------------------------
# 📤 Step 8: Export to Excel
# ---------------------------------------------
with pd.ExcelWriter("supplement_analysis_final.xlsx") as writer:
    df.to_excel(writer, sheet_name="DataFrame")
    monthly_df.to_excel(writer, sheet_name="Monthly Summary")
    quarterly_df.to_excel(writer, sheet_name="Quarterly Summary")
    vitamin.to_excel(writer, sheet_name="Vitamin")
    mineral.to_excel(writer, sheet_name="Mineral")
    protein.to_excel(writer, sheet_name="Protein")
    profit_df.to_excel(writer, sheet_name="Profit per Unit")
    df_monthly.to_excel(writer, sheet_name="Discount & Return")

    pd.DataFrame({
        "Source": ["Walmart Supplement Sales"],
        "Region": ["Canada"],
        "Year": [2024],
        "Prepared by": ["Parham Parvizi"]
    }).to_excel(writer, sheet_name="Info", index=False)
