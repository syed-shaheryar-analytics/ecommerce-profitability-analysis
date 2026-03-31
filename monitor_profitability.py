# -*- coding: utf-8 -*-
"""
Created on Tue Mar 31 00:08:31 2026

@author: syed shaheryar
"""

import pandas as pd
from pathlib import Path
import smtplib
from email.message import EmailMessage


INPUT_FILE = Path("orders_raw.xlsx")   
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

ALERT_MARGIN_THRESHOLD = 0.20  
SEND_EMAIL = True  # Change to True only after email details are set correctly


SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "shaheryarumts@gmail.com"
APP_PASSWORD = "your app password"
RECIPIENT_EMAIL = "shaheryarumts@gmail.com"


df = pd.read_excel(INPUT_FILE)


df.columns = df.columns.str.strip()


for col in ["Sales", "Profit", "Discount", "Quantity"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

if "Order Date" in df.columns:
    df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")

total_revenue = df["Sales"].sum()
total_profit = df["Profit"].sum()
profit_margin = total_profit / total_revenue if total_revenue != 0 else 0

alert_triggered = profit_margin < ALERT_MARGIN_THRESHOLD


category_summary = (
    df.groupby("Category", as_index=False)
      .agg(
          Revenue=("Sales", "sum"),
          Profit=("Profit", "sum"),
          Avg_Discount=("Discount", "mean")
      )
)
category_summary["Profit_Margin"] = category_summary["Profit"] / category_summary["Revenue"]

segment_summary = (
    df.groupby("Segment", as_index=False)
      .agg(
          Revenue=("Sales", "sum"),
          Profit=("Profit", "sum"),
          Avg_Discount=("Discount", "mean")
      )
)
segment_summary["Profit_Margin"] = segment_summary["Profit"] / segment_summary["Revenue"]


def discount_group(d):
    if pd.isna(d):
        return "Unknown"
    if d == 0:
        return "0%"
    elif d <= 0.10:
        return "0-10%"
    elif d <= 0.20:
        return "10-20%"
    elif d <= 0.30:
        return "20-30%"
    else:
        return "30%+"

df["Discount_Group"] = df["Discount"].apply(discount_group)

discount_summary = (
    df.groupby("Discount_Group", as_index=False)
      .agg(
          Revenue=("Sales", "sum"),
          Profit=("Profit", "sum"),
          Orders=("Sales", "count")
      )
)
discount_summary["Profit_Margin"] = discount_summary["Profit"] / discount_summary["Revenue"]

yearly_summary = None
if "Order Date" in df.columns:
    df["Year"] = df["Order Date"].dt.year
    yearly_summary = (
        df.groupby("Year", as_index=False)
          .agg(
              Revenue=("Sales", "sum"),
              Profit=("Profit", "sum")
          )
    )
    yearly_summary["Profit_Margin"] = yearly_summary["Profit"] / yearly_summary["Revenue"]

def send_alert_email(total_revenue, total_profit, profit_margin):
    msg = EmailMessage()
    msg["Subject"] = "ALERT: Profit Margin Below Threshold"
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECIPIENT_EMAIL

    body = f"""
Profitability Alert Triggered

Total Revenue: €{total_revenue:,.2f}
Total Profit: €{total_profit:,.2f}
Profit Margin: {profit_margin:.2%}

Profit margin has dropped below the defined threshold of {ALERT_MARGIN_THRESHOLD:.0%}.
Please review pricing, discounts, and cost structure.
"""
    msg.set_content(body)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)


output_file = OUTPUT_DIR / "performance_summary.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    
    pd.DataFrame([{
        "Total_Revenue": total_revenue,
        "Total_Profit": total_profit,
        "Profit_Margin": profit_margin,
        "Alert_Triggered": alert_triggered
    }]).to_excel(writer, sheet_name="KPI", index=False)

  
    category_summary.to_excel(writer, sheet_name="Category", index=False)
    segment_summary.to_excel(writer, sheet_name="Segment", index=False)
    discount_summary.to_excel(writer, sheet_name="Discount", index=False)

    if yearly_summary is not None:
        yearly_summary.to_excel(writer, sheet_name="Yearly", index=False)


print("\n=== PROFITABILITY MONITORING SUMMARY ===")
print(f"Total Revenue : €{total_revenue:,.2f}")
print(f"Total Profit  : €{total_profit:,.2f}")
print(f"Profit Margin : {profit_margin:.2%}")

if alert_triggered:
    print("\nALERT: Profit margin is below the threshold!")
    print(f"Threshold was {ALERT_MARGIN_THRESHOLD:.0%}. Review pricing, discounts, and cost structure.")

    if SEND_EMAIL:
        send_alert_email(total_revenue, total_profit, profit_margin)
        print("Alert email sent successfully.")
    else:
        print("Email sending is turned OFF. No email sent.")
else:
    print("\nStatus OK: Profit margin is above the threshold.")

print(f"\nSummary saved to: {output_file}")
