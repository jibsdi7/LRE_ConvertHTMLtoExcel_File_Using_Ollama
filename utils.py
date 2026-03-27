import zipfile
import os
import pandas as pd
from bs4 import BeautifulSoup


# 📦 Unzip file
def unzip_file(zip_path, extract_to):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print(f"✅ Extracted: {zip_path}")


def extract_full_summary(html_file):
    import pandas as pd
    from bs4 import BeautifulSoup

    with open(html_file, "r", encoding="utf-8", errors="ignore") as f:
        soup = BeautifulSoup(f, "lxml")

    # =============================
    # 1️⃣ TRANSACTION TABLE
    # =============================
    table = soup.find("table", {"id": "TransactionsTable"})

    transactions_df = None

    if table:
        rows = []
        for tr in table.find_all("tr"):
            cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cols:
                rows.append(cols)

        if len(rows) > 1:
            transactions_df = pd.DataFrame(rows[1:], columns=rows[0])
            transactions_df.columns = transactions_df.columns.str.strip()

    # =============================
    # 2️⃣ SUMMARY (FIXED LOGIC)
    # =============================
    summary_data = {}

    # 🔥 Extract PERIOD
    full_text = soup.get_text("\n")
    for line in full_text.split("\n"):
        if line.strip().startswith("Period:"):
            summary_data["Period"] = line.replace("Period:", "").strip()

    # 🔥 Extract ALL <td> values
    tds = [td.get_text(strip=True) for td in soup.find_all("td")]

    # 🎯 Keys we want
    stats_keys = [
        "Maximum Running Vusers",
        "Total Throughput (bytes)",
        "Average Throughput (B/s)",
        "Total Hits:",
        "Average Hits per Second",
        "Passed Transactions Ratio"
    ]

    # 🔥 Pair KEY → VALUE (td[i] → td[i+1])
    for i in range(len(tds) - 1):
        key = tds[i]
        value = tds[i + 1]

        if key in stats_keys:
            summary_data[key.replace(":", "")] = value

    # 📊 Convert to DataFrame
    summary_df = pd.DataFrame(
        list(summary_data.items()),
        columns=["Metric", "Value"]
    )

    return transactions_df, summary_df


# 📊 Create final Excel with multiple sheets
def create_full_excel(report_folder, output_file):
    html_file = os.path.join(report_folder, "summary.html")

    if not os.path.exists(html_file):
        print("❌ summary.html not found")
        return

    transactions_df, summary_df = extract_full_summary(html_file)

    if transactions_df is None:
        print("❌ No transaction table found")
        return

    # 📂 Ensure output folder exists
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    # 📊 Write to Excel
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        transactions_df.to_excel(writer, sheet_name="Transactions", index=False)

        if not summary_df.empty:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

    print(f"🎯 Full Excel created: {output_file}")