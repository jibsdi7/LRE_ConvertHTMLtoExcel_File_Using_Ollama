import os
from utils import unzip_file, create_full_excel

# ✅ Correct columns (based on your data)



def get_latest_zip(folder):
    zip_files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith(".zip")
    ]

    if not zip_files:
        return None

    return max(zip_files, key=os.path.getmtime)


def run():
    input_folder = r"C:\Ahold\Projects\BIDM\Result"
    extract_root = r"C:\Ahold\Projects\BIDM\Result\Extract"
    output_folder = "output"

    os.makedirs(extract_root, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    # 🎯 Get latest ZIP
    latest_zip = get_latest_zip(input_folder)

    if not latest_zip:
        print("❌ No ZIP file found")
        return

    print(f"📦 Latest ZIP: {latest_zip}")

    file_name = os.path.basename(latest_zip).replace(".zip", "")

    extract_path = os.path.join(extract_root, file_name)
    os.makedirs(extract_path, exist_ok=True)

    # ✅ Avoid re-extracting
    if not os.listdir(extract_path):
        unzip_file(latest_zip, extract_path)
    else:
        print("⏭️ Already extracted")

    # 📂 Go to Report folder
    report_folder = os.path.join(extract_path, "Report")

    if not os.path.exists(report_folder):
        print("⚠️ Report folder not found")
        return

    # 📊 Output file
    output_file = os.path.join(output_folder, file_name + ".xlsx")

    # 📊 Convert → single sheet
    create_full_excel(report_folder, output_file)


if __name__ == "__main__":
    run()