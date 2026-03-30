import os
from utils import unzip_file, create_full_excel
from llama_agent import decide_input

DEFAULT_PATH = r"C:\Ahold\Projects\BIDM\Result"
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

def run_agent(prompt):
    # 🧠 Get input from LLaMA
    config = decide_input(prompt)

    # 🎯 Default path (used if user doesn't provide one)
    DEFAULT_PATH = r"C:\Ahold\Projects\BIDM\Result"

    # ✅ Dynamic input folder
    
    input_folder = config.get("input_path", DEFAULT_PATH)

     #Clean path (handles quotes/spaces from LLM)
    input_folder = input_folder.strip().replace('"', '')

    if not os.path.exists(input_folder):
        print(f"❌ Invalid path: {input_folder}")
        return

    print(f"📂 Using folder: {input_folder}")

    # 🔥 Fully dynamic structure
    extract_root = os.path.join(input_folder, "Extract")
    output_folder = os.path.join(input_folder, "Output")

    os.makedirs(extract_root, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    # 🎯 Get latest ZIP
    latest_zip = get_latest_zip(input_folder)

    if not latest_zip:
        print("❌ No ZIP file found")
        return

    print(f"📦 Latest ZIP: {latest_zip}")

    file_name = os.path.basename(latest_zip).replace(".zip", "")

    # 📂 Extract path per file
    extract_path = os.path.join(extract_root, file_name)
    os.makedirs(extract_path, exist_ok=True)

    # 📦 Extract only if not already done
    if not os.listdir(extract_path):
        unzip_file(latest_zip, extract_path)
    else:
        print("⏭️ Already extracted")

    # 📂 Locate report folder
    report_folder = os.path.join(extract_path, "Report")

    if not os.path.exists(report_folder):
        print("⚠️ Report folder not found")
        return

    # 📊 Output file
    output_file = os.path.join(output_folder, file_name + ".xlsx")

    # 📊 Generate Excel
    create_full_excel(report_folder, output_file)

    print(f"🎯 Done! File saved at:\n{output_file}")

if __name__ == "__main__":
    prompt = input("💬 Enter your requirements: ")
    run_agent(prompt)


