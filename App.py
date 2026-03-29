from fastapi import FastAPI, Body
import os

from utils import unzip_file, create_full_excel
from llama_agent import decide_input

app = FastAPI()

DEFAULT_PATH = r"C:\Ahold\Projects\BIDM\Result"


# 🎯 Get latest ZIP
def get_latest_zip(folder):
    zip_files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith(".zip")
    ]

    if not zip_files:
        return None

    return max(zip_files, key=os.path.getmtime)


# 🚀 AGENT ENDPOINT (SAME AS main.py)
@app.post("/agent")
def run_agent_api(data: dict = Body(...)):
    try:
        prompt = data.get("prompt", "")

        if not prompt:
            return {"error": "Prompt is required"}

        print(f"💬 Prompt: {prompt}")

        # 🧠 Get input from LLaMA
        config = decide_input(prompt)

        # ✅ Dynamic input folder
        #input_folder = config.get("input_path", DEFAULT_PATH)

        input_folder = config.get("input_path")

# 🔥 Validate LLaMA output
        if not input_folder or not os.path.exists(str(input_folder)):
            print("⚠️ Invalid or no path from LLaMA → using default")
            input_folder = DEFAULT_PATH

        # ✅ Clean path
        input_folder = str(input_folder).strip().replace('"', '')

        if not os.path.exists(input_folder):
            return {"error": f"Invalid path: {input_folder}"}

        print(f"📂 Using folder: {input_folder}")

        # 📁 Create folders
        extract_root = os.path.join(input_folder, "Extract")
        output_folder = os.path.join(input_folder, "Output")

        os.makedirs(extract_root, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)

        # 📦 Get latest ZIP
        latest_zip = get_latest_zip(input_folder)

        if not latest_zip:
            return {"error": "No ZIP file found"}

        print(f"📦 Latest ZIP: {latest_zip}")

        file_name = os.path.basename(latest_zip).replace(".zip", "")

        # 📂 Extract path
        extract_path = os.path.join(extract_root, file_name)
        os.makedirs(extract_path, exist_ok=True)

        # 📦 Extract if needed
        if not os.listdir(extract_path):
            unzip_file(latest_zip, extract_path)
        else:
            print("⏭️ Already extracted")

        # 📂 Report folder
        report_folder = os.path.join(extract_path, "Report")

        if not os.path.exists(report_folder):
            return {"error": "Report folder not found"}

        # 📊 Output file
        output_file = os.path.join(output_folder, file_name + ".xlsx")

        # 📊 Generate Excel
        create_full_excel(report_folder, output_file)

        print(f"🎯 Done: {output_file}")

        return {
            "message": "Report generated successfully",
            "file": output_file
        }

    except Exception as e:
        import traceback
        print("❌ ERROR:", traceback.format_exc())
        return {"error": str(e)}