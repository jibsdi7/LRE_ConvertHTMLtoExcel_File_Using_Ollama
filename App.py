from fastapi import FastAPI, Body
import os
import re

from utils import unzip_file, create_full_excel, compare_multiple_excels, send_email_outlook
from llama_agent import decide_input

app = FastAPI()

DEFAULT_PATH = r"C:\Ahold\Projects\BIDM\Result"


# 📧 Extract emails from prompt
def extract_emails(prompt):
    emails = re.findall(r'[\w\.-]+@[\w\.-]+', prompt)
    return list(set(emails))


# 🎯 Get all ZIP files
def get_zip_files(folder):
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith(".zip")
    ]


# 🎯 Extract count from prompt
def extract_count(prompt):
    match = re.search(r'(\d+)', prompt)
    if match:
        return int(match.group(1))
    return 1


# 🚀 AGENT ENDPOINT
@app.post("/agent")
def run_agent_api(data: dict = Body(...)):
    try:
        prompt = data.get("prompt", "").lower()

        if not prompt:
            return {"error": "Prompt is required"}

        print(f"💬 Prompt: {prompt}")

        # 🧠 LLaMA input parsing
        config = decide_input(prompt)
        input_folder = config.get("input_path")

        # 🔥 Safe fallback
        if not input_folder or not os.path.exists(str(input_folder)):
            print("⚠️ Using default path")
            input_folder = DEFAULT_PATH

        input_folder = str(input_folder).strip().replace('"', '')

        if not os.path.exists(input_folder):
            return {"error": f"Invalid path: {input_folder}"}

        print(f"📂 Using folder: {input_folder}")

        # 📁 Setup folders
        extract_root = os.path.join(input_folder, "Extract")
        output_folder = os.path.join(input_folder, "Output")

        os.makedirs(extract_root, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)

        # 📦 Get ZIP files
        zip_files = get_zip_files(input_folder)

        if not zip_files:
            return {"error": "No ZIP files found"}

        zip_files.sort(key=os.path.getmtime, reverse=True)

        # =============================
        # 🔥 FILE SELECTION
        # =============================
        if "all" in prompt:
            selected_files = zip_files
            print("📊 Processing ALL reports...")
        else:
            count = extract_count(prompt)
            selected_files = zip_files[:count]
            print(f"📊 Processing latest {len(selected_files)} report(s)...")

        output_files = []
        comparison_file = None  # 🔥 important

        # =============================
        # 🔥 PROCESS FILES
        # =============================
        for zip_path in selected_files:
            try:
                file_name = os.path.basename(zip_path).replace(".zip", "")

                extract_path = os.path.join(extract_root, file_name)
                os.makedirs(extract_path, exist_ok=True)

                if not os.listdir(extract_path):
                    unzip_file(zip_path, extract_path)
                else:
                    print(f"⏭️ Already extracted: {file_name}")

                report_folder = os.path.join(extract_path, "Report")

                if not os.path.exists(report_folder):
                    print(f"⚠️ Skipping: {file_name}")
                    continue

                output_file = os.path.join(output_folder, file_name + ".xlsx")

                create_full_excel(report_folder, output_file)

                output_files.append(output_file)

                print(f"✅ Done: {output_file}")

            except Exception as e:
                print(f"❌ Error processing {zip_path}: {e}")

        # =============================
        # 🔥 COMPARE LOGIC
        # =============================
        if "compare" in prompt:
            if len(output_files) < 2:
                return {"error": "Need at least 2 reports to compare"}

            comparison_file = os.path.join(
                output_folder,
                f"comparison_latest_{len(output_files)}.xlsx"
            )

            compare_multiple_excels(output_files, comparison_file)

            print(f"📊 Comparison created: {comparison_file}")

        # =============================
        # 📧 EMAIL LOGIC
        # =============================
        if "email" in prompt or "send" in prompt:
            recipients = extract_emails(prompt)

            if not recipients:
                return {"error": "No recipient email found"}

            file_to_send = None

            # 🔥 PRIORITY: send comparison if exists
            if comparison_file and os.path.exists(comparison_file):
                file_to_send = comparison_file
            elif output_files:
                file_to_send = output_files[0]

            if not file_to_send:
                return {"error": "No file available to send"}

            send_email_outlook(
                subject="LRE Report",
                body="Please find attached report.",
                to=";".join(recipients),
                attachment_path=file_to_send,
                sender="dibyendu.dey@adusa.com"
            )

            return {
                "message": "Email sent successfully",
                "recipients": recipients,
                "file": file_to_send
            }

        # =============================
        # 🎯 NORMAL RESPONSE
        # =============================
        if not output_files:
            return {"error": "No reports processed"}

        if comparison_file:
            return {
                "message": f"Comparison of {len(output_files)} reports generated",
                "file": comparison_file
            }

        if len(output_files) == 1:
            return {
                "message": "Report processed",
                "file": output_files[0]
            }

        return {
            "message": f"{len(output_files)} reports processed",
            "files": output_files
        }

    except Exception as e:
        import traceback
        print("❌ ERROR:", traceback.format_exc())
        return {"error": str(e)}