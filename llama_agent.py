#Author Dibyendu Dey
import json
import re

def decide_input(prompt: str):
    try:
        print("🧠 LLaMA Raw Response:", prompt)

        # =============================
        # ✅ 1. Try JSON extraction
        # =============================
        json_match = re.search(r'\{.*?\}', prompt, re.DOTALL)

        if json_match:
            try:
                json_str = json_match.group(0)
                return json.loads(json_str)
            except Exception as e:
                print("⚠️ JSON parse failed, fallback to regex:", e)

        # =============================
        # 🔥 2. Fallback → Extract path from prompt
        # =============================
        path_match = re.search(r'[a-zA-Z]:\\[^\s]+', prompt)

        if path_match:
            extracted_path = path_match.group(0)
            print(f"✅ Extracted path: {extracted_path}")
            return {"input_path": extracted_path}

        # =============================
        # ⚠️ Nothing found
        # =============================
        print("⚠️ No path found in prompt")
        return {}

    except Exception as e:
        print("❌ decide_input error:", e)
        return {}