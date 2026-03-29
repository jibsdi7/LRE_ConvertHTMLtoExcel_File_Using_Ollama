import ollama
import json


def decide_input(prompt):
    try:
        response = ollama.chat(
            model="llama3",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Extract the folder path from the user prompt. "
                        "Return ONLY valid JSON in this format: "
                        '{"input_path": "C:\\\\path\\\\to\\\\folder"} '
                        "If no path is found, return {}"
                    )
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )

        content = response["message"]["content"].strip()

        print("🧠 LLaMA Raw Response:", content)

        # 🔥 Try parsing safely
        try:
            return json.loads(content)
        except:
            # 🔁 Extract JSON manually if extra text present
            start = content.find("{")
            end = content.rfind("}") + 1

            if start != -1 and end != -1:
                json_str = content[start:end]
                return json.loads(json_str)

        return {}

    except Exception as e:
        print(f"❌ LLaMA error: {e}")
        return {}