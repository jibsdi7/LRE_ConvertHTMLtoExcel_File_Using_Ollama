import ollama
import json


def decide_input(prompt):
    try:
        response = ollama.chat(
            model="llama3",
            messages=[
                {
                    "role": "system",
                    "content": "Extract the folder path from the user prompt. Return JSON with key 'input_path'."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )

        content = response["message"]["content"]

        # Extract JSON
        start = content.find("{")
        end = content.rfind("}") + 1
        json_str = content[start:end]

        return json.loads(json_str)

    except Exception as e:
        print(f"❌ LLaMA error: {e}")
        return {}