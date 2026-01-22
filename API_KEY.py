import os

# ============== 填入你的 KEY ==============
MY_KEY = "AIzaSyAXVGZ9PEV6W_CnPCL8lOxKef1WvmUszHs"
# ========================================

user_home = os.path.expanduser("~")
target_file = os.path.join(user_home, ".continue", "config.json")

# 使用你列表里存在的 gemini-2.5-pro 和 gemini-2.5-flash
config_content = f'''{{
  "models": [
    {{
      "title": "Gemini 2.5 Pro (最新版)",
      "provider": "gemini",
      "model": "gemini-2.5-pro",
      "apiKey": "{MY_KEY}"
    }},
    {{
      "title": "Gemini 2.5 Flash (极速)",
      "provider": "gemini",
      "model": "gemini-2.5-flash",
      "apiKey": "{MY_KEY}"
    }}
  ],
  "tabAutocompleteModel": {{
    "title": "Autocomplete",
    "provider": "gemini",
    "model": "gemini-2.5-flash",
    "apiKey": "{MY_KEY}"
  }},
  "requestOptions": {{
    "proxy": "http://127.0.0.1:10809"
  }}
}}'''

try:
    with open(target_file, "w", encoding="utf-8") as f:
        f.write(config_content)
    print(f"✅ 成功升级！配置文件已更新为 Gemini 2.5: {target_file}")
    print("👉 现在重启 PyCharm，选择 'Gemini 2.5 Pro'，肯定能通！")
except Exception as e:
    print(f"❌ 写入失败: {e}")