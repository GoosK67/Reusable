import ollama

# Direct test of the fix
prompt = "Return JSON: {\"group\": \"Executive Summary & Product Overview\", \"reason\": \"test\"}"

response = ollama.chat(
    model='qwen2.5:3b-instruct',
    messages=[{'role': 'user', 'content': prompt}],
    options={'temperature': 0}
)

print("Response object:", type(response))
print("Has message attr:", hasattr(response, 'message'))
print("Message:", response.message if hasattr(response, 'message') else "NO")

# This is the fix we applied
try:
    raw_content = response.message.content if hasattr(response, 'message') else ""
    print("Raw content:", repr(raw_content))
except Exception as e:
    print("Error:", e)
    raw_content = ""

print("Content empty?:", not raw_content or not raw_content.strip())
