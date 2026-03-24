import ollama
import json

# Test what Ollama actually returns
prompt = "Hello, what is 2+2?"

print("Testing Ollama response format...")
response = ollama.chat(
    model='qwen2.5:3b-instruct',
    messages=[{'role': 'user', 'content': prompt}],
    options={'temperature': 0}
)

print("\n=== RESPONSE TYPE ===")
print(f"Type: {type(response)}")
print(f"Class name: {response.__class__.__name__}")

# Try accessing as object attribute
print("\n=== ACCESSING AS OBJECT ===")
print(f"response.message: {response.message}")
print(f"response.message.content: {response.message.content}")

# Try converting to dict
print("\n=== CONVERTING TO DICT ===")
try:
    response_dict = response.model_dump() if hasattr(response, 'model_dump') else vars(response)
    print(json.dumps(response_dict, indent=2, default=str)[:500])
except Exception as e:
    print(f"Error: {e}")
    # Try another approach
    print(f"Dir: {[x for x in dir(response) if not x.startswith('_')]}")

