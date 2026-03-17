from pathlib import Path
import json

EXTRACTED = Path("extracted")
OUTPUT = Path("output")
OUTPUT.mkdir(exist_ok=True)

def main():
    for js in EXTRACTED.glob("*.json"):
        print(f"Ready for Copilot processing: {js}")
        print(" → Open this file in VS Code")
        print(" → Select entire JSON")
        print(" → Right click → Copilot: Generate using prompt file")
        print(" → Choose: generate-presales-guide.prompt.md")
        print(" → Paste result into:")
        out = OUTPUT / f"{js.stem}_presales.md"
        print(f"   {out}")

if __name__ == "__main__":
    main()