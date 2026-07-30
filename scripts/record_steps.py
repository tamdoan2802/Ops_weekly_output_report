import os
import subprocess
import sys

def main():
    print("==========================================================")
    print(" Launching Playwright Code Generator")
    print("==========================================================")
    print("1. A browser window will open to https://mydaily.myteamsolution.com.vn/")
    print("2. The Playwright Inspector window will also open.")
    print("3. Log in and perform the exact steps to configure and download the report.")
    print("4. Copy the generated Python code from the Inspector.")
    print("5. Paste the relevant steps into Data_loader.py where indicated.")
    print("==========================================================\n")
    
    try:
        # Run the playwright codegen command
        subprocess.run([sys.executable, "-m", "playwright", "codegen", "https://mydaily.myteamsolution.com.vn/"])
    except FileNotFoundError:
        print("Error: Playwright module not found. Please install it using:")
        print("pip install playwright")
        print("playwright install")

if __name__ == "__main__":
    main()
