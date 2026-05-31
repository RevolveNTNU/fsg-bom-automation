import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.config import Config
from src.browser import FSGBrowser

def main():
    config = Config()
    config.load()
    
    print("Connecting to browser...")
    with FSGBrowser(config) as browser:
        print("Logging in...")
        if not browser.login():
            print("Login failed or needs manual input")
        
        print("Navigating to BOM...")
        browser.goto_bom()
        
        # Click new to open dialog
        print("Opening 'New' dialog...")
        browser.page.get_by_text("New", exact=True).click()
        browser.page.wait_for_selector(".DTE_Action_Create", timeout=15000)
        
        # Get systems
        systems = browser.page.eval_on_selector(
            "#DTE_Field_system",
            "el => Array.from(el.options).map(o => ({value: o.value, text: o.text}))"
        )
        print("\n=== SYSTEMS ON WEBSITE ===")
        for sys in systems:
            print(f"Value: '{sys['value']}' | Text: '{sys['text']}'")
            
        print("\n=== SYSTEM MAPPINGS IN CONFIG ===")
        import yaml
        with open("BOMs/config.yaml") as f:
            data = yaml.safe_load(f)
            sys_map = data.get('system_map', {})
            for k, v in sys_map.items():
                print(f"Config Key: '{k}' | Config Label: '{v}'")

if __name__ == "__main__":
    main()
