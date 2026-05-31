import time
from typing import List, Dict, Optional
from playwright.sync_api import sync_playwright

class FSGBrowser:
    def __init__(self, config):
        self.config = config
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None

    def __enter__(self):
        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.launch(headless=False)
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.browser:
            self.browser.close()
        if self.pw:
            self.pw.stop()

    def login(self):
        if not self.config.username or not self.config.password:
            return False
        self.page.goto(self.config.login_url)
        self.page.fill("#tx-felogin-input-username", self.config.username)
        self.page.fill("#tx-felogin-input-password", self.config.password)
        self.page.click('input[name="submit"]')
        self.page.wait_for_load_state("networkidle")
        return True

    def goto_bom(self):
        self.page.goto(self.config.bom_url)

    def fetch_site_options(self, system_label: Optional[str] = None) -> List[str]:
        try:
            self.page.get_by_text("New", exact=True).click()
            self.page.wait_for_selector(".DTE_Action_Create", timeout=5000)
            
            if system_label:
                self.page.locator("#DTE_Field_system").select_option(label=system_label)
                self.page.locator("#DTE_Field_system").dispatch_event("change")
                time.sleep(0.5)

            options = self.page.eval_on_selector(
                "#DTE_Field_assembly",
                "el => Array.from(el.options).map(o => o.text)",
            ) or []
            self.page.keyboard.press("Escape")
            return options
        except Exception:
            return []

    def scrape_existing_parts(self, matcher) -> Dict[str, Dict]:
        try:
            self.page.wait_for_selector("#bom-table", timeout=10000)
            
            # Try to show "All" entries if it's a DataTables table
            try:
                if self.page.locator("select[name='bom-table_length']").is_visible():
                    self.page.locator("select[name='bom-table_length']").select_option("-1")
                    time.sleep(3.0) # Wait for table to reload with all entries
            except Exception:
                pass

            data = self.page.evaluate("""() => {
                const results = [];
                const table = document.querySelector('#bom-table');
                if (!table) return [];

                // 1. Get header mapping
                const ths = Array.from(table.querySelectorAll('thead th'));
                const headers = ths.map(th => th.innerText.toLowerCase().trim());
                
                const findIdx = (aliases) => {
                    return headers.findIndex(h => aliases.some(a => h.includes(a)));
                };
                
                const idxMap = {
                    system: findIdx(['system', 'sys']),
                    assembly: findIdx(['assembly', 'asm', 'assy']),
                    part: findIdx(['part', 'name', 'designation', 'description']),
                    id_col: findIdx(['id', 'part-id']),
                };

                // 2. Scrape all rows with state for grouped rows
                let lastSystem = "";
                let lastAssembly = "";

                const rows = table.querySelectorAll('tbody tr');
                rows.forEach(tr => {
                    if (tr.classList.contains('empty') || tr.innerText.includes('No data')) return;
                    
                    const cells = tr.querySelectorAll('td');
                    if (cells.length < 4) return;

                    let currentSys = idxMap.system !== -1 ? cells[idxMap.system].innerText.trim() : "";
                    let currentAsm = idxMap.assembly !== -1 ? cells[idxMap.assembly].innerText.trim() : "";
                    let currentPart = idxMap.part !== -1 ? cells[idxMap.part].innerText.trim() : "";
                    let currentIdVal = idxMap.id_col !== -1 ? cells[idxMap.id_col].innerText.trim() : "";

                    // If part cell is empty but it's a part row, try to find ANY non-empty cell after assembly
                    if (!currentPart && tr.id.startsWith('bompart_')) {
                         for (let i = Math.max(idxMap.assembly, 2) + 1; i < cells.length; i++) {
                             const txt = cells[i].innerText.trim();
                             if (txt && txt.length > 1) {
                                 currentPart = txt;
                                 break;
                             }
                         }
                    }

                    // Update state if we found a new system/assembly (often in bomassembly_ rows)
                    if (currentSys && currentSys.length > 1) lastSystem = currentSys;
                    if (currentAsm && currentAsm.length > 1) lastAssembly = currentAsm;

                    // A row is a PART row if it has an ID starting with 'bompart_'
                    if (tr.id && tr.id.startsWith('bompart_')) {
                        results.push({
                            row_id: tr.id,
                            site_id: currentIdVal,
                            system: lastSystem,
                            assembly: lastAssembly,
                            part: currentPart
                        });
                    }
                });
                return results;
            }""")
            
            existing = {}
            for r in data:
                # Extract System Code
                sys = ""
                # 1. Try from the System column text (e.g. "FR - Chassis" -> "FR")
                if r.get('system'):
                    sys = r['system'].split(' ')[0].strip().upper()
                
                # 2. Try from the ID column (e.g. '359-AT-00008' -> 'AT')
                if not sys and r.get('site_id'):
                    parts = str(r['site_id']).split('-')
                    if len(parts) >= 2:
                        sys = parts[1].upper()
                
                # 3. Try from the Row ID if it follows a format like 'DT_12345'
                if not sys and r.get('row_id'):
                    row_id_parts = str(r['row_id']).split('_')
                    if len(row_id_parts) >= 2 and len(row_id_parts[0]) <= 3:
                        sys = row_id_parts[0].upper()

                key = matcher.canonical_key(sys, r.get('assembly') or "", r.get('part') or "")
                existing[key] = r
                
            return existing
        except Exception as e:
            print(f"Error scraping existing parts: {e}")
            return {}

    def create_part(self, item: Dict):
        # Clean start: if a dialog is already open (from a previous failure), try to close it
        if self.page.locator(".DTE_Action_Create").is_visible():
            self.page.keyboard.press("Escape")
            time.sleep(0.5)
            # If still visible, it might be a confirm dialog or stuck
            if self.page.locator(".DTE_Action_Create").is_visible():
                self.page.get_by_text("Cancel").click(force=True)
                time.sleep(0.5)

        self.page.get_by_text("New", exact=True).click()
        self.page.wait_for_selector(".DTE_Action_Create", timeout=10000)
        
        # Select system and trigger change event
        self.page.locator("#DTE_Field_system").select_option(label=item['system_label'])
        self.page.locator("#DTE_Field_system").dispatch_event("change")
        
        # Wait for the specific assembly option to be loaded in the dropdown
        try:
            target_asm = item['assembly']
            self.page.wait_for_function(
                """(asm) => {
                    const el = document.querySelector("#DTE_Field_assembly");
                    return el && Array.from(el.options).some(o => o.text === asm);
                }""",
                target_asm,
                timeout=5000
            )
        except Exception:
            # If it doesn't appear in 5s, let the next step try (and potentially fail with a better error)
            pass
            
        self.page.locator("#DTE_Field_assembly").select_option(label=item['assembly'])
        self.page.locator("#DTE_Field_part").fill(item['part'])
        
        if item['makebuy'] == 'm':
            self.page.locator("#DTE_Field_makebuy_0").check()
        else:
            self.page.locator("#DTE_Field_makebuy_1").check()
        
        if item['comments']:
            self.page.locator("#DTE_Field_comments").fill(item['comments'])
        if item['quantity']:
            self.page.locator("#DTE_Field_quantity").fill(item['quantity'])
        
        self.page.get_by_text("Create", exact=True).click()
        # Increased timeout to 20s as server can be slow ("processing" state)
        self.page.wait_for_selector(".DTE_Action_Create", state="hidden", timeout=20000)
