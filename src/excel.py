import os
import glob
import pandas as pd
import openpyxl
import questionary
from typing import List, Dict, Any, Optional, Tuple

class ExcelProcessor:
    SKIP_COLORS = {
        "FF00FF00", "0000FF00", "00FF00", "#00FF00",
        "FFFF0000", "00FF0000", "FF0000", "#FF0000",
    }
    GREEN_TARGET = (0, 255, 0)
    RED_TARGET = (255, 0, 0)
    COLOR_TOLERANCE = 110

    def __init__(self, boms_dir: str):
        self.boms_dir = boms_dir

    def discover_files(self) -> List[str]:
        search_dir = os.path.join(os.getcwd(), self.boms_dir)
        if not os.path.isdir(search_dir):
            os.makedirs(search_dir, exist_ok=True)
            return []
        return sorted(glob.glob(os.path.join(search_dir, "*.xlsx")))

    def _normalize_color(self, value: Any) -> Optional[str]:
        if not value:
            return None
        raw = str(value).strip()
        if raw.startswith("#"):
            raw = raw[1:]
        if raw.lower().startswith("0x"):
            raw = raw[2:]
        raw = raw.upper()
        if raw == "00000000":
            return None
        if len(raw) == 6:
            return raw
        if len(raw) == 8:
            return raw[2:]
        return None

    def _hex_to_rgb(self, hex_value: str) -> Optional[Tuple[int, int, int]]:
        if not hex_value or len(hex_value) != 6:
            return None
        try:
            return tuple(int(hex_value[i:i+2], 16) for i in (0, 2, 4))
        except ValueError:
            return None

    def _is_close_color(self, rgb: Tuple[int, int, int], target: Tuple[int, int, int]) -> bool:
        distance = sum((c - t) ** 2 for c, t in zip(rgb, target))
        return distance <= self.COLOR_TOLERANCE ** 2

    def _match_skip_color(self, raw: Any) -> Optional[str]:
        normalized = self._normalize_color(raw)
        if normalized in self.SKIP_COLORS:
            if normalized.endswith("00FF00"):
                return "green (done)"
            if normalized.endswith("FF0000"):
                return "red (skip)"

        rgb = self._hex_to_rgb(normalized)
        if rgb is None:
            return None

        if self._is_close_color(rgb, self.GREEN_TARGET):
            return "green (done)"
        if self._is_close_color(rgb, self.RED_TARGET):
            return "red (skip)"
        return None

    def get_cell_color(self, sheet, row: int, col: int = 1) -> Optional[str]:
        try:
            fill = sheet.cell(row=row, column=col).fill
            if not fill or not fill.patternType:
                return None
            color = fill.start_color or fill.fgColor
            if not color:
                return None
            
            raw = None
            if hasattr(color, "rgb") and color.rgb:
                raw = color.rgb
            elif hasattr(color, "index") and color.index:
                raw = color.index
            return self._normalize_color(raw)
        except Exception:
            # Return None if the color value can't be normalized
            return None

    def should_skip_row_color(self, sheet, row: int) -> Optional[str]:
        for col in range(1, min(sheet.max_column, 5) + 1):
            color = self.get_cell_color(sheet, row, col)
            skip = self._match_skip_color(color)
            if skip:
                return skip
        return None

    def process_file(self, filepath: str, run_system: str = "ALL", matcher: Any = None) -> tuple[list[dict[str, Any]], dict[str, int]]:
        target_systems = [run_system.upper()] if run_system and run_system != "ALL" else []
            
        stats = {
            "total_excel_rows": 0,
            "empty_rows": 0,
            "example_rows": 0,
            "system_mismatch": 0,
            "skipped_by_color": 0,
            "valid_parts": 0
        }
        
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        df = pd.read_excel(filepath)
        stats["total_excel_rows"] = len(df)
        
        raw_cols = [str(c).strip().lower() for c in df.columns]
        
        def find_col(aliases):
            for a in aliases:
                if a in raw_cols:
                    return raw_cols.index(a)
            return None

        col_map = {
            "system": find_col(["system", "sys"]),
            "assembly": find_col(["assembly", "asm", "assy"]),
            "part": find_col(["part", "part name", "designation"]),
            "quantity": find_col(["part_quantity", "quantity", "qty", "amount"]),
            "makebuy": find_col(["make o. buy", "m/b", "makebuy"]),
            "comments": find_col(["part_comments", "comments", "notes", "comment"])
        }

        if col_map["system"] is None or col_map["part"] is None:
            raise ValueError(f"Could not identify required 'system' or 'part' columns in {filepath}")

        filtered = []
        for idx, row in df.iterrows():
            excel_row = idx + 2
            sys_val = str(row.iloc[col_map["system"]] if col_map["system"] is not None else "").strip().upper()
            
            # Normalize system code (e.g., "C&B" -> "FR")
            if matcher:
                sys_val = matcher.normalize_system_code(sys_val)
                
            part_val = str(row.iloc[col_map["part"]] if col_map["part"] is not None else "").strip()
            
            if not sys_val or sys_val == "NAN" or not part_val or part_val == "NAN":
                stats["empty_rows"] += 1
                continue

            # Check for "BIS HIER NUR BEISPIEL" marker (means everything before was examples)
            row_content = " ".join(str(val) for val in row.values).upper()
            if "BIS HIER" in row_content and "BEISPIEL" in row_content:
                # Add everything processed so far to examples and reset
                stats["example_rows"] += stats["valid_parts"] + stats["system_mismatch"] + stats["skipped_by_color"] + 1
                stats["valid_parts"] = 0
                stats["system_mismatch"] = 0
                stats["skipped_by_color"] = 0
                filtered = []
                continue

            if any(x in sys_val for x in ["BEISPIEL", "EXAMPLE"]) or \
               any(x in part_val.upper() for x in ["BEISPIEL", "EXAMPLE"]):
                stats["example_rows"] += 1
                continue

            if target_systems and sys_val not in target_systems:
                stats["system_mismatch"] += 1
                continue

            if self.should_skip_row_color(sheet, excel_row):
                stats["skipped_by_color"] += 1
                continue

            asm_val = str(row.iloc[col_map["assembly"]] if col_map["assembly"] is not None else "").strip()
            qty_val = str(row.iloc[col_map["quantity"]] if col_map["quantity"] is not None else "").strip()
            mb_val = str(row.iloc[col_map["makebuy"]] if col_map["makebuy"] is not None else "m").strip().lower()[:1] or "m"
            comm_val = str(row.iloc[col_map["comments"]] if col_map["comments"] is not None else "").strip().replace("nan", "")

            if len(comm_val) > 40:
                print(f"\n[!] Long comment detected at row {excel_row}:")
                print(f"Original: {comm_val}")
                print(f"Length: {len(comm_val)}")
                
                choice = questionary.select(
                    "How would you like to handle this comment?",
                    choices=[
                        "Edit (with pre-fill)",
                        "Truncate to 40 chars",
                        "Skip - do NOT upload"
                    ]
                ).ask()

                if choice == "Edit (with pre-fill)":
                    comm_val = questionary.text(
                        "Enter new comment (<= 40 chars):",
                        default=comm_val,
                        validate=lambda text: len(text) <= 40 or "Must be 40 characters or less"
                    ).ask()
                elif choice == "Truncate to 40 chars":
                    comm_val = comm_val[:40]
                elif choice == "Skip - do NOT upload":
                    with open("bom_log.txt", "a") as log:
                        log.write(f"Row {excel_row}: Action=Skip, Original Comment='{comm_val}'\n")
                    continue
                
                with open("bom_log.txt", "a") as log:
                    log.write(f"Row {excel_row}: Action={choice}, Final Comment='{comm_val}'\n")

            filtered.append({
                "row": excel_row,
                "system": sys_val,
                "assembly": asm_val,
                "part": part_val,
                "makebuy": mb_val,
                "quantity": qty_val,
                "comments": comm_val,
            })
            stats["valid_parts"] += 1

        return filtered, stats
