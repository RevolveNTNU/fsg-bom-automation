import re
import yaml
import os
from typing import List, Optional

class AssemblyMatcher:
    def __init__(self, config_path: str = "BOMs/config.yaml"):
        self.load_config(config_path)

    def load_config(self, path: str):
        if not os.path.exists(path):
            raise FileNotFoundError(f"Configuration file not found: {path}")
        with open(path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
            self.ASSEMBLY_REMAP = data.get('assembly_mappings', {})
            self.SYSTEM_MAP = data.get('system_map', {})

    @staticmethod
    def _normalize(s: str) -> str:
        return re.sub(r"\W+", "", str(s or "").lower())

    def canonical_key(self, system: str, assembly: str, part: str) -> str:
        s = system.strip().upper()
        a = self._normalize(assembly)
        p = self._normalize(part)
        return f"{s}_{a}_{p}"

    def resolve_label(self, target: str, site_options: List[str], allowed: Optional[List[str]] = None) -> Optional[str]:
        target_clean = target.strip()
        resolved = self.ASSEMBLY_REMAP.get(target_clean.lower(), target_clean).strip()
        target_lower = resolved.lower()
        target_norm = self._normalize(resolved)

        options = site_options
        if allowed:
            allowed_set = {a.strip().lower() for a in allowed if a.strip()}
            options = [opt for opt in site_options if opt.strip().lower() in allowed_set]

        if not options:
            return None

        # 1. Exact match
        for opt in options:
            if opt.strip() == resolved:
                return opt

        # 2. Case-insensitive
        for opt in options:
            if opt.strip().lower() == target_lower:
                return opt

        # 3. Substring
        for opt in options:
            ol = opt.strip().lower()
            if target_lower in ol or ol in target_lower:
                return opt

        # 4. Normalized
        for opt in options:
            if self._normalize(opt) == target_norm:
                return opt

        # 5. Normalized substring
        for opt in options:
            on = self._normalize(opt)
            if target_norm in on or on in target_norm:
                return opt

        return None

    def get_system_label(self, code: str) -> str:
        return self.SYSTEM_MAP.get(code.upper(), code.upper())
