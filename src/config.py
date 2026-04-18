import os
import argparse
import keyring
from dotenv import load_dotenv
from dataclasses import dataclass, field
from typing import List

@dataclass
class Config:
    team_id: str = ""
    username: str = ""
    password: str = ""
    test_mode: bool = True
    test_limit: int = 3
    dry_run: bool = False
    default_system: str = ""
    allowed_assemblies: List[str] = field(default_factory=list)
    log_file: str = "bom_log.txt"
    boms_dir: str = "BOMs"
    
    # Timing & Robustness
    base_delay: float = 2.0
    burst_limit: int = 10
    burst_cooldown: float = 15.0
    max_retries: int = 2
    request_timeout: float = 30000.0 # ms
    
    part_max_length: int = 25
    comments_max_length: int = 40

    def load(self):
        # Explicitly reload from .env
        load_dotenv(override=True)
        
        # 1. Credentials
        self.team_id = os.getenv("TEAM_ID", self.team_id).strip()
        self.username = os.getenv("FSG_USERNAME", self.username).strip()
        self.password = os.getenv("FSG_PASSWORD", self.password).strip()
        
        # Try keyring if not in env
        if not self.password and self.username:
            stored = keyring.get_password("fsg_bom_automation", self.username)
            if stored:
                self.password = stored

        # 2. Boolean flags
        if os.getenv("TEST_MODE") is not None:
            self.test_mode = os.getenv("TEST_MODE", "").lower() == "true"
        if os.getenv("DRY_RUN") is not None:
            self.dry_run = os.getenv("DRY_RUN", "").lower() == "true"
        
        # 3. Numeric values
        try:
            self.test_limit = int(os.getenv("TEST_LIMIT", self.test_limit))
        except Exception:
            pass
        
        try:
            self.base_delay = float(os.getenv("BASE_DELAY", self.base_delay))
        except Exception:
            pass
        
        try:
            self.burst_limit = int(os.getenv("BURST_LIMIT", self.burst_limit))
        except Exception:
            pass
        
        try:
            self.burst_cooldown = float(os.getenv("BURST_COOLDOWN", self.burst_cooldown))
        except Exception:
            pass
        
        try:
            self.max_retries = int(os.getenv("MAX_RETRIES", self.max_retries))
        except Exception:
            pass
        
        try:
            self.request_timeout = float(os.getenv("REQUEST_TIMEOUT", str(self.request_timeout/1000.0))) * 1000
        except Exception:
            pass

        # 4. Strings & Lists
        self.default_system = os.getenv("DEFAULT_SYSTEM", self.default_system).strip().upper()
        
        allowed = os.getenv("ALLOWED_ASSEMBLIES", "").strip()
        if allowed:
            self.allowed_assemblies = [a.strip() for a in allowed.split(",") if a.strip()]
            
        self.log_file = os.getenv("LOG_FILE", self.log_file)
        self.boms_dir = os.getenv("BOMS_DIR", self.boms_dir)

    def parse_args(self):
        parser = argparse.ArgumentParser(description="FSG CCBOM Automation Tool")
        parser.add_argument("--team-id", help="FSG Team ID")
        parser.add_argument("--username", help="FSG Username")
        parser.add_argument("--password", help="FSG Password")
        parser.add_argument("--set-password", action="store_true", help="Store password in system keyring")
        parser.add_argument("--test-mode", action="store_true", default=None, help="Enable test mode (limit uploads)")
        parser.add_argument("--no-test-mode", action="store_false", dest="test_mode", help="Disable test mode")
        parser.add_argument("--dry-run", action="store_true", default=None, help="Simulate uploads without hitting the server")
        parser.add_argument("--limit", type=int, help="Upload limit for test mode")
        parser.add_argument("--system", help="Pre-select system code (e.g. BR, DT)")
        
        args = parser.parse_args()
        
        if args.team_id:
            self.team_id = args.team_id
        if args.username:
            self.username = args.username
        if args.password:
            self.password = args.password
        
        # Only override if explicit CLI flag is used (not None)
        if args.test_mode is not None:
            self.test_mode = args.test_mode
        if args.dry_run is not None:
            self.dry_run = args.dry_run
        
        if args.limit:
            self.test_limit = args.limit
        if args.system:
            self.default_system = args.system.upper()
        
        if args.set_password:
            self._store_password()
            print("Password stored in keyring. You can now run without --password.")
            exit(0)

    def _store_password(self):
        if not self.username:
            self.username = input("Enter FSG Username: ").strip()
        import getpass
        pwd = getpass.getpass(f"Enter FSG Password for {self.username}: ")
        keyring.set_password("fsg_bom_automation", self.username, pwd)

    @property
    def base_url(self):
        return "https://www.formulastudent.de"

    @property
    def login_url(self):
        return f"{self.base_url}/login"

    @property
    def bom_url(self):
        if not self.team_id:
            return ""
        return f"{self.base_url}/teams/fse/details/bom/tid/{self.team_id}"
