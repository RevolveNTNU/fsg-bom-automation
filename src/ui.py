import threading
from datetime import datetime
from rich.console import Console, Group
from rich.panel import Panel
from rich.table import Table
from rich.live import Live
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn

class UI:
    def __init__(self, log_file: str):
        self.console = Console()
        self.log_file = log_file

    def log(self, message: str, status: str = "INFO") -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        # Console output colors
        color = {
            "INFO": "white",
            "OK": "green",
            "SKIP": "blue",
            "WARN": "yellow",
            "ERROR": "red",
            "DRY": "magenta",
            "WAIT": "cyan"
        }.get(status, "white")
        
        # Pretty console logging
        self.console.print(f"[dim]{ts}[/] [{color}]{status:^5}[/] {message}")
        
        # Log file formatting
        # Using a structured format: TIMESTAMP | STATUS | MESSAGE
        with open(self.log_file, "a", encoding="utf-8") as f:
            log_entry = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {status:^5} | {message}\n"
            f.write(log_entry)

    def prompt_ask(self, question):
        result = [None]
        def _target(): result[0] = question.ask()
        thread = threading.Thread(target=_target, daemon=True)
        thread.start()
        try:
            while thread.is_alive():
                thread.join(timeout=0.1)
        except KeyboardInterrupt:
            raise
        return result[0]

    def show_summary(self, count: int, filename: str, system: str, test_mode: bool, dry_run: bool):
        table = Table(title="Upload Configuration", show_header=False)
        table.add_row("File", filename)
        table.add_row("System", system)
        table.add_row("Parts", str(count))
        table.add_row("Test Mode", "[yellow]ENABLED[/]" if test_mode else "[green]DISABLED[/]")
        table.add_row("Dry Run", "[yellow]ENABLED[/]" if dry_run else "[green]DISABLED[/]")
        self.console.print("\n", Panel(table, title="Ready to Upload", expand=False))

    def create_dashboard(self, total: int):
        status_table = Table(title="Live Upload Status", box=None)
        status_table.add_column("Row", justify="right", style="cyan")
        status_table.add_column("Part Name", style="white")
        status_table.add_column("Status", justify="center")
        status_table.add_column("Message", style="dim")

        progress = Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TextColumn("• [blue]ETA: {task.fields[eta]}"),
            expand=True
        )
        task_id = progress.add_task("Uploading parts...", total=total, eta="Calculating...")
        
        live = Live(Panel(Group(status_table, progress), title="BOM Upload Dashboard"), refresh_per_second=4)
        return live, status_table, progress, task_id

    def update_eta(self, progress, task_id, start_time, current_idx, total):
        import time
        elapsed = time.time() - start_time
        if current_idx > 0:
            avg_per_item = elapsed / current_idx
            remaining = total - current_idx
            eta_secs = int(avg_per_item * remaining)
            mins, secs = divmod(eta_secs, 60)
            eta_str = f"{mins}m {secs}s" if mins > 0 else f"{secs}s"
            progress.update(task_id, eta=eta_str)
