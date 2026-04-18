# Release v1.0.0 — Refactor & Tests

Date: 2026-04-18

This release reorganises the project into a proper package, introduces a small CLI entrypoint, improves configuration handling, adds browser automation abstractions, and brings a test-suite and safer defaults. It lays a much more maintainable foundation for future work while keeping the user-facing workflow familiar.

Complete update of every single aspect including:
- UI/UX
- rate-limiting/delay
- reliability
- ease of use and configuration
- and more

Highlights
- Structured package layout under `src/` and a small runner at `main.py`.
- New `Config` class: environment (`.env`) + CLI args + optional `keyring` password storage.
- Browser automation refactor: `src/browser.py` using Playwright for scraping and part creation.
- Added a pytest suite under `.tests/` (unit + integration checks).
- New BOM mappings and system labels in `BOMs/config.yaml` to improve assembly matching.
- Safer defaults in `.env.example`: `DRY_RUN=true` and `TEST_MODE=true`.
- `bom_automation.py` moved to `.old/bom_automation.py` (preserved for compatibility).
- Documentation and README improvements to guide setup and usage.

Notable files (summary)
- [src/__init__.py](src/__init__.py#L1-L3) — package version bumped to `1.1.0`.
- [main.py](main.py) — new lightweight entrypoint.
- [src/config.py](src/config.py) — centralized configuration and CLI parsing.
- [src/browser.py](src/browser.py) — Playwright-based browser helpers.
- [BOMs/config.yaml](BOMs/config.yaml) — assembly remapping + system_map.
- [.env.example](.env.example) — updated and safer defaults.
- [README.md](README.md) — install & usage polish.
- [.tests/](.tests/) — test suite added.

Breaking changes & migration notes
- The primary entrypoint is now `main.py` — call `python main.py` (or keep using the old script at `.old/bom_automation.py`).
- Review `.env.example` and set `TEAM_ID` before running. Consider storing your password in the system keyring with `--set-password`.