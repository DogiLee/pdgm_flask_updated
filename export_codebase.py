#!/usr/bin/env python3
"""PDGM codebase'ini tek Markdown dosyasina dump eder (GPT / Claude icin).

Kullanim:
  python3 export_codebase.py
  python3 export_codebase.py -o codebase_export.md
  python3 export_codebase.py --docs   # docs/ dahil
  python3 export_codebase.py --all-md # tum .md raporlari dahil

Varsayilan olarak .env, data/, .venv ve binary dosyalar HARIC tutulur.
"""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path

ROOT = Path(__file__).resolve().parent

SKIP_DIRS = {
    ".venv",
    "venv",
    "__pycache__",
    ".git",
    "node_modules",
    "data",
    ".cursor",
    ".idea",
    ".vscode",
}

SKIP_FILES = {
    ".env",
    "gizli.key",
    "kullanicilar.json",
    "sunucu.lock",
    "uygulama.log",
    "export_codebase.py",  # kendini dump etme (istege bagli degistirilebilir)
}

SKIP_SUFFIXES = {
    ".pyc",
    ".pyo",
    ".xlsx",
    ".xls",
    ".xlsm",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".webp",
    ".ico",
    ".pdf",
    ".zip",
    ".gz",
    ".7z",
    ".exe",
    ".dll",
    ".so",
    ".dylib",
    ".lock",
}

# Buyuk onceki dump / analiz dosyalari — --all-md olmadan atla
LARGE_MD_SKIP = {
    "codebase.md",
    "claude_code_change.md",
    "claude_analiz.md",
    "HARDENING_DEGISIKLIK_OZETI.md",
}

LANG = {
    ".py": "python",
    ".html": "html",
    ".htm": "html",
    ".css": "css",
    ".js": "javascript",
    ".json": "json",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".ps1": "powershell",
    ".sh": "bash",
    ".toml": "toml",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".ini": "ini",
    ".cfg": "ini",
    ".example": "dotenv",
    ".gitignore": "gitignore",
}


def fence_lang(path: Path) -> str:
    if path.name == ".gitignore":
        return "gitignore"
    if path.name.endswith(".env.example") or path.name == ".env.example":
        return "dotenv"
    return LANG.get(path.suffix.lower(), "")


def should_skip(path: Path, include_docs: bool, all_md: bool) -> bool:
    rel = path.relative_to(ROOT)
    parts = set(rel.parts)

    if parts & SKIP_DIRS:
        return True
    if path.name in SKIP_FILES:
        return True
    if path.suffix.lower() in SKIP_SUFFIXES:
        return True
    if path.name.startswith("codebase") and path.suffix == ".md" and path.name != "README.md":
        if not all_md:
            return True
    if path.name in LARGE_MD_SKIP and not all_md:
        return True
    if rel.parts and rel.parts[0] == "docs" and not include_docs and not all_md:
        return True
    return False


def collect_files(include_docs: bool, all_md: bool) -> list[Path]:
    files: list[Path] = []
    for path in sorted(ROOT.rglob("*")):
        if not path.is_file():
            continue
        if should_skip(path, include_docs, all_md):
            continue
        # yalnizca text dosyalari
        try:
            path.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue
        files.append(path)
    return files


def group_key(path: Path) -> str:
    rel = path.relative_to(ROOT).as_posix()
    if path.suffix == ".py":
        return "PYTHON"
    if "templates/" in rel or path.suffix in {".html", ".htm"}:
        return "TEMPLATES"
    if "static/" in rel or path.suffix == ".css":
        return "STATIC"
    if path.suffix == ".md" or path.parts[-2:] == ("docs", path.name):
        return "DOCS"
    if path.suffix in {".bat", ".ps1", ".sh"}:
        return "SCRIPTS"
    return "CONFIG / DIGER"


def write_dump(out: Path, files: list[Path]) -> None:
    by_group: dict[str, list[Path]] = {}
    for f in files:
        by_group.setdefault(group_key(f), []).append(f)

    order = ["PYTHON", "TEMPLATES", "STATIC", "SCRIPTS", "CONFIG / DIGER", "DOCS"]
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    lines: list[str] = []
    lines.append("# Repository Codebase Dump\n")
    lines.append(f"Olusturulma tarihi: {now}\n")
    lines.append(f"Repository: `{ROOT.name}`\n")
    lines.append(f"Toplam dosya: **{len(files)}**\n")
    lines.append("")
    lines.append("---\n")
    lines.append("\n## Dosya Listesi\n")
    lines.append("```text")
    for g in order:
        group = by_group.get(g) or []
        if not group:
            continue
        lines.append(g)
        for f in group:
            lines.append(f"├── {f.relative_to(ROOT).as_posix()}")
        lines.append("")
    lines.append("```\n")
    lines.append("\n---\n")

    section = 1
    for g in order:
        group = by_group.get(g) or []
        if not group:
            continue
        lines.append(f"\n# {section}. {g}\n")
        section += 1
        for f in group:
            rel = f.relative_to(ROOT).as_posix()
            lang = fence_lang(f)
            content = f.read_text(encoding="utf-8")
            if not content.endswith("\n"):
                content += "\n"
            lines.append(f"\n## `{rel}`\n")
            lines.append(f"\n```{lang}")
            lines.append(content.rstrip("\n"))
            lines.append("```\n")

    out.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="PDGM codebase Markdown export")
    parser.add_argument(
        "-o",
        "--output",
        default="codebase_export.md",
        help="Cikti dosyasi (varsayilan: codebase_export.md)",
    )
    parser.add_argument(
        "--docs",
        action="store_true",
        help="docs/ klasorunu dahil et",
    )
    parser.add_argument(
        "--all-md",
        action="store_true",
        help="Buyuk MD raporlarini ve docs/ klasorunu dahil et",
    )
    parser.add_argument(
        "--include-self",
        action="store_true",
        help="export_codebase.py dosyasini da dahil et",
    )
    args = parser.parse_args()

    if args.include_self:
        SKIP_FILES.discard("export_codebase.py")

    include_docs = args.docs or args.all_md
    files = collect_files(include_docs=include_docs, all_md=args.all_md)
    out = Path(args.output)
    if not out.is_absolute():
        out = ROOT / out

    write_dump(out, files)
    size_kb = out.stat().st_size / 1024
    print(f"OK: {len(files)} dosya -> {out}")
    print(f"Boyut: {size_kb:.1f} KB")
    print("Not: .env, data/, .venv ve secret dosyalar HARIC.")


if __name__ == "__main__":
    main()
