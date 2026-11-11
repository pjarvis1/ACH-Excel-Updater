#!/usr/bin/env python3
"""
Read app.toml and emit build settings for GitHub Actions.

Writes key=value lines to $GITHUB_OUTPUT so later jobs can use
${{ needs.config.outputs.<key> }}. Run it locally with no environment
variable set and it just prints what it would emit, which makes it easy
to debug without pushing.

tomllib is in the standard library from Python 3.11 onward, so this has
no dependencies.
"""

from __future__ import annotations

import json
import os
import sys
import tomllib
from pathlib import Path

CONFIG_PATH = Path(__file__).resolve().parent.parent / "app.toml"

VALID_TARGETS = {
    "windows": "windows-latest",
    "macos-apple-silicon": "macos-latest",
    "macos-intel": "macos-15-intel",
    "linux": "ubuntu-latest",
}

ILLEGAL_NAME_CHARS = set(' /\\:*?"<>|')


def fail(message: str) -> None:
    print(f"::error file=app.toml::{message}")
    sys.exit(1)


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        fail(f"Config file not found at {CONFIG_PATH}")
    try:
        return tomllib.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except tomllib.TOMLDecodeError as exc:
        fail(f"app.toml is not valid TOML: {exc}")
    return {}


def build_outputs(cfg: dict) -> dict[str, str]:
    app = cfg.get("app", {})
    pyi = cfg.get("pyinstaller", {})
    build = cfg.get("build", {})

    # ---- app name -------------------------------------------------
    name = str(app.get("name", "")).strip()
    if not name:
        fail("app.name is required")
    bad = ILLEGAL_NAME_CHARS.intersection(name)
    if bad:
        fail(f"app.name cannot contain {sorted(bad)} -- got {name!r}")

    entrypoint = str(app.get("entrypoint", "app.py"))
    if not (CONFIG_PATH.parent / entrypoint).exists():
        fail(f"app.entrypoint points at {entrypoint}, which does not exist")

    python_version = str(app.get("python_version", "3.12"))

    # ---- assemble the PyInstaller command --------------------------
    onefile = bool(pyi.get("onefile", True))
    windowed = bool(pyi.get("windowed", True))

    args: list[str] = ["--clean", "--noconfirm", "--name", name]
    args.append("--onefile" if onefile else "--onedir")
    if windowed:
        args.append("--windowed")

    for module in pyi.get("hidden_imports", []):
        args += ["--hidden-import", str(module)]
    for module in pyi.get("exclude_modules", []):
        args += ["--exclude-module", str(module)]
    for extra in pyi.get("extra_args", []):
        args.append(str(extra))

    # NOTE: entrypoint is deliberately NOT appended here. The workflow adds it
    # last, after the per-platform icon flag, so the positional script argument
    # always comes at the very end of the command.

    # ---- per-platform icon (optional) ------------------------------
    # Windows needs a .ico, macOS needs a .icns -- they can't share one flag,
    # so each platform gets its own, picked here. Missing files are simply
    # skipped, so the template still works for apps with no icon.
    assets = CONFIG_PATH.parent / "assets"
    have_ico = (assets / "icon.ico").exists()
    have_icns = (assets / "icon.icns").exists()

    def icon_arg_for(label: str) -> str:
        if label == "windows" and have_ico:
            return "--icon=assets/icon.ico"
        if label.startswith("macos") and have_icns:
            return "--icon=assets/icon.icns"
        return ""

    # ---- what each platform should package -------------------------
    # --windowed always produces a .app bundle on macOS, in both
    # onefile and onedir mode. Windows gets an .exe or a folder.
    windows_target = f"dist/{name}.exe" if onefile else f"dist/{name}"
    macos_target = f"dist/{name}.app" if windowed else f"dist/{name}"

    # ---- which runners to spin up ----------------------------------
    targets = build.get("targets", list(VALID_TARGETS)[:3])
    unknown = [t for t in targets if t not in VALID_TARGETS]
    if unknown:
        fail(f"Unknown build.targets entries: {unknown}. Valid: {sorted(VALID_TARGETS)}")
    if not targets:
        fail("build.targets is empty -- nothing to build")

    matrix = {
        "include": [
            {"os": VALID_TARGETS[t], "label": t, "icon_arg": icon_arg_for(t)}
            for t in targets
        ]
    }

    return {
        "app_name": name,
        "display_name": str(app.get("display_name", name)),
        "entrypoint": entrypoint,
        "python_version": python_version,
        "pyi_args": " ".join(args),
        "windows_target": windows_target,
        "macos_target": macos_target,
        "retention_days": str(build.get("artifact_retention_days", 3)),
        "matrix": json.dumps(matrix),
    }


def main() -> None:
    outputs = build_outputs(load_config())

    github_output = os.environ.get("GITHUB_OUTPUT")
    for key, value in outputs.items():
        print(f"{key}={value}")
        if github_output:
            with open(github_output, "a", encoding="utf-8") as handle:
                handle.write(f"{key}={value}\n")


if __name__ == "__main__":
    main()