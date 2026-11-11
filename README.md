# App Distribution Template

Turn a Python script into a double-clickable app for Windows and macOS, built
automatically by GitHub Actions and published to a download page.

Your users install nothing. No Python, no `pip`, no command line.

---

## Why this exists

PyInstaller can only build for the platform it runs on — you cannot make a Mac
app from a Windows machine. This template solves that by letting GitHub's
runners do the building: one Windows machine, one Apple Silicon Mac, one Intel
Mac, all in parallel, every time you tag a version.

Everything configurable lives in one file, `app.toml`. You should never need to
edit the workflow.

---

## Quick start

1. Click **Use this template → Create a new repository** at the top of this page
2. Replace `app.py` with your script
3. List your dependencies in `requirements.txt`
4. Set `name` and `display_name` in `app.toml`
5. Update the `icon.ico` and `icon.icns` files in `/assets`
6. Commit, push, then:

```bash
git tag v1.0.0
git push origin v1.0.0
```

A few minutes later your **Releases** page has three downloads.

Check your config before pushing:

```bash
python scripts/build_config.py
```

That prints exactly what the workflow will use and fails loudly on common
mistakes — a name with spaces, a missing entrypoint, an unknown build target.

---

## What's in the repo

```
├── app.py                      ← you replace this with your program (keeping the name app.py)
├── requirements.txt            ← you configure this file with your dependencies
├── app.toml                    ← you configure this with your app name
├── scripts/
│   └── build_config.py         reads app.toml, feeds the workflow
└── .github/workflows/
    └── build.yml               don't edit this
```

---

## Configuring `app.toml`

### `[app]`

| Key | What it does |
|---|---|
| `name` | Executable name. **No spaces, slashes, or colons.** Becomes `YourApp.exe` / `YourApp.app` |
| `display_name` | Friendly name, used in release titles. Spaces fine |
| `entrypoint` | The Python file PyInstaller starts from |
| `python_version` | Version the runners install |

### `[pyinstaller]`

| Key | What it does |
|---|---|
| `onefile` | `true` = one clickable file, slower to start. `false` = a folder, starts fast |
| `windowed` | `true` = no console window; required for GUI apps and produces a real `.app` on macOS |
| `hidden_imports` | Modules PyInstaller can't detect on its own — see [below](#the-build-passed-but-the-app-wont-start) |
| `exclude_modules` | Modules to leave out, to shrink the download |
| `extra_args` | Any other PyInstaller flag, one per entry |

### `[build]`

| Key | What it does |
|---|---|
| `targets` | Which platforms to build. Valid: `windows`, `macos-apple-silicon`, `macos-intel`, `linux` |
| `artifact_retention_days` | How long intermediate build files are kept |

### Common adjustments

| Want to... | Change |
|---|---|
| Skip Mac builds entirely | Remove entries from `build.targets` |
| Fix a "module not found" crash | Add to `hidden_imports` |
| Shrink the download | Add to `exclude_modules` |
| Faster startup | `onefile = false` |
| Console app instead of GUI | `windowed = false` |

---

## Running a build

**Test run, no version number:**
Actions tab → **Build** → **Run workflow**. The zips appear under **Artifacts**
at the bottom of the run page. No Release is created.

**Real release:**

```bash
git tag v1.0.0
git push origin v1.0.0
```

The tag must start with a lowercase `v`. `1.0.0` won't trigger anything.

Need to redo a tag?

```bash
git tag -d v1.0.0
git push origin :refs/tags/v1.0.0
git tag v1.0.0
git push origin v1.0.0
```

### How it fits together

```
app.toml → config job → build job (one per platform) → release job
```

The `config` job reads `app.toml` once and publishes the values as job outputs.
Everything downstream — including which runners start — comes from there.

---

## Sending it to your users

Your Releases page is public and needs no login:

```
https://github.com/USERNAME/REPO/releases
```

GitHub also gives you a permanent link to the newest version of each file, so
you can give one URL to your users and never update it:

```
https://github.com/USERNAME/REPO/releases/latest/download/YourApp-windows.zip
```

> **Private repos:** release assets are private too. Downloaders must be
> repo collaborators *and* signed in. For any real audience, either keep the
> repo public or download the zips yourself and host them somewhere your
> users already are.

### Warn users about the first launch

The binaries are unsigned, so both operating systems will object once.

**Windows** — SmartScreen says "Windows protected your PC."
Click *More info* → *Run anyway*.

**macOS** — the app is blocked on open. Go to
**System Settings → Privacy & Security**, scroll down, and click *Open Anyway*
next to the message about your app. Only needed once.

Nobody discovers the macOS steps on their own. Put them in your own README,
or screenshot them.

### Which Mac download?

Apple menu → **About This Mac**. "Apple M1/M2/M3/M4" means Apple Silicon;
anything else is Intel.

---

## Troubleshooting

### The build passed but the app won't start

Almost always a **hidden import**. PyInstaller never runs your code — it reads
your source, follows `import` statements, and bundles what it finds. Anything
imported by name at *runtime* is invisible to it:

```python
matplotlib.use("TkAgg")   # "TkAgg" is a string, not an import
```

matplotlib turns that string into `matplotlib.backends.backend_tkagg` later, at
which point the module isn't in the bundle. The fix:

```toml
hidden_imports = ["matplotlib.backends.backend_tkagg"]
```

The signature of this problem is that **the build is green and the app fails on
a clean machine.** Suspect it any time a dependency uses plugins, drivers,
entry points, or `importlib.import_module`.

If your app is `windowed`, wrap your entrypoint so failures are visible rather
than silent:

```python
if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        show_some_dialog("Unexpected error", traceback.format_exc())
```

Without something like that, a `--windowed` app just disappears and you get bug
reports that say "it doesn't work."

### The app takes 10+ seconds to open

Normal for `onefile` with large dependencies — it unpacks to a temp directory on
every launch. Set `onefile = false` for near-instant startup, at the cost of
users seeing a folder instead of a single file.

### Build fails on `--add-data`

The separator differs by platform: `file.txt:.` on macOS, `file.txt;.` on
Windows. If you need data files on both, use a `.spec` file instead of
`extra_args`.

### The download is enormous

Check whether a dependency is pulling in something large you never use. Add it
to `exclude_modules` — but only if nothing imports it, or the app will die
immediately.

### Node.js deprecation warnings

Cosmetic. They mean an action declares an old Node runtime. Bump the `@v`
versions when convenient.

---

## Cost

Free and unlimited on public repos.

On private repos it draws from your monthly Actions minutes, and macOS is
metered at **10x** wall time (Windows 2x, Linux 1x). A three-platform build runs
roughly 110 charged minutes, so the 2,000-minute free tier covers about 18
releases a month. Free accounts have a $0 spending limit by default, so builds
stop rather than bill you.

Private repos also cap artifact storage at 500 MB, which fills quickly with
50 MB binaries — that's what `artifact_retention_days` is for.

---

## Code signing (optional)

Signing removes the security warnings. It costs money and is only worth it for
a wide or non-technical audience.

- **macOS** — $99/year Apple Developer account, then `codesign` and `notarytool`
  steps in the workflow with your certificate stored as a repo secret.
- **Windows** — an EV code-signing certificate, roughly $250–400/year, and since
  2023 it requires a hardware token or cloud HSM.

Everything here works unsigned. Your users just click through one warning.

---

## Limits worth knowing

- Builds only for the platforms GitHub offers runners for
- Intel Mac support ends when the macOS 15 runner retires in Fall 2027
- One binary per architecture; universal2 builds fail if any dependency lacks a
  universal wheel, which is why Intel and Apple Silicon are built separately
- Entries in `extra_args` must not contain spaces — they expand unquoted
