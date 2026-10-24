#!/usr/bin/env python3
"""Publish a new version of 5belan so every installed copy (Windows / macOS / Linux) updates itself.

Usage (run from the app folder, the one that holds deepseek_chat.py):
    python publish.py --repo OWNER/REPO --bump patch --notes "what changed" --push

What it does:
  1. optionally bumps APP_VERSION in deepseek_chat.py
  2. writes update.json (version, notes, SHA-256 of every app file)
  3. with --push: git add / commit / push, so the files go live on GitHub
Installed apps read update.json from raw.githubusercontent.com and download only the files that changed.
"""
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(HERE, "deepseek_chat.py")
FILES = ["deepseek_chat.py", "guard.py", "5belan.png", "5belan.ico", "requirements.txt"]
GITIGNORE = """config.json
chats.json
*.log
*.log.old
*.bak*
*.broken
run_state.json
resume.json
.venv/
_last_good/
_last_try/
_update_backup/
_brand/
_brand_probe.txt
__pycache__/
*.pyc
update_source.txt
"""


def read(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        return f.read()


def write(path, text):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(text)


def get_version(src):
    m = re.search(r'^APP_VERSION = "([^"]+)"', src, re.M)
    if not m:
        sys.exit("ERROR: APP_VERSION not found in deepseek_chat.py")
    return m.group(1)


def bump(v, kind):
    p = [int(x) for x in (v.split(".") + ["0", "0", "0"])[:3]]
    if kind == "major":
        p = [p[0] + 1, 0, 0]
    elif kind == "minor":
        p = [p[0], p[1] + 1, 0]
    else:
        p[2] += 1
    return ".".join(map(str, p))


def sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def run(args):
    print("$ " + " ".join(args))
    return subprocess.call(args, cwd=HERE)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", help="GitHub OWNER/REPO (sets the default update address inside the app)")
    ap.add_argument("--branch", default="main")
    ap.add_argument("--bump", choices=["patch", "minor", "major"], help="raise the version number")
    ap.add_argument("--version", help="set an exact version")
    ap.add_argument("--notes", default="", help="short text shown to users before they update")
    ap.add_argument("--pip", nargs="*", default=[], help="extra Python packages the new version needs")
    ap.add_argument("--push", action="store_true", help="git add, commit and push")
    a = ap.parse_args()

    src = read(APP)
    if a.repo:
        url = "https://raw.githubusercontent.com/%s/%s/update.json" % (a.repo, a.branch)
        src = re.sub(r'^DEFAULT_UPDATE_URL = ".*?"', 'DEFAULT_UPDATE_URL = "%s"' % url, src, flags=re.M)
        print("Update address set to " + url)
    ver = get_version(src)
    if a.version:
        new = a.version
    elif a.bump:
        new = bump(ver, a.bump)
    else:
        new = ver
    if new != ver:
        src = re.sub(r'^APP_VERSION = "[^"]+"', 'APP_VERSION = "%s"' % new, src, flags=re.M)
    try:
        compile(src, "deepseek_chat.py", "exec")
    except SyntaxError as e:
        sys.exit("ERROR: deepseek_chat.py has a syntax error, not publishing: %s" % e)
    write(APP, src)

    files = []
    for n in FILES:
        p = os.path.join(HERE, n)
        if os.path.exists(p):
            files.append({"name": n, "sha256": sha(p)})
    manifest = {"version": new, "notes": a.notes, "files": files}
    if a.pip:
        manifest["pip"] = a.pip
    with open(os.path.join(HERE, "update.json"), "w", encoding="utf-8", newline="\n") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print("update.json written: version %s, %d files" % (new, len(files)))

    gi = os.path.join(HERE, ".gitignore")
    if not os.path.exists(gi):
        with open(gi, "w", newline="\n") as f:
            f.write(GITIGNORE)
    if a.push:
        if not os.path.isdir(os.path.join(HERE, ".git")):
            if run(["git", "init", "-b", a.branch]) != 0:
                sys.exit("ERROR: git is not installed. Install Git from https://git-scm.com and run again.")
            if a.repo:
                run(["git", "remote", "add", "origin", "https://github.com/%s.git" % a.repo])
        run(["git", "add", "-A"])
        run(["git", "commit", "-m", "Release %s" % new])
        if run(["git", "push", "-u", "origin", a.branch]) != 0:
            sys.exit("ERROR: push failed. Check the repository name and that you are signed in to GitHub (git credential manager).")
        print("Published version %s. Installed apps will offer the update within a few hours (or press 'check for updates')." % new)
    else:
        print("Nothing pushed (add --push to publish).")


if __name__ == "__main__":
    main()
