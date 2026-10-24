# -*- coding: utf-8 -*-
"""تشغيل آمن لتطبيق deepseek_chat.py  (guard.py)

ليه هذا الملف موجود؟
  لو تعدّلنا على الملف، وصلحنا شي وجاء خطأ، البرنامج ما يفتح ويقعد عالق.
  guard.py يحل هالمشكلة بالترتيب:

  1. يتأكد أن الملف سليم صياغة قبل أي تشغيل.
  2. يحفظ نسخة من آخر تشغيل ناجح في مجلد _last_good.
  3. لو الملف فيه خطأ صياغة أو البرنامج طار بخطأ:
       - يحفظ الملف الخربان باسم deepseek_chat.py.broken
       - يرجّع نسخة _last_good ويشغّلها تلقائياً بدون ما تتدخل.
  4. لو ما فيه نسخة _last_good بعد، ياخذ أحدث نسخة احتياطية سليمة (.bak / .bak2 ...)
     ويسويها أساس - عشان ما يصير أن الملف خربان وما فيه شي نرجّع له.
  5. يكتشف الانهيار حتى لو البرنامج رجّع كود خروج صفر:
     يقارن حجم error.log قبل وبعد التشغيل.

يشتغل تلقائياً من run.bat. ما يحتاج تشغّله بنفسك.
"""

import glob
import os
import re
import shutil
import struct
import subprocess
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
GOOD_DIR = os.path.join(BASE, "_last_good")
TRY_DIR = os.path.join(BASE, "_last_try")
ERROR_LOG = os.path.join(BASE, "error.log")
MIN_UPTIME = 10.0   # أقل مدة نعتبرها "تشغيل ناجح"
RUNTIME_MIN = 20.0  # لو اشتغل أكثر من كذا ثم طاح، نعتبره كراش وقت تشغيل مو خطأ في الملف
MAX_RESTARTS = 3
MAX_FIXES = 2       # أقصى عدد مرات رجوع تلقائي قبل ما نوقف


def say(msg):
    """يطبع بدون ما يطيح لو الكونسول ما يعرض العربي."""
    try:
        print(msg, flush=True)
    except Exception:
        try:
            sys.stdout.write(str(msg).encode("ascii", "replace").decode("ascii") + "\n")
            sys.stdout.flush()
        except Exception:
            pass


# ------------------------------------------------------------------ قراءة الفحص
def read_source(path):
    """يقرأ الملف كـ نص، ويتجاوز BOM لو موجود (هذا اللي كان يكسر الفحص قبل)."""
    with open(path, "rb") as f:
        data = f.read()
    for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
        try:
            return data.decode(enc)
        except Exception:
            continue
    return data.decode("utf-8", errors="replace")


def syntax_error(path):
    """يرجّع نص الخطأ لو فيه مشكلة صياغة، وإلا None."""
    try:
        compile(read_source(path), path, "exec")
    except SyntaxError as e:
        return "SyntaxError في السطر %s: %s" % (e.lineno, e.msg)
    except Exception as e:
        return "%s: %s" % (type(e).__name__, e)
    return None


# ------------------------------------------------------------------ النسخ
def good_path(name):
    return os.path.join(GOOD_DIR, name)


def has_good(name):
    return os.path.exists(good_path(name))


def copy_file(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copyfile(src, dst)


def save_good(path, name):
    copy_file(path, good_path(name))


def same_file(a, b):
    try:
        if not os.path.exists(a) or not os.path.exists(b):
            return False
        with open(a, "rb") as x, open(b, "rb") as y:
            return x.read() == y.read()
    except Exception:
        return False


def backup_rank(p):
    """النسخ الأحدث أولاً: الرقم في الاسم، وإذا ما فيه رقم نعتبره الأول."""
    tail = p.rsplit(".bak", 1)[-1]
    n = int(tail) if tail.isdigit() else 1
    try:
        return (n, os.path.getmtime(p))
    except OSError:
        return (n, 0)


def backup_candidates(name):
    found = []
    for pat in (name + ".bak", name + ".bak*"):
        for p in glob.glob(os.path.join(BASE, pat)):
            if os.path.isfile(p) and p not in found:
                found.append(p)
    found.sort(key=backup_rank, reverse=True)
    return found


def bootstrap_good(name, app):
    """يتأكد أن فيه نسخة سليمة محفوظة، ويسوّي واحدة من الملف أو من نسخة احتياطية."""
    if has_good(name):
        return "kept"
    if not syntax_error(app):
        save_good(app, name)
        return "current"
    for p in backup_candidates(name):
        if not syntax_error(p):
            save_good(p, name)
            return os.path.basename(p)
    return "none"


def keep_broken(app, name):
    try:
        shutil.copyfile(app, os.path.join(BASE, name + ".broken"))
        return True
    except Exception:
        return False


def restore(app, name):
    """يرجّع النسخة السليمة، ويحفظ الملف الحالي باسم .broken قبل ما يستبدله."""
    src = good_path(name)
    if not os.path.exists(src):
        return False
    keep_broken(app, name)
    shutil.copyfile(src, app)
    return True


# ------------------------------------------------------------------ سجل الأخطاء
def log_size():
    try:
        return os.path.getsize(ERROR_LOG)
    except OSError:
        return 0


def tail_log(lines=12):
    if not os.path.exists(ERROR_LOG):
        return "(ما فيه error.log)"
    try:
        rows = read_source(ERROR_LOG).strip().splitlines()
    except Exception:
        return "(تعذّرت قراءة error.log)"
    return "\n".join(rows[-lines:]) if rows else "(error.log فاضي)"


# ------------------------------------------------------------------ الاسم في مدير المهام
# مدير المهام (ويندوز) يعرض "وصف الملف التنفيذي"، ومراقب النظام (لينكس/ماك) يعرض اسم الملف.
# فنسوي نسخة من بايثون باسم 5belan، وفيها بيانات الإصدار "خبيلان". أي فشل = نكمّل بالعادي.
BRAND_NAME = "\u062e\u0628\u064a\u0644\u0627\u0646"  # خبيلان
BRAND_EXE = "5belan"
BRAND_STAMP = "v1"


def _app_version():
    try:
        with open(os.path.join(BASE, "deepseek_chat.py"), encoding="utf-8-sig") as f:
            m = re.search(r'^APP_VERSION\s*=\s*"([\d.]+)"', f.read(), re.M)
        return m.group(1) if m else "0.0.0"
    except Exception:
        return "0.0.0"


def _ver_block(key, value=b"", children=b"", text=False):
    """كتلة واحدة من مورد VS_VERSIONINFO."""
    k = (key + "\0").encode("utf-16-le")
    body = k
    body += b"\0" * (-(6 + len(body)) % 4)
    body += value
    if children:
        body += b"\0" * (-(6 + len(body)) % 4)
        body += children
    vlen = len(value) // 2 if text else len(value)
    return struct.pack("<HHH", 6 + len(body), vlen, 1 if text else 0) + body


def _ver_pad(b):
    return b + b"\0" * (-len(b) % 4)


def version_resource(version):
    nums = [int(x) for x in re.findall(r"\d+", version)][:4]
    nums += [0] * (4 - len(nums))
    fixed = struct.pack("<13I", 0xFEEF04BD, 0x00010000, (nums[0] << 16) | nums[1], (nums[2] << 16) | nums[3],
                        (nums[0] << 16) | nums[1], (nums[2] << 16) | nums[3], 0x3F, 0, 0x00040004, 1, 0, 0, 0)
    fields = [("CompanyName", BRAND_NAME), ("FileDescription", BRAND_NAME), ("FileVersion", version),
              ("InternalName", BRAND_EXE), ("OriginalFilename", BRAND_EXE + ".exe"),
              ("ProductName", BRAND_NAME), ("ProductVersion", version)]
    strings = b"".join(_ver_pad(_ver_block(k, (v + "\0").encode("utf-16-le"), text=True)) for k, v in fields)
    table = _ver_block("040904B0", children=strings)
    sfi = _ver_block("StringFileInfo", children=_ver_pad(table))
    var = _ver_block("Translation", struct.pack("<HH", 0x0409, 0x04B0))
    vfi = _ver_block("VarFileInfo", children=_ver_pad(var))
    return _ver_block("VS_VERSION_INFO", fixed, _ver_pad(sfi) + _ver_pad(vfi))


def _set_version_info(path, data):
    import ctypes
    k = ctypes.WinDLL("kernel32", use_last_error=True)
    k.BeginUpdateResourceW.restype = ctypes.c_void_p
    k.BeginUpdateResourceW.argtypes = [ctypes.c_wchar_p, ctypes.c_int]
    k.UpdateResourceW.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ushort,
                                  ctypes.c_void_p, ctypes.c_ulong]
    k.EndUpdateResourceW.argtypes = [ctypes.c_void_p, ctypes.c_int]
    h = k.BeginUpdateResourceW(path, 0)
    if not h:
        return False
    buf = ctypes.create_string_buffer(data, len(data))
    ok = k.UpdateResourceW(h, 16, 1, 0x0409, buf, len(data))  # RT_VERSION, id 1, en-US
    return bool(k.EndUpdateResourceW(h, 0 if ok else 1)) and bool(ok)


def _venv_home():
    cfg = os.path.join(sys.prefix, "pyvenv.cfg")
    try:
        with open(cfg, encoding="utf-8-sig") as f:
            for line in f:
                if line.split("=")[0].strip().lower() == "home":
                    return line.split("=", 1)[1].strip()
    except OSError:
        pass
    return ""


def _branded_windows():
    if sys.prefix == getattr(sys, "base_prefix", sys.prefix):
        return ""
    scripts = os.path.join(sys.prefix, "Scripts")
    dest = os.path.join(scripts, BRAND_EXE + ".exe")
    stamp_file = os.path.join(scripts, BRAND_EXE + ".stamp")
    home = _venv_home()
    stamp = "%s|%s|%s" % (BRAND_STAMP, home, sys.version)
    try:
        with open(stamp_file, encoding="utf-8") as f:
            if f.read() == stamp and os.path.exists(dest):
                return dest
    except OSError:
        pass
    src = os.path.join(home, "pythonw.exe")
    if not home or not os.path.exists(src):
        return ""
    try:
        for n in os.listdir(home):
            low = n.lower()
            if low in ("python3.dll", "python%d%d.dll" % sys.version_info[:2]) or low.startswith("vcruntime140"):
                shutil.copyfile(os.path.join(home, n), os.path.join(scripts, n))
        tmp = dest + ".new"
        shutil.copyfile(src, tmp)
        if not _set_version_info(tmp, version_resource(_app_version())):
            os.remove(tmp)
            return ""
        os.replace(tmp, dest)
        # تجربة حقيقية: لازم يشتغل ويشوف نفس البيئة
        probe = os.path.join(BASE, "_brand_probe.txt")
        try:
            os.remove(probe)
        except OSError:
            pass
        subprocess.run([dest, "-c", "import sys;open(r'%s','w').write(sys.prefix)" % probe],
                       timeout=30, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        with open(probe, encoding="utf-8") as f:
            good = os.path.normcase(f.read().strip()) == os.path.normcase(sys.prefix)
        os.remove(probe)
        if not good:
            os.remove(dest)
            return ""
        with open(stamp_file, "w", encoding="utf-8") as f:
            f.write(stamp)
        return dest
    except Exception as e:
        say("(الاسم الجديد ما اشتغل: %s)" % e)
        try:
            if os.path.exists(dest):
                os.remove(dest)
        except OSError:
            pass
        return ""


def _branded_posix():
    if sys.prefix == getattr(sys, "base_prefix", sys.prefix):
        return ""
    link = os.path.join(os.path.dirname(sys.executable), BRAND_EXE)
    try:
        if not os.path.lexists(link):
            os.symlink(os.path.basename(sys.executable), link)
        return link if os.access(link, os.X_OK) else ""
    except OSError:
        return ""


def branded_python():
    """مسار بايثون باسم 5belan، أو نص فاضي لو ما نقدر."""
    try:
        return _branded_windows() if os.name == "nt" else _branded_posix()
    except Exception:
        return ""


def relaunch_branded():
    """يعيد تشغيل الحارس بالاسم الجديد. يرجّع True لو سلّم الشغل لنسخة ثانية."""
    if os.environ.get("BELAN_BRANDED"):
        return False
    if os.name == "nt" and not os.path.basename(sys.executable).lower().startswith("pythonw"):
        return False  # شغّله بنفسك من كونسول: لا تفصله عنه
    exe = branded_python()
    if not exe or os.path.normcase(exe) == os.path.normcase(sys.executable):
        return False
    env = dict(os.environ)
    env["BELAN_BRANDED"] = "1"
    args = [exe, os.path.abspath(__file__)] + sys.argv[1:]
    try:
        if os.name == "nt":
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
            subprocess.Popen(args, cwd=BASE, env=env, creationflags=flags, close_fds=True)
            return True
        os.execve(exe, args, env)
    except Exception:
        return False
    return False


# ------------------------------------------------------------------ التشغيل
def run_once(app):
    return subprocess.call([sys.executable, app], cwd=BASE)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--brand":  # يستدعيه المثبّت
        exe = branded_python()
        print("branded: " + (exe or "no"))
        return 0
    if relaunch_branded():
        return 0
    name = sys.argv[1] if len(sys.argv) > 1 else "deepseek_chat.py"
    app = os.path.join(BASE, name)

    say("=" * 60)
    say("  Safe launcher - guard.py")
    say("  لو صار خطأ، آخر نسخة شغّالة ترجع تلقائياً")
    say("=" * 60)

    if not os.path.exists(app):
        say("!! ما لقيت الملف: " + app)
        return 1

    src = bootstrap_good(name, app)
    if src == "kept":
        say(">> عندنا نسخة شغّالة محفوظة.")
    elif src == "current":
        say(">> ما كان فيه نسخة محفوظة، فحفظت الملف الحالي كأساس شغّال.")
    elif src == "none":
        say(">> ما فيه نسخة محفوظة ولا نسخة احتياطية سليمة. بنشغّل ونشوف.")
    else:
        say(">> ما كان فيه نسخة محفوظة، فأخذت %s كأساس شغّال." % src)

    fixes = 0
    restarts = 0
    try:
        while True:
            err = syntax_error(app)
            if err:
                say("")
                say("!! الملف فيه مشكلة: " + err)
                if fixes < MAX_FIXES and restore(app, name):
                    fixes += 1
                    say(">> رجّعت النسخة الشغّالة، وملفك محفوظ في %s.broken" % name)
                    continue
                say(">> ما فيه نسخة أرجّع لها. وقفت هنا.")
                return 1

            try:
                copy_file(app, os.path.join(TRY_DIR, name))
            except Exception:
                pass

            before = log_size()
            t0 = time.time()
            say("")
            say(">> اشغّل البرنامج...  سكّر نافذة البرنامج عشان توقف.")
            code = run_once(app)
            uptime = time.time() - t0
            native = code < 0 or code >= 0xC0000000  # a crash inside Qt/Windows (access violation...), never a problem of the file itself

            if code != 0 and (uptime >= RUNTIME_MIN or native):
                # it ran fine for a while and then died: a runtime crash, not a broken file.
                # Never swap the file for an old copy here; just restart so the app can offer to resume.
                say("")
                say("!! The app crashed while running (exit code %s) after %.0f s. A native crash is not a problem of the file: not restoring any old copy." % (code, uptime))
                try:
                    with open(os.path.join(BASE, "events.log"), "a", encoding="utf-8") as lf:
                        lf.write("%s  guard: app exited with code %s after %.0f s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), code, uptime))
                except Exception:
                    pass
                if native and restarts < MAX_RESTARTS:
                    restarts += 1
                    say(">> Restarting automatically (%d of %d)..." % (restarts, MAX_RESTARTS))
                    time.sleep(2)
                    continue
                return code
            if code != 0:
                say("")
                say("!! البرنامج طار بخطأ، كود الخروج %s، بعد %.1f ثانية." % (code, uptime))
            elif log_size() > before:
                say("")
                say("!! البرنامج كتب خطأ في error.log بعد %.1f ثانية." % uptime)
            else:
                if uptime >= MIN_UPTIME:
                    save_good(app, name)
                    say(">> التشغيل نجح بعد %.0f ثانية. حدّثت النسخة الشغّالة." % uptime)
                else:
                    say(">> البرنامج اشتغل %.0f ثانية وسطّر عادي." % uptime)
                return 0

            say("---- آخر سطور error.log ----")
            say(tail_log())

            if same_file(app, good_path(name)):
                say(">> الملف الحالي هو نفسه النسخة الشغّالة، يعني المشكلة مو من تعديل.")
                say(">> شف error.log أو الحزم المثبّتة.")
                return code or 1

            if fixes >= MAX_FIXES:
                say(">> رجّعت النسخة %d مرة وما نفع. وقفت." % fixes)
                say(">> راجع الملف: %s.broken" % name)
                return code or 1

            if not restore(app, name):
                say(">> ما لقيت نسخة أرجّع لها.")
                return code or 1

            fixes += 1
            say("")
            say(">> رجّعت آخر نسخة شغّالة. ملفك محفوظ في %s.broken" % name)
            say(">> أعيد التشغيل تلقائياً، محاولة %d من %d..." % (fixes, MAX_FIXES))
            time.sleep(1)
    except KeyboardInterrupt:
        say(">> أوقفتها بنفسك.")
        return 130
    except Exception as e:
        say("!! صار خطأ في الحماية نفسها: %s" % e)
        say(">> بنشغّل البرنامج مباشرة عشان ما تعلق.")
        return run_once(app)


if __name__ == "__main__":
    sys.exit(main())
