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
import shutil
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


# ------------------------------------------------------------------ التشغيل
def run_once(app):
    return subprocess.call([sys.executable, app], cwd=BASE)


def main():
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
