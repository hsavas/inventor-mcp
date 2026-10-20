"""Tek-instance testi: iki surec x iki thread ayni anda get_app() cagirir.

Kullanim:  .venv\\Scripts\\python.exe tests\\check_single_instance.py
Basarili sonuc: Inventor surec sayisi en fazla 1 artar ve tum cagrilar ayni
Inventor PID'ine baglanir.
"""

import os
import subprocess
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def worker_process() -> None:
    from inventor.connection import get_app
    import win32process

    pids = []

    def call():
        app = get_app()
        hwnd = app.MainFrameHWND
        pids.append(win32process.GetWindowThreadProcessId(hwnd)[1])

    ts = [threading.Thread(target=call) for _ in range(2)]
    [t.start() for t in ts]
    [t.join() for t in ts]
    print("PIDS", *pids, flush=True)
    input()  # ust surec bitene kadar ROT kaydini canli tut


def inventor_pids() -> set:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq Inventor.exe", "/FO", "CSV", "/NH"],
        capture_output=True, text=True,
    ).stdout
    return {int(l.split('","')[1]) for l in out.splitlines() if l.startswith('"Inventor')}


if __name__ == "__main__" and "--worker" in sys.argv:
    worker_process()
elif __name__ == "__main__":
    before = inventor_pids()
    procs = [
        subprocess.Popen([sys.executable, __file__, "--worker"], stdin=subprocess.PIPE,
                         stdout=subprocess.PIPE, text=True)
        for _ in range(2)
    ]
    lines = [p.stdout.readline().strip() for p in procs]
    after = inventor_pids()
    for p in procs:
        p.communicate("\n")
    seen = {int(x) for l in lines for x in l.split()[1:]}
    print("worker output:", lines)
    print("new Inventor processes:", sorted(after - before))
    print("connected PIDs:", sorted(seen))
    ok = len(after - before) <= 1 and len(seen) == 1
    print("RESULT:", "PASS" if ok else "FAIL")
