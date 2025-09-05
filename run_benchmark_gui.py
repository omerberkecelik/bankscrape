# run_benchmark_gui.py
# Minimal Tkinter GUI that runs your original run_benchmark.py and shows progress.

import sys, os, threading, queue, subprocess, re, time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

BANK_COUNT = 9  # number of banks your script processes

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bank Benchmark")
        self.geometry("840x520")

        self.start_btn = ttk.Button(self, text="Start", command=self.start)
        self.start_btn.pack(pady=8)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", maximum=BANK_COUNT)
        self.progress.pack(fill="x", padx=10)

        self.log = scrolledtext.ScrolledText(self, height=22, state="disabled")
        self.log.pack(fill="both", expand=True, padx=10, pady=8)

        self.status = ttk.Label(self, text="Ready.")
        self.status.pack(anchor="w", padx=10, pady=(0,8))

        self.q = queue.Queue()
        self.proc = None
        self.done_banks = 0

        self.after(100, self._drain_queue)

    def _append(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _reader(self, cmd):
        try:
            self.proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1
            )
            for line in self.proc.stdout:
                self.q.put(line)
            self.proc.wait()
            self.q.put(f"\n[EXIT] Process finished with code {self.proc.returncode}\n")
        except Exception as e:
            self.q.put(f"\n[ERROR] {e}\n")

    def start(self):
        if self.proc and self.proc.poll() is None:
            return
        self.done_banks = 0
        self.progress['value'] = 0
        self.log.configure(state="normal")
        self.log.delete("1.0","end")
        self.log.configure(state="disabled")
        self.status.config(text="Running...")
        self.start_btn.config(state="disabled")

        # run the original script unchanged
        cmd = [sys.executable, os.path.join(os.path.dirname(__file__), "run_benchmark.py")]
        threading.Thread(target=self._reader, args=(cmd,), daemon=True).start()

    def _drain_queue(self):
        try:
            while True:
                line = self.q.get_nowait()
                self._append(line)

                # advance on markers your script already prints
                if line.startswith("[OK] Filled column"):
                    self.done_banks = min(BANK_COUNT, self.done_banks + 1)
                    self.progress['value'] = self.done_banks
                if line.startswith("[DONE]"):
                    self.status.config(text="Done. See Benchmark_Results.xlsx")
                    self.start_btn.config(state="normal")
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

if __name__ == "__main__":
    app = App()
    app.mainloop()

