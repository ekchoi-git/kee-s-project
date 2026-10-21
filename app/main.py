"""대화창. Excel·AI 작업은 별도 작업 스레드 하나에서만 처리해 창이 멈추지 않게 한다."""
import queue, threading, traceback
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

import pythoncom

from .excel_ops import PlanError

EXAMPLES = ["회사별 연도별 금액 합계를 수식으로 만들어줘", "같은 내용을 피벗 테이블로 만들어줘",
            "금액 1억 넘는 것만 보이게 하고 큰 순서로 정렬해줘"]


class Worker(threading.Thread):
    """COM 개체는 만든 스레드에서만 써야 하므로 Excel 관련 일은 전부 여기서 순서대로 처리."""

    def __init__(self, results):
        super().__init__(daemon=True)
        self.jobs, self.results, self.session = queue.Queue(), results, None

    def submit(self, fn, done):
        self.jobs.put((fn, done))

    def run(self):
        pythoncom.CoInitialize()
        while True:
            fn, done = self.jobs.get()
            try:
                if self.session is None:
                    from .core import Session
                    self.session = Session()
                res, err = fn(self.session), None
            except PlanError as e:
                res, err = None, str(e)
            except Exception as e:
                traceback.print_exc()
                res, err = None, f"예상하지 못한 오류: {e}"
            self.results.put((done, res, err, bool(self.session and self.session.last)))


class App:
    def __init__(self, root):
        self.root, self.results = root, queue.Queue()
        self.worker = Worker(self.results)
        self.worker.start()
        root.title("Excel AI 도우미")
        root.geometry("640x720")

        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="작업 대상").pack(side="left")
        self.target = ttk.Combobox(top, state="readonly", width=48)
        self.target.pack(side="left", padx=6, fill="x", expand=True)
        self.target.bind("<<ComboboxSelected>>", lambda e: self.on_target())
        ttk.Button(top, text="새로고침", command=self.refresh).pack(side="left")

        self.log = ScrolledText(root, wrap="word", state="disabled", font=("맑은 고딕", 10), padx=8, pady=8)
        self.log.pack(fill="both", expand=True, padx=8)
        self.log.tag_config("me", foreground="#1a56db", font=("맑은 고딕", 10, "bold"))
        self.log.tag_config("err", foreground="#c81e1e")
        self.log.tag_config("info", foreground="#6b7280")

        ex = ttk.Frame(root, padding=(8, 4))
        ex.pack(fill="x")
        for t in EXAMPLES:
            ttk.Button(ex, text=t[:14] + "…", command=lambda t=t: self.set_input(t)).pack(side="left", padx=2)

        bottom = ttk.Frame(root, padding=8)
        bottom.pack(fill="x")
        self.entry = ttk.Entry(bottom, font=("맑은 고딕", 11))
        self.entry.pack(side="left", fill="x", expand=True)
        self.entry.bind("<Return>", lambda e: self.send())
        self.send_btn = ttk.Button(bottom, text="보내기", command=self.send)
        self.send_btn.pack(side="left", padx=4)
        self.undo_btn = ttk.Button(bottom, text="직전 AI 작업 되돌리기", command=self.undo, state="disabled")
        self.undo_btn.pack(side="left")

        self.status = ttk.Label(root, text="", padding=(8, 0, 8, 6), foreground="#6b7280")
        self.status.pack(fill="x")
        self.say("Excel에서 파일을 연 뒤, 작업 대상을 확인하고 하고 싶은 일을 말로 적어주세요.", "info")
        root.after(100, self.poll)
        self.refresh()

    # --- 화면 ---
    def say(self, text, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n\n", tag)
        self.log.configure(state="disabled")
        self.log.see("end")

    def set_input(self, text):
        self.entry.delete(0, "end")
        self.entry.insert(0, text)
        self.entry.focus()

    def busy(self, msg):
        self.status.configure(text=msg)
        for b in (self.send_btn, self.undo_btn):
            b.configure(state="disabled")

    def idle(self):
        self.status.configure(text="")
        self.send_btn.configure(state="normal")

    def poll(self):
        while not self.results.empty():
            done, res, err, can_undo = self.results.get()
            done(res, err)
            # 되돌릴 수 있는지는 화면이 추측하지 않고 실제 세션 상태를 따름
            self.undo_btn.configure(state="normal" if can_undo and self.worker.jobs.empty() else "disabled")
        self.root.after(100, self.poll)

    # --- 동작 ---
    def refresh(self):
        def got(res, err):
            if err:
                self.say(err, "err")
                return
            items, active = res
            self.items = items
            self.target["values"] = [f"{wb} › {ws}" for wb, ws in items]
            if items:
                self.target.current(items.index(active) if active in items else 0)
        self.worker.submit(lambda s: s.targets(), got)

    def on_target(self):
        # 대상이 바뀌면 이전 대화 맥락을 그대로 쓰지 않음 (PRD FR-02)
        self.worker.submit(lambda s: s.planner.reset(), lambda r, e: None)
        self.say(f"작업 대상: {self.target.get()} (이전 대화 맥락은 초기화)", "info")

    def send(self):
        text = self.entry.get().strip()
        if not text or str(self.send_btn["state"]) == "disabled":
            return
        if not getattr(self, "items", None) or self.target.current() < 0:
            self.say("작업 대상이 없습니다. Excel에서 파일을 열고 새로고침을 눌러주세요.", "err")
            self.refresh()
            return
        wb, ws = self.items[self.target.current()]  # 요청 시점의 대상으로 고정
        self.entry.delete(0, "end")
        self.say("나: " + text, "me")
        self.busy("AI가 요청을 해석하고 Excel에 적용하는 중…")

        def got(res, err):
            self.idle()
            if err:
                self.say("✖ " + err, "err")
                return
            msg, changed = res
            self.say(("✔ " if changed else "") + msg)

        self.worker.submit(lambda s: s.send(text, wb, ws), got)

    def undo(self):
        self.busy("되돌리는 중…")

        def got(res, err):
            self.idle()
            if err:
                self.say("✖ " + err, "err")
            else:
                self.say("↩ " + res, "info")

        self.worker.submit(lambda s: s.undo(), got)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
