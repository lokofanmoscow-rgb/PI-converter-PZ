"""ПИ → ПЗ конвертер: окно программы (tkinter)."""
import os
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, scrolledtext

import converter

APP_TITLE = "ПИ → ПЗ конвертер"


def resource_path(name: str) -> str:
    """Путь к файлу рядом с EXE/скриптом, иначе — встроенному в EXE (PyInstaller)."""
    exe_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
    local = os.path.join(exe_dir, name)
    if os.path.exists(local):
        return local
    bundle_dir = getattr(sys, "_MEIPASS", exe_dir)
    return os.path.join(bundle_dir, name)


def open_folder(path: str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606 — открыть папку в Проводнике
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class App:
    def __init__(self, root: tk.Tk, initial_path: str = ""):
        self.root = root
        self.pi_path = tk.StringVar()
        self.status = tk.StringVar(value="Выберите файл ПИ (.docx)")
        self.result_path = None

        root.title(APP_TITLE)
        root.geometry("760x460")
        root.minsize(560, 360)

        frame = tk.Frame(root, padx=12, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text=APP_TITLE, font=("Segoe UI", 14, "bold")).pack(anchor="w")

        row = tk.Frame(frame)
        row.pack(fill=tk.X, pady=(10, 4))
        tk.Button(row, text="Выбрать ПИ (.docx)", width=20, command=self.choose).pack(side=tk.LEFT)
        tk.Entry(row, textvariable=self.pi_path, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        row2 = tk.Frame(frame)
        row2.pack(fill=tk.X, pady=4)
        self.btn_create = tk.Button(row2, text="Создать ПЗ", width=20, command=self.create, state=tk.DISABLED)
        self.btn_create.pack(side=tk.LEFT)
        self.btn_open = tk.Button(row2, text="Открыть папку с результатом", command=self.open_result)
        self.lbl_status = tk.Label(row2, textvariable=self.status, anchor="w", justify=tk.LEFT)
        self.lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        self.log = scrolledtext.ScrolledText(frame, height=14, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        if initial_path:
            self.set_file(initial_path)

    # --- журнал ------------------------------------------------------------
    def write_log(self, text: str) -> None:
        def append():
            self.log.configure(state=tk.NORMAL)
            self.log.insert(tk.END, text + "\n")
            self.log.see(tk.END)
            self.log.configure(state=tk.DISABLED)

        self.root.after(0, append)

    def clear_log(self) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.configure(state=tk.DISABLED)

    # --- действия ----------------------------------------------------------
    def choose(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите файл ПИ",
            filetypes=[("Документ Word", "*.docx"), ("Все файлы", "*.*")],
        )
        if path:
            self.set_file(path)

    def set_file(self, path: str) -> None:
        self.pi_path.set(os.path.normpath(path))
        self.btn_create.configure(state=tk.NORMAL)
        self.btn_open.pack_forget()
        self.status.set("Файл выбран. Нажмите «Создать ПЗ».")
        self.clear_log()

    def create(self) -> None:
        path = self.pi_path.get()
        if not path:
            return
        self.btn_create.configure(state=tk.DISABLED)
        self.btn_open.pack_forget()
        self.status.set("Создаю ПЗ…")
        self.clear_log()
        threading.Thread(target=self._worker, args=(path,), daemon=True).start()

    def _worker(self, path: str) -> None:
        try:
            template = resource_path(converter.TEMPLATE_NAME)
            result = converter.convert(path, template, log=self.write_log)
            self.root.after(0, self._done, result, None)
        except converter.ConversionError as e:
            self.root.after(0, self._done, None, str(e))
        except PermissionError as e:
            self.root.after(0, self._done, None,
                            f"Нет доступа к файлу (возможно, ПЗ уже открыт в Word — закройте его):\n{e.filename}")
        except Exception as e:  # noqa: BLE001 — показать любую ошибку пользователю
            self.write_log(traceback.format_exc())
            self.root.after(0, self._done, None, f"Непредвиденная ошибка: {e}")

    def _done(self, result, error) -> None:
        self.btn_create.configure(state=tk.NORMAL)
        if error:
            self.status.set("Ошибка — ПЗ не создан.")
            self.write_log("ОШИБКА: " + error)
            messagebox.showerror(APP_TITLE, error)
            return
        self.result_path = result
        self.status.set(f"Готово: {os.path.basename(result)}")
        self.btn_open.pack(side=tk.LEFT, padx=(8, 0), before=self.lbl_status)

    def open_result(self) -> None:
        if self.result_path:
            open_folder(os.path.dirname(self.result_path))


def batch(args) -> int:
    """PI_to_PZ.exe --batch ПИ.docx [--out ПАПКА] — без окна (для проверки сборки)."""
    out_dir = None
    if "--out" in args:
        k = args.index("--out")
        out_dir = args[k + 1]
        args = args[:k] + args[k + 2:]
    log_lines = []
    try:
        converter.convert(args[0], resource_path(converter.TEMPLATE_NAME), out_dir, log_lines.append)
        code = 0
    except Exception:  # noqa: BLE001
        log_lines.append(traceback.format_exc())
        code = 1
    if out_dir:
        with open(os.path.join(out_dir, "PI_to_PZ.log"), "w", encoding="utf-8") as f:
            f.write("\n".join(log_lines))
    return code


def main() -> None:
    args = sys.argv[1:]
    if args and args[0] == "--batch":
        sys.exit(batch(args[1:]))
    root = tk.Tk()
    App(root, args[0] if args else "")  # путь, если ПИ перетащили на значок программы
    root.mainloop()


if __name__ == "__main__":
    main()
