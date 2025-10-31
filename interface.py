"""Простой графический интерфейс для запуска process_103."""
from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from process_103 import ProcessResult, process_excel


class ProcessApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Обработка Excel в формат 103")
        self.resizable(False, False)

        self._input_var = tk.StringVar()
        self._output_var = tk.StringVar()
        self._template_var = tk.StringVar()
        self._inn_var = tk.StringVar()
        self._region_var = tk.StringVar(value="Челябинская область")
        self._status_var = tk.StringVar(value="Готово к обработке")

        self._build_layout()

    def _build_layout(self) -> None:
        padding = {"padx": 8, "pady": 4, "sticky": "w"}

        frame = tk.Frame(self)
        frame.grid(column=0, row=0, padx=12, pady=12)

        tk.Label(frame, text="Исходный файл Excel:").grid(column=0, row=0, **padding)
        tk.Entry(frame, textvariable=self._input_var, width=45).grid(
            column=1, row=0, **padding
        )
        tk.Button(frame, text="Обзор", command=self._browse_input).grid(
            column=2, row=0, padx=4, pady=4
        )

        tk.Label(frame, text="Файл шаблона (опционально):").grid(column=0, row=1, **padding)
        tk.Entry(frame, textvariable=self._template_var, width=45).grid(
            column=1, row=1, **padding
        )
        tk.Button(frame, text="Обзор", command=self._browse_template).grid(
            column=2, row=1, padx=4, pady=4
        )

        tk.Label(frame, text="Выходной файл Excel:").grid(column=0, row=2, **padding)
        tk.Entry(frame, textvariable=self._output_var, width=45).grid(
            column=1, row=2, **padding
        )
        tk.Button(frame, text="Сохранить как", command=self._browse_output).grid(
            column=2, row=2, padx=4, pady=4
        )

        tk.Label(frame, text="ИНН отправителя:").grid(column=0, row=3, **padding)
        tk.Entry(frame, textvariable=self._inn_var, width=20).grid(column=1, row=3, **padding)

        tk.Label(frame, text="Регион (oblo):").grid(column=0, row=4, **padding)
        tk.Entry(frame, textvariable=self._region_var, width=45).grid(
            column=1, row=4, **padding
        )

        self._run_button = tk.Button(frame, text="Обработать", command=self._start_processing)
        self._run_button.grid(column=0, row=5, columnspan=3, pady=(12, 4))

        tk.Label(frame, textvariable=self._status_var, fg="gray").grid(
            column=0, row=6, columnspan=3, sticky="w", padx=8
        )

    def _browse_input(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Выберите исходный Excel",
            filetypes=(("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")),
        )
        if file_path:
            self._input_var.set(file_path)
            default_output = Path(file_path).with_name("103_result.xlsx")
            if not self._output_var.get():
                self._output_var.set(str(default_output))

    def _browse_template(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Выберите шаблон Таблица1",
            filetypes=(("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")),
        )
        if file_path:
            self._template_var.set(file_path)

    def _browse_output(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Сохранить результат", defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"),)
        )
        if file_path:
            self._output_var.set(file_path)

    def _start_processing(self) -> None:
        if not self._input_var.get():
            messagebox.showwarning("Не выбран файл", "Укажите путь к исходному файлу Excel")
            return
        if not self._output_var.get():
            messagebox.showwarning("Не выбран файл", "Укажите путь к выходному файлу Excel")
            return

        self._run_button.configure(state=tk.DISABLED)
        self._status_var.set("Идет обработка...")
        threading.Thread(target=self._process_in_thread, daemon=True).start()

    def _process_in_thread(self) -> None:
        try:
            result = process_excel(
                self._input_var.get(),
                self._output_var.get(),
                inn=self._inn_var.get() or None,
                template_path=self._template_var.get() or None,
                region=self._region_var.get() or "Челябинская область",
            )
        except Exception as exc:
            self._show_error(str(exc))
        else:
            self._show_success(result)
        finally:
            self._run_button.configure(state=tk.NORMAL)

    def _show_success(self, result: ProcessResult) -> None:
        messagebox.showinfo(
            "Готово",
            f"Создан файл {result.output_path}\n"
            f"Всего записей: {result.total_rows}\n"
            f"tipo=2: {result.tipo2_rows}",
        )
        self._status_var.set("Обработка завершена успешно")

    def _show_error(self, message: str) -> None:
        messagebox.showerror("Ошибка", message)
        self._status_var.set("Ошибка: " + message)


def main() -> None:
    app = ProcessApp()
    app.mainloop()


if __name__ == "__main__":
    main()
