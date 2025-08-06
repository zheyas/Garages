import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import back


class GarageApp:
    COLUMNS = (
        back.RENT_GARAGE_COL,
        back.RENT_DATE_COL,
        back.RENT_SUM_COL,
        back.STATUS_COL,
    )
    TAG_COLORS = {
        back.STATUS_PAID:    "#d4edda",
        back.STATUS_OVERDUE: "#f8d7da",
        back.STATUS_PENDING: "#fff3cd",
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Мониторинг оплат аренды гаражей")
        self.root.geometry("900x600")

        # Верхняя панель
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", pady=10)
        tk.Button(
            btn_frame, text="Загрузить аренду",
            command=self.load_rent
        ).pack(side="left", padx=5)
        tk.Button(
            btn_frame, text="Загрузить выписку",
            command=self.load_statement
        ).pack(side="left", padx=5)
        tk.Button(
            btn_frame, text="Проверить оплаты",
            command=self.check_payments
        ).pack(side="left", padx=5)

        # Таблица
        tree_frame = tk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree = ttk.Treeview(
            tree_frame, columns=self.COLUMNS, show="headings"
        )
        for col in self.COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center",
                             width=100 if col != back.RENT_SUM_COL else 120,
                             stretch=True)
        vsb = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # Цвета по статусам
        for status, color in self.TAG_COLORS.items():
            self.tree.tag_configure(status, background=color)

        # Хранилища
        self.rent_df = None
        self.statement_df = None
        self.result_df = None

    def _load_file(self, loader, attr_name, description):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        try:
            df = loader(path)
            setattr(self, attr_name, df)
            messagebox.showinfo("Успех", f"{description} загружен.")
        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось загрузить {description.lower()}:\n{e}"
            )

    def load_rent(self):
        self._load_file(
            back.load_rent_excel,
            "rent_df",
            "файл аренды"
        )

    def load_statement(self):
        self._load_file(
            back.load_statement,
            "statement_df",
            "файл выписки"
        )

    def check_payments(self):
        if self.rent_df is None or self.statement_df is None:
            messagebox.showwarning(
                "Внимание",
                "Загрузите оба файла прежде чем проверять оплаты."
            )
            return
        try:
            self.result_df = back.analyze_payments(
                self.rent_df, self.statement_df
            )
        except Exception as e:
            messagebox.showerror("Ошибка при анализе", str(e))
            return

        # Сохранение отчёта "Сохранить как..."
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="result.xlsx"
        )
        if save_path:
            try:
                back.save_to_excel(self.result_df, save_path)
                messagebox.showinfo(
                    "Готово", f"Отчёт сохранён:\n{save_path}"
                )
            except Exception as e:
                messagebox.showerror("Ошибка при сохранении", str(e))
                return
        else:
            messagebox.showinfo("Отмена", "Сохранение отчёта отменено.")

        # Обновляем таблицу
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for _, row in self.result_df.iterrows():
            vals = (
                row[back.RENT_GARAGE_COL],
                row[back.RENT_DATE_COL].strftime("%d.%m.%Y"),
                f"{row[back.RENT_SUM_COL]:,.2f} ₽",
                row[back.STATUS_COL],
            )
            self.tree.insert(
                "", "end", values=vals, tags=(row[back.STATUS_COL],)
            )

        # График
        self.show_chart_window()

    def show_chart_window(self):
        if self.result_df is None:
            return
        win = tk.Toplevel(self.root)
        win.title("График: количество гаражей по статусам")
        win.geometry("500x400")

        fig, ax = plt.subplots(figsize=(5, 4), tight_layout=True)
        counts = self.result_df[back.STATUS_COL].value_counts()
        statuses = counts.index.tolist()
        values = counts.values.tolist()
        color_map = {
            back.STATUS_PAID:    "#28a745",
            back.STATUS_OVERDUE: "#dc3545",
            back.STATUS_PENDING: "#ffc107",
        }
        bar_colors = [color_map.get(s, "#6c757d") for s in statuses]
        x = list(range(len(statuses)))
        ax.bar(x, values, color=bar_colors)
        ax.set_title("Число гаражей по статусам")
        ax.set_ylabel("Гаражей")
        ax.set_xticks(x)
        ax.set_xticklabels(statuses, rotation=45, ha="right")

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
