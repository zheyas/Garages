import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

import back


class GarageApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Мониторинг оплат аренды гаражей")
        self.root.geometry("900x600")

        # Верхняя панель с кнопками
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", pady=10)

        tk.Button(
            btn_frame, text="Загрузить Arenda.xlsx",
            command=self.load_arenda)\
            .pack(side="left", padx=5)
        tk.Button(
            btn_frame, text="Загрузить выписку",
            command=self.load_vypiska)\
            .pack(side="left", padx=5)
        tk.Button(
            btn_frame, text="Проверить оплаты",
            command=self.check_payments)\
            .pack(side="left", padx=5)

        # Фрейм для таблицы
        tree_frame = tk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("Номер гаража", "Дата оплаты", "Сумма оплаты", "Статус")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(
                col, anchor="center",
                width=100 if col != "Сумма оплаты" else 120,
                stretch=True
            )
        # Вертикальный скроллбар
        vsb = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # Теги для раскраски строк
        self.tree.tag_configure("Получен", background="#d4edda")
        self.tree.tag_configure("Просрочен", background="#f8d7da")
        self.tree.tag_configure("Срок не наступил", background="#fff3cd")

        # Храним загруженные DataFrame и результат
        self.arenda_df = None
        self.vypiska_df = None
        self.result_df = None

    def load_arenda(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        try:
            df = back.load_excel(path)
            df = back.normalize_arenda_columns(df)
            self.arenda_df = df
            messagebox.showinfo("Успех", "Файл аренды загружен.")
        except Exception as e:
            messagebox.showerror(
                "Ошибка", f"Не удалось загрузить файл аренды:\n{e}"
            )

    def load_vypiska(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        try:
            self.vypiska_df = back.load_vypiska_operations(path)
            messagebox.showinfo("Успех", "Файл выписки загружен.")
        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось загрузить файл выписки:\n{e}")

    def check_payments(self):
        if self.arenda_df is None or self.vypiska_df is None:
            messagebox.showwarning(
                "Внимание",
                "Загрузите оба файла прежде чем проверять оплаты."
            )
            return

        try:
            self.result_df = back.analyze_payments(
                self.arenda_df, self.vypiska_df
            )
            back.save_to_excel(self.result_df, "result.xlsx")
        except Exception as e:
            messagebox.showerror("Ошибка при анализе", str(e))
            return

        # Обновляем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.result_df.iterrows():
            vals = (
                row["Номер гаража"],
                row["Дата оплаты"].strftime("%d.%m.%Y"),
                f"{row['Сумма оплаты']:,.2f} ₽",
                row["Статус"]
            )
            self.tree.insert("", "end", values=vals, tags=(row["Статус"],))

        messagebox.showinfo(
            "Готово",
            "Проверка завершена, отчет сохранён в result.xlsx."
        )

        # Открываем отдельное окно с графиком
        self.show_chart_window()

    def show_chart_window(self):
        if self.result_df is None:
            return

        # Окно для графика
        win = tk.Toplevel(self.root)
        win.title("График: количество гаражей по статусам")
        win.geometry("500x400")

        # Подготовка фигуры
        fig, ax = plt.subplots(figsize=(5, 4), tight_layout=True)

        # Данные для построения
        counts = self.result_df["Статус"].value_counts()
        statuses = counts.index.tolist()
        values = counts.values.tolist()

        # Цветовая карта
        color_map = {
            "Получен": "#28a745",
            "Просрочен": "#dc3545",
            "Срок не наступил": "#ffc107"
        }
        bar_colors = [color_map.get(s, "#6c757d") for s in statuses]

        # Явно задаём позиции по оси X
        x = list(range(len(statuses)))

        # Строим столбики
        ax.bar(x, values, color=bar_colors)

        # Подписываем оси и метки
        ax.set_title("Число гаражей по статусам")
        ax.set_ylabel("Гаражей")
        ax.set_xticks(x)
        ax.set_xticklabels(statuses, rotation=45, ha="right")

        # Встраиваем в tkinter
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
