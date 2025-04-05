import tkinter as tk
from tkinter import ttk, messagebox
import datetime
import json
import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учет посещаемости студентов")
        self.root.geometry("400x150")
        self.root.configure(padx=20, pady=20)

        # Регистрация шрифта
        pdfmetrics.registerFont(TTFont('OpenSans', 'OpenSans-VariableFont_wdth,wght.ttf'))
        self.font_name = 'OpenSans'

        self.load_config()
        self.attendance_data = self.load_attendance_data()
        self.create_main_form()

    def load_config(self):
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                self.students = config["students"]
                self.subjects = config["subjects"]
        else:
            raise Exception("Файл config.json не найден!")

    def load_attendance_data(self):
        if os.path.exists("attendance_data.json"):
            with open("attendance_data.json", "r", encoding="utf-8") as f:
                return json.load(f)
        else:
            return {subject: {} for subject in self.subjects}

    def save_attendance_data(self):
        with open("attendance_data.json", "w", encoding="utf-8") as f:
            json.dump(self.attendance_data, f, ensure_ascii=False, indent=4)

    def create_main_form(self):
        # Заголовок
        ttk.Label(self.root, text="Учет посещаемости студентов", font=("Helvetica", 16, "bold")).pack(pady=(0, 20))

        # Рамка для кнопок
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", pady=10)

        # Кнопки с увеличенными размерами и отступами
        ttk.Button(button_frame, text="Проставить явку", command=self.mark_attendance, width=20).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Сгенерировать отчет", command=self.open_report_window, width=20).pack(side="left", padx=10)

    def mark_attendance(self):
        mark_window = tk.Toplevel(self.root)
        mark_window.title("Проставить явку")
        mark_window.geometry("600x800")  # Увеличил размер окна
        mark_window.configure(padx=20, pady=20)

        # Заголовок
        ttk.Label(mark_window, text="Проставить явку", font=("Helvetica", 14, "bold")).pack(pady=(0, 20))

        # Поля ввода в рамке
        input_frame = ttk.LabelFrame(mark_window, text="Параметры занятия", padding=10)
        input_frame.pack(fill="x", pady=10)

        ttk.Label(input_frame, text="Предмет:").grid(row=0, column=0, sticky="w", pady=5)
        subject_combo = ttk.Combobox(input_frame, values=list(self.subjects.keys()), width=40)
        subject_combo.grid(row=0, column=1, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Дата (ДД.ММ.ГГГГ):").grid(row=1, column=0, sticky="w", pady=5)
        date_entry = ttk.Entry(input_frame, width=40)
        date_entry.grid(row=1, column=1, pady=5, sticky="ew")
        date_entry.insert(0, datetime.datetime.now().strftime("%d.%m.%Y"))

        ttk.Label(input_frame, text="Тип занятия:").grid(row=2, column=0, sticky="w", pady=5)
        type_combo = ttk.Combobox(input_frame, state="readonly", width=40)
        type_combo.grid(row=2, column=1, pady=5, sticky="ew")

        input_frame.columnconfigure(1, weight=1)

        # Список студентов
        students_frame = ttk.LabelFrame(mark_window, text="Студенты", padding=10)
        students_frame.pack(fill="both", expand=True, pady=10)

        canvas = tk.Canvas(students_frame)
        scrollbar = ttk.Scrollbar(students_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        student_marks = {}
        confirmed_var = tk.BooleanVar(value=False)

        def update_types(event):
            subject = subject_combo.get()
            if not subject:
                return
            types = []
            if self.subjects[subject]["lectures"]:
                types.append("Лекция")
            if self.subjects[subject]["practices"]:
                types.append("Практика")
            if self.subjects[subject]["labs"]:
                types.extend([f"Лабораторная работа - {sg}" for sg in self.subjects[subject]["labs"].keys()])
            type_combo["values"] = types
            type_combo.set(types[0] if types else "")

        def update_students(event):
            for widget in scrollable_frame.winfo_children():
                widget.destroy()
            student_marks.clear()

            subject = subject_combo.get()
            class_type = type_combo.get()
            if not subject or not class_type:
                return

            if "Лабораторная работа" in class_type:
                subgroup = class_type.split(" - ")[1]
                students = self.subjects[subject]["labs"].get(subgroup, [])
            else:
                students = self.subjects[subject]["students"]

            for student in students:
                frame = ttk.Frame(scrollable_frame)
                frame.pack(fill="x", pady=5)
                ttk.Label(frame, text=student, width=45).pack(side="left", padx=5)
                mark_combo = ttk.Combobox(frame, values=["есть", "н", "б"], state="readonly", width=10)
                mark_combo.set("есть")
                mark_combo.pack(side="left")
                student_marks[student] = mark_combo

        subject_combo.bind("<<ComboboxSelected>>", update_types)
        type_combo.bind("<<ComboboxSelected>>", update_students)

        # Чекбокс и кнопка
        ttk.Checkbutton(mark_window, text="Подтвердить явку", variable=confirmed_var).pack(pady=10)
        ttk.Button(mark_window, text="Сохранить", command=lambda: save_marks(), width=20).pack(pady=10)

        def save_marks():
            subject = subject_combo.get()
            date = date_entry.get()
            class_type = type_combo.get()

            if not subject or not date or not class_type:
                messagebox.showerror("Ошибка", "Заполните все поля!", parent=mark_window)
                return

            try:
                datetime.datetime.strptime(date, "%d.%m.%Y")
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты! Используйте ДД.ММ.ГГГГ", parent=mark_window)
                return

            if date not in self.attendance_data[subject]:
                self.attendance_data[subject][date] = {}
            if class_type not in self.attendance_data[subject][date]:
                self.attendance_data[subject][date][class_type] = {}

            for student, mark_combo in student_marks.items():
                self.attendance_data[subject][date][class_type][student] = mark_combo.get()

            self.attendance_data[subject][date][class_type]["confirmed"] = confirmed_var.get()
            self.save_attendance_data()
            messagebox.showinfo("Успех", "Явка проставлена и сохранена!", parent=mark_window)
            mark_window.destroy()

    def open_report_window(self):
        report_window = tk.Toplevel(self.root)
        report_window.title("Сгенерировать отчет")
        report_window.geometry("500x310")  # Увеличил размер окна
        report_window.configure(padx=20, pady=20)

        # Заголовок
        ttk.Label(report_window, text="Сгенерировать отчет", font=("Helvetica", 14, "bold")).pack(pady=(0, 20))

        # Поля ввода в рамке
        input_frame = ttk.LabelFrame(report_window, text="Параметры отчета", padding=10)
        input_frame.pack(fill="x", pady=10)

        ttk.Label(input_frame, text="Предмет:").grid(row=0, column=0, sticky="w", pady=5)
        subject_combo = ttk.Combobox(input_frame, values=list(self.subjects.keys()), width=40)
        subject_combo.grid(row=0, column=1, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Дата начала (ДД.ММ.ГГГГ):").grid(row=1, column=0, sticky="w", pady=5)
        start_date_entry = ttk.Entry(input_frame, width=40)
        start_date_entry.grid(row=1, column=1, pady=5, sticky="ew")
        start_date_entry.insert(0, (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%d.%m.%Y"))

        ttk.Label(input_frame, text="Дата окончания (ДД.ММ.ГГГГ):").grid(row=2, column=0, sticky="w", pady=5)
        end_date_entry = ttk.Entry(input_frame, width=40)
        end_date_entry.grid(row=2, column=1, pady=5, sticky="ew")
        end_date_entry.insert(0, datetime.datetime.now().strftime("%d.%m.%Y"))

        input_frame.columnconfigure(1, weight=1)

        # Кнопка
        ttk.Button(report_window, text="Сгенерировать PDF", command=lambda: generate(), width=20).pack(pady=20)

        def generate():
            subject = subject_combo.get()
            start_date = start_date_entry.get()
            end_date = end_date_entry.get()

            if not subject or not start_date or not end_date:
                messagebox.showerror("Ошибка", "Заполните все поля!", parent=report_window)
                return

            try:
                start = datetime.datetime.strptime(start_date, "%d.%m.%Y")
                end = datetime.datetime.strptime(end_date, "%d.%m.%Y")
                if start > end:
                    messagebox.showerror("Ошибка", "Дата начала не может быть позже даты окончания!", parent=report_window)
                    return
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты! Используйте ДД.ММ.ГГГГ", parent=report_window)
                return

            self.generate_pdf(subject, start, end)
            report_window.destroy()

    # Оставшиеся методы (wrap_text, draw_table, generate_pdf) без изменений
    def wrap_text(self, c, text, max_width, font_name, font_size, wrap=True):
        c.setFont(font_name, font_size)
        if not wrap:
            return [text]
        lines = []
        words = text.split()
        current_line = ""
        for word in words:
            test_line = current_line + (word if not current_line else " " + word)
            if c.stringWidth(test_line, font_name, font_size) <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        return lines

    def draw_table(self, c, data, x_start, y_start, col_widths, row_height, height, title, date_headers, confirmed_status, hide_class_type=False):
        y = y_start
        rows_per_page = int((height - 150) / row_height) - 3
        header_height = 40
        x = x_start
        wrapped_lines = self.wrap_text(c, data[0][0], col_widths[0] - 10, self.font_name, 10, wrap=False)
        y_temp = y - 10
        for line in wrapped_lines:
            c.drawCentredString(x + col_widths[0] / 2, y_temp, line)
            y_temp -= 12
        x += col_widths[0]
        date_idx = 0
        for date, types in date_headers.items():
            date_width = sum(col_widths[date_idx + 1:date_idx + 1 + len(types)])
            c.drawCentredString(x + date_width / 2, y - 10, date)
            if not hide_class_type:
                for t_idx, class_type in enumerate(types):
                    c.drawCentredString(x + col_widths[date_idx + 1] / 2, y - 25, class_type)
                    x += col_widths[date_idx + 1]
                    date_idx += 1
            else:
                for _ in types:
                    x += col_widths[date_idx + 1]
                    date_idx += 1
        y -= header_height
        c.line(x_start, y, x_start + sum(col_widths), y)
        for i in range(1, len(data)):
            if (i - 1) % rows_per_page == 0 and i != 1:
                y -= row_height
                x = x_start
                c.drawCentredString(x + col_widths[0] / 2, y + 5, "Подтверждено:")
                x += col_widths[0]
                date_idx = 0
                for date, types in date_headers.items():
                    for class_type in types:
                        is_confirmed = confirmed_status.get(date, {}).get(class_type, False)
                        c.drawCentredString(x + col_widths[date_idx + 1] / 2, y + 5, "+" if is_confirmed else "-")
                        x += col_widths[date_idx + 1]
                        date_idx += 1
                c.line(x_start, y, x_start + sum(col_widths), y)
                c.setFont(self.font_name, 12)
                c.drawString(50, 50, "Подпись преподавателя: ____________________")
                c.showPage()
                c.setFont(self.font_name, 16)
                c.drawString(50, height - 50, title)
                c.setFont(self.font_name, 10)
                y = height - 80
                x = x_start
                wrapped_lines = self.wrap_text(c, data[0][0], col_widths[0] - 10, self.font_name, 10, wrap=False)
                y_temp = y - 10
                for line in wrapped_lines:
                    c.drawCentredString(x + col_widths[0] / 2, y_temp, line)
                    y_temp -= 12
                x += col_widths[0]
                date_idx = 0
                for date, types in date_headers.items():
                    date_width = sum(col_widths[date_idx + 1:date_idx + 1 + len(types)])
                    c.drawCentredString(x + date_width / 2, y - 10, date)
                    if not hide_class_type:
                        for t_idx, class_type in enumerate(types):
                            c.drawCentredString(x + col_widths[date_idx + 1] / 2, y - 25, class_type)
                            x += col_widths[date_idx + 1]
                            date_idx += 1
                    else:
                        for _ in types:
                            x += col_widths[date_idx + 1]
                            date_idx += 1
                y -= header_height
                c.line(x_start, y, x_start + sum(col_widths), y)
            y -= row_height
            x = x_start
            for j in range(len(data[i])):
                wrap = False if j == 0 else True
                wrapped_lines = self.wrap_text(c, data[i][j], col_widths[j] - 10, self.font_name, 10, wrap=wrap)
                y_temp = y + 5
                for line in wrapped_lines[:2]:
                    c.drawCentredString(x + col_widths[j] / 2, y_temp, line)
                    y_temp -= 12
                x += col_widths[j]
            c.line(x_start, y, x_start + sum(col_widths), y)
        y -= row_height
        x = x_start
        c.drawCentredString(x + col_widths[0] / 2, y + 5, "Подтверждено:")
        x += col_widths[0]
        date_idx = 0
        for date, types in date_headers.items():
            for class_type in types:
                is_confirmed = confirmed_status.get(date, {}).get(class_type, False)
                c.drawCentredString(x + col_widths[date_idx + 1] / 2, y + 5, "+" if is_confirmed else "-")
                x += col_widths[date_idx + 1]
                date_idx += 1
        c.line(x_start, y, x_start + sum(col_widths), y)
        x = x_start
        for w in col_widths:
            c.line(x, y_start, x, y)
            x += w
        c.line(x, y_start, x, y)

    def generate_pdf(self, subject, start_date, end_date):
        c = canvas.Canvas(f"attendance_report_{subject}.pdf", pagesize=landscape(A4))
        width, height = landscape(A4)
        attendance = self.attendance_data.get(subject, {})
        filtered_dates = {}
        for date_str in attendance:
            try:
                date = datetime.datetime.strptime(date_str, "%d.%m.%Y")
                if start_date <= date <= end_date:
                    filtered_dates[date_str] = attendance[date_str]
            except ValueError:
                continue
        if self.subjects[subject]["lectures"] or self.subjects[subject]["practices"]:
            c.setFont(self.font_name, 16)
            c.drawString(50, height - 50, subject)
            c.setFont(self.font_name, 10)
            dates_types = {}
            confirmed_status = {}
            for date, types in filtered_dates.items():
                for class_type in types:
                    if "Лабораторная работа" not in class_type:
                        if date not in dates_types:
                            dates_types[date] = []
                            confirmed_status[date] = {}
                        if class_type not in dates_types[date]:
                            dates_types[date].append(class_type)
                            confirmed_status[date][class_type] = types[class_type].get("confirmed", False)
            if dates_types:
                students = self.subjects[subject]["students"]
                dates = sorted(dates_types.keys())
                header = ["ФИО студента"]
                column_mapping = []
                for date in dates:
                    for class_type in dates_types[date]:
                        header.append(class_type)
                        column_mapping.append((date, class_type))
                data = [header]
                for student in sorted(students):
                    row = [student]
                    for date, class_type in column_mapping:
                        mark = filtered_dates.get(date, {}).get(class_type, {}).get(student, "")
                        row.append(mark)
                    data.append(row)
                x_start = 50
                y_start = height - 80
                col_widths = [150] + [60] * (len(header) - 1)
                row_height = 20
                hide_class_type = bool(self.subjects[subject]["labs"]) and not self.subjects[subject]["practices"]
                self.draw_table(c, data, x_start, y_start, col_widths, row_height, height, subject, dates_types, confirmed_status, hide_class_type)
                c.setFont(self.font_name, 12)
                c.drawString(50, 50, "Подпись преподавателя: ____________________")
                c.showPage()
        if self.subjects[subject]["labs"]:
            for subgroup, students in self.subjects[subject]["labs"].items():
                c.setFont(self.font_name, 16)
                c.drawString(50, height - 50, f"{subject} - Лабораторная работа - {subgroup}")
                c.setFont(self.font_name, 10)
                dates = sorted([d for d, t in filtered_dates.items() if f"Лабораторная работа - {subgroup}" in t])
                header = ["ФИО студента"] + dates
                data = [header]
                dates_types = {date: [f"Лабораторная работа - {subgroup}"] for date in dates}
                confirmed_status = {date: {f"Лабораторная работа - {subgroup}": filtered_dates[date][f"Лабораторная работа - {subgroup}"].get("confirmed", False)} for date in dates}
                for student in sorted(students):
                    row = [student]
                    for date in dates:
                        mark = filtered_dates.get(date, {}).get(f"Лабораторная работа - {subgroup}", {}).get(student, "")
                        row.append(mark)
                    data.append(row)
                x_start = 50
                y_start = height - 80
                col_widths = [150] + [60] * (len(header) - 1)
                row_height = 20
                self.draw_table(c, data, x_start, y_start, col_widths, row_height, height, f"{subject} - Лабораторная работа - {subgroup}", dates_types, confirmed_status, hide_class_type=True)
                c.setFont(self.font_name, 12)
                c.drawString(50, 50, "Подпись преподавателя: ____________________")
                c.showPage()
        c.save()
        messagebox.showinfo("Успех", f"PDF сгенерирован: attendance_report_{subject}.pdf")

def main():
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()