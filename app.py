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
        self.root.geometry("800x600")

        # Регистрация шрифта с поддержкой кириллицы
        try:
            pdfmetrics.registerFont(TTFont('OpenSans', 'OpenSans-VariableFont_wdth,wght.ttf'))
            self.font_name = 'OpenSans'
        except:
            raise Exception("Шрифт OpenSans-VariableFont_wdth,wght.ttf не найден. Пожалуйста, скачайте его и поместите в директорию со скриптом.")

        # Данные о студентах, предметах и типах занятий
        self.students = [
            "Иванов Иван Иванович",
            "Петров Петр Петрович",
            "Сидоров Сидор Сидорович",
            "Козлов Алексей Викторович",
            "Смирнова Анна Сергеевна",
            "Васильева Мария Александровна",
            "Кузнецова Ольга Дмитриевна",
            "Попова Екатерина Ивановна"
        ]

        # Определяем подгруппы для каждого предмета отдельно
        self.subjects = {
            "Программирование": {
                "lectures": True,
                "practices": True,
                "labs": {},  # Нет лабораторных
                "students": self.students
            },
            "Математика": {
                "lectures": True,
                "practices": False,
                "labs": {  # Лабораторные с разделением на подгруппы
                    "Подгруппа 1": self.students[:4],  # Первые 4 студента
                    "Подгруппа 2": self.students[4:]   # Оставшиеся студенты
                },
                "students": self.students
            },
            "Физика": {
                "lectures": True,
                "practices": True,
                "labs": {  # Другие подгруппы для Физики
                    "Группа А": self.students[:3],  # Первые 3 студента
                    "Группа Б": self.students[3:6],  # Следующие 3 студента
                    "Группа В": self.students[6:]   # Оставшиеся студенты
                },
                "students": self.students
            }
        }

        # Загрузка данных о явке из файла
        self.attendance_data = self.load_attendance_data()

        # Главная форма
        self.create_main_form()

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
        # Убираем список предметов, оставляем только кнопки
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Проставить явку", command=self.mark_attendance).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Сгенерировать отчет", command=self.open_report_window).pack(side="left", padx=5)

    def mark_attendance(self):
        mark_window = tk.Toplevel(self.root)
        mark_window.title("Проставить явку")
        mark_window.geometry("500x700")

        # Выбор предмета
        ttk.Label(mark_window, text="Предмет:").pack(anchor="w", padx=10, pady=5)
        subject_combo = ttk.Combobox(mark_window, values=list(self.subjects.keys()))
        subject_combo.pack(fill="x", padx=10)

        # Выбор даты
        ttk.Label(mark_window, text="Дата (ДД.ММ.ГГГГ):").pack(anchor="w", padx=10, pady=5)
        date_entry = ttk.Entry(mark_window)
        date_entry.pack(fill="x", padx=10)
        date_entry.insert(0, datetime.datetime.now().strftime("%d.%m.%Y"))

        # Выбор типа занятия
        ttk.Label(mark_window, text="Тип занятия:").pack(anchor="w", padx=10, pady=5)
        type_combo = ttk.Combobox(mark_window, state="readonly")
        type_combo.pack(fill="x", padx=10)

        # Список студентов с выбором явки
        students_frame = ttk.LabelFrame(mark_window, text="Студенты")
        students_frame.pack(fill="both", expand=True, padx=10, pady=5)
        canvas = tk.Canvas(students_frame)
        scrollbar = ttk.Scrollbar(students_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

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
                types.append("Лекции")
            if self.subjects[subject]["practices"]:
                types.append("Практика")
            if self.subjects[subject]["labs"]:
                types.extend([f"Лабораторные - {sg}" for sg in self.subjects[subject]["labs"].keys()])
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

            if "Лабораторные" in class_type:
                subgroup = class_type.split(" - ")[1]
                students = self.subjects[subject]["labs"].get(subgroup, [])
            else:
                students = self.subjects[subject]["students"]

            for student in students:
                frame = ttk.Frame(scrollable_frame)
                frame.pack(fill="x", pady=2)
                ttk.Label(frame, text=student, width=40).pack(side="left")
                mark_combo = ttk.Combobox(frame, values=["есть", "н", "б"], state="readonly", width=10)
                mark_combo.set("есть")
                mark_combo.pack(side="left", padx=5)
                student_marks[student] = mark_combo

        subject_combo.bind("<<ComboboxSelected>>", update_types)
        type_combo.bind("<<ComboboxSelected>>", update_students)

        # Чекбокс для подтверждения преподавателем
        ttk.Checkbutton(mark_window, text="Подтвердить явку", variable=confirmed_var).pack(pady=5)

        def save_marks():
            subject = subject_combo.get()
            date = date_entry.get()
            class_type = type_combo.get()

            if not subject or not date or not class_type:
                messagebox.showerror("Ошибка", "Заполните все поля!")
                return

            try:
                datetime.datetime.strptime(date, "%d.%m.%Y")
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты! Используйте ДД.ММ.ГГГГ")
                return

            if date not in self.attendance_data[subject]:
                self.attendance_data[subject][date] = {}
            if class_type not in self.attendance_data[subject][date]:
                self.attendance_data[subject][date][class_type] = {}

            # Сохраняем явку студентов
            for student, mark_combo in student_marks.items():
                mark = mark_combo.get()
                self.attendance_data[subject][date][class_type][student] = mark

            # Сохраняем статус подтверждения
            self.attendance_data[subject][date][class_type]["confirmed"] = confirmed_var.get()

            self.save_attendance_data()
            messagebox.showinfo("Успех", "Явка проставлена и сохранена!")
            mark_window.destroy()

        ttk.Button(mark_window, text="Сохранить", command=save_marks).pack(pady=10)

    def open_report_window(self):
        report_window = tk.Toplevel(self.root)
        report_window.title("Сгенерировать отчет")
        report_window.geometry("400x300")

        ttk.Label(report_window, text="Выберите предмет для отчета:").pack(anchor="w", padx=10, pady=5)
        subject_combo = ttk.Combobox(report_window, values=list(self.subjects.keys()))
        subject_combo.pack(fill="x", padx=10)

        # Выбор интервала дат
        ttk.Label(report_window, text="Дата начала (ДД.ММ.ГГГГ):").pack(anchor="w", padx=10, pady=5)
        start_date_entry = ttk.Entry(report_window)
        start_date_entry.pack(fill="x", padx=10)
        start_date_entry.insert(0, (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%d.%m.%Y"))

        ttk.Label(report_window, text="Дата окончания (ДД.ММ.ГГГГ):").pack(anchor="w", padx=10, pady=5)
        end_date_entry = ttk.Entry(report_window)
        end_date_entry.pack(fill="x", padx=10)
        end_date_entry.insert(0, datetime.datetime.now().strftime("%d.%m.%Y"))

        def generate():
            subject = subject_combo.get()
            start_date = start_date_entry.get()
            end_date = end_date_entry.get()

            if not subject or not start_date or not end_date:
                messagebox.showerror("Ошибка", "Заполните все поля!")
                return

            try:
                start = datetime.datetime.strptime(start_date, "%d.%m.%Y")
                end = datetime.datetime.strptime(end_date, "%d.%m.%Y")
                if start > end:
                    messagebox.showerror("Ошибка", "Дата начала не может быть позже даты окончания!")
                    return
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты! Используйте ДД.ММ.ГГГГ")
                return

            self.generate_pdf(subject, start, end)
            report_window.destroy()

        ttk.Button(report_window, text="Сгенерировать PDF", command=generate).pack(pady=10)

    def wrap_text(self, c, text, max_width, font_name, font_size, wrap=True):
        """Разбивает текст на несколько строк, если он не помещается в заданную ширину. Если wrap=False, текст не переносится."""
        c.setFont(font_name, font_size)
        if not wrap:
            return [text]  # Возвращаем текст без переноса
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

    def draw_table(self, c, data, x_start, y_start, col_widths, row_height, height, title, date_headers, confirmed_status):
        """Отрисовка таблицы с учетом переноса на новую страницу и объединенных заголовков дат."""
        y = y_start
        rows_per_page = int((height - 150) / row_height) - 3  # Учитываем заголовок, подпись и строку подтверждения
        header_height = 40  # Увеличиваем высоту заголовка для двухуровневых заголовков

        # Отрисовка заголовков на первой странице
        x = x_start
        # Первая колонка: ФИО студента (без переноса)
        wrapped_lines = self.wrap_text(c, data[0][0], col_widths[0] - 10, self.font_name, 10, wrap=False)
        y_temp = y - 10
        for line in wrapped_lines:
            c.drawCentredString(x + col_widths[0] / 2, y_temp, line)
            y_temp -= 12
        x += col_widths[0]

        # Отрисовка объединенных заголовков дат
        date_idx = 0
        for date, types in date_headers.items():
            # Отрисовка даты (объединяет несколько колонок)
            date_width = sum(col_widths[date_idx + 1:date_idx + 1 + len(types)])
            c.drawCentredString(x + date_width / 2, y - 10, date)
            # Отрисовка типов занятий под датой
            for t_idx, class_type in enumerate(types):
                c.drawCentredString(x + col_widths[date_idx + 1] / 2, y - 25, class_type)
                x += col_widths[date_idx + 1]
                date_idx += 1

        # Начальная горизонтальная линия
        y -= header_height
        c.line(x_start, y, x_start + sum(col_widths), y)

        # Отрисовка строк данных
        for i in range(1, len(data)):  # Пропускаем первую строку (заголовок)
            # Проверяем, нужно ли начать новую страницу
            if (i - 1) % rows_per_page == 0 and i != 1:
                # Добавляем строку подтверждения перед переходом на новую страницу
                y -= row_height
                x = x_start
                c.drawCentredString(x + col_widths[0] / 2, y + 5, "Подтверждено:")
                x += col_widths[0]
                date_idx = 0
                for date, types in date_headers.items():
                    for class_type in types:
                        is_confirmed = confirmed_status.get(date, {}).get(class_type, False)
                        c.drawCentredString(x + col_widths[date_idx + 1] / 2, y + 5, "✓" if is_confirmed else "")
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

                # Отрисовка заголовков на новой странице
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
                    for t_idx, class_type in enumerate(types):
                        c.drawCentredString(x + col_widths[date_idx + 1] / 2, y - 25, class_type)
                        x += col_widths[date_idx + 1]
                        date_idx += 1

                y -= header_height
                c.line(x_start, y, x_start + sum(col_widths), y)

            # Отрисовка строк
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

            # Отрисовка горизонтальных линий
            c.line(x_start, y, x_start + sum(col_widths), y)

        # Последняя строка таблицы
        y -= row_height
        c.line(x_start, y, x_start + sum(col_widths), y)

        # Добавляем строку подтверждения в конце таблицы
        y -= row_height
        x = x_start
        c.drawCentredString(x + col_widths[0] / 2, y + 5, "Подтверждено:")
        x += col_widths[0]
        date_idx = 0
        for date, types in date_headers.items():
            for class_type in types:
                is_confirmed = confirmed_status.get(date, {}).get(class_type, False)
                c.drawCentredString(x + col_widths[date_idx + 1] / 2, y + 5, "✓" if is_confirmed else "")
                x += col_widths[date_idx + 1]
                date_idx += 1
        c.line(x_start, y, x_start + sum(col_widths), y)

        # Вертикальные линии
        x = x_start
        for w in col_widths:
            c.line(x, y_start, x, y - row_height)
            x += w
        c.line(x, y_start, x, y - row_height)

    def generate_pdf(self, subject, start_date, end_date):
        c = canvas.Canvas(f"attendance_report_{subject}.pdf", pagesize=landscape(A4))
        width, height = landscape(A4)  # Размеры страницы: ширина 842, высота 595

        # Фильтрация дат в заданном интервале
        attendance = self.attendance_data.get(subject, {})
        filtered_dates = {}
        for date_str in attendance:
            try:
                date = datetime.datetime.strptime(date_str, "%d.%m.%Y")
                if start_date <= date <= end_date:
                    filtered_dates[date_str] = attendance[date_str]
            except ValueError:
                continue

        # Если есть лекции и практики (на одном листе)
        if self.subjects[subject]["practices"]:
            c.setFont(self.font_name, 16)
            c.drawString(50, height - 50, subject)
            c.setFont(self.font_name, 10)

            # Собираем все даты и типы занятий
            dates_types = {}
            confirmed_status = {}
            for date, types in filtered_dates.items():
                for class_type in types:
                    if "Лабораторные" not in class_type:
                        if date not in dates_types:
                            dates_types[date] = []
                            confirmed_status[date] = {}
                        if class_type not in dates_types[date]:
                            dates_types[date].append(class_type)
                            confirmed_status[date][class_type] = filtered_dates[date][class_type].get("confirmed", False)

            # Получаем всех студентов
            students = self.subjects[subject]["students"]

            # Создаем таблицу
            dates = sorted(dates_types.keys())
            header = ["ФИО студента"]
            column_mapping = []  # Для отслеживания соответствия колонок и типов занятий
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

            # Отрисовка таблицы
            x_start = 50
            y_start = height - 80
            col_widths = [150] + [60] * (len(header) - 1)  # Ширина колонок
            row_height = 20

            self.draw_table(c, data, x_start, y_start, col_widths, row_height, height, subject, dates_types, confirmed_status)

            # Подпись преподавателя
            c.setFont(self.font_name, 12)
            c.drawString(50, 50, "Подпись преподавателя: ____________________")
            c.showPage()

        # Если есть лабораторные (отдельные листы с разделением по подгруппам)
        if self.subjects[subject]["labs"]:
            # Лекции на отдельном листе
            if self.subjects[subject]["lectures"]:
                c.setFont(self.font_name, 16)
                c.drawString(50, height - 50, f"{subject} - Лекции")
                c.setFont(self.font_name, 10)

                dates = sorted([d for d, t in filtered_dates.items() if "Лекции" in t])
                header = ["ФИО студента"] + dates
                data = [header]
                students = self.subjects[subject]["students"]
                dates_types = {date: ["Лекции"] for date in dates}
                confirmed_status = {date: {"Лекции": filtered_dates[date]["Лекции"].get("confirmed", False)} for date in dates}

                for student in sorted(students):
                    row = [student]
                    for date in dates:
                        mark = filtered_dates.get(date, {}).get("Лекции", {}).get(student, "")
                        row.append(mark)
                    data.append(row)

                # Отрисовка таблицы
                x_start = 50
                y_start = height - 80
                col_widths = [150] + [40] * (len(header) - 1)
                row_height = 20

                self.draw_table(c, data, x_start, y_start, col_widths, row_height, height, f"{subject} - Лекции", dates_types, confirmed_status)

                c.setFont(self.font_name, 12)
                c.drawString(50, 50, "Подпись преподавателя: ____________________")
                c.showPage()

            # Лабораторные по подгруппам (каждая на отдельном листе)
            for subgroup, students in self.subjects[subject]["labs"].items():
                c.setFont(self.font_name, 16)
                c.drawString(50, height - 50, f"{subject} - Лабораторные - {subgroup}")
                c.setFont(self.font_name, 10)

                dates = sorted([d for d, t in filtered_dates.items() if f"Лабораторные - {subgroup}" in t])
                header = ["ФИО студента"] + dates
                data = [header]
                dates_types = {date: [f"Лабораторные - {subgroup}"] for date in dates}
                confirmed_status = {date: {f"Лабораторные - {subgroup}": filtered_dates[date][f"Лабораторные - {subgroup}"].get("confirmed", False)} for date in dates}

                for student in sorted(students):
                    row = [student]
                    for date in dates:
                        mark = filtered_dates.get(date, {}).get(f"Лабораторные - {subgroup}", {}).get(student, "")
                        row.append(mark)
                    data.append(row)

                # Отрисовка таблицы
                x_start = 50
                y_start = height - 80
                col_widths = [150] + [40] * (len(header) - 1)
                row_height = 20

                self.draw_table(c, data, x_start, y_start, col_widths, row_height, height, f"{subject} - Лабораторные - {subgroup}", dates_types, confirmed_status)

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