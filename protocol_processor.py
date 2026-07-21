import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import docx
import re
from datetime import datetime
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

class ProtocolProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Обробник протоколів - Конвертер у таблиці")
        self.root.geometry("800x600")
        
        self.input_file = None
        self.output_data = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # Заголовок
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.pack(fill=tk.X)
        
        ttk.Label(title_frame, text="Обробник протоколів засідань", 
                 font=('Arial', 16, 'bold')).pack()
        
        # Зона для drag & drop
        drop_frame = ttk.LabelFrame(self.root, text="1. Виберіть файл протоколу", padding="20")
        drop_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.drop_label = ttk.Label(drop_frame, 
                                    text="Перетягніть файл .docx сюди\nабо натисніть кнопку нижче",
                                    font=('Arial', 12),
                                    background='#e8f4f8',
                                    relief=tk.RIDGE,
                                    padding=50)
        self.drop_label.pack(fill=tk.BOTH, expand=True)
        
        # Реєструємо drag & drop
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)
        
        # Кнопка вибору файлу
        btn_frame = ttk.Frame(drop_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Вибрати файл", 
                  command=self.select_file).pack(side=tk.LEFT, padx=5)
        
        # Інформація про вибраний файл
        self.file_info = ttk.Label(drop_frame, text="", foreground='green')
        self.file_info.pack()
        
        # Зона результату
        result_frame = ttk.LabelFrame(self.root, text="2. Результат обробки", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Текстове поле для результату
        text_frame = ttk.Frame(result_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.result_text = tk.Text(text_frame, height=10, yscrollcommand=scrollbar.set,
                                   font=('Courier', 9))
        self.result_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.result_text.yview)
        
        # Кнопки дій
        action_frame = ttk.Frame(result_frame)
        action_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(action_frame, text="Обробити файл", 
                  command=self.process_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Зберегти результат", 
                  command=self.save_result).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Очистити", 
                  command=self.clear_all).pack(side=tk.LEFT, padx=5)
        
        # Інфо про сортування
        ttk.Label(action_frame, text="📊 Автоматичне сортування: дата → виконавець", 
                 foreground='blue').pack(side=tk.LEFT, padx=20)
    
    def on_drop(self, event):
        """Обробка drag & drop"""
        file_path = event.data
        # Видаляємо фігурні дужки, якщо вони є
        file_path = file_path.strip('{}')
        
        if file_path.endswith('.docx'):
            self.input_file = file_path
            self.file_info.config(text=f"✓ Вибрано: {file_path.split('/')[-1]}")
            self.drop_label.config(text=f"Файл завантажено:\n{file_path.split('/')[-1]}")
        else:
            messagebox.showerror("Помилка", "Будь ласка, виберіть файл .docx")
    
    def select_file(self):
        """Вибір файлу через діалог"""
        file_path = filedialog.askopenfilename(
            title="Виберіть файл протоколу",
            filetypes=[("Word документи", "*.docx"), ("Всі файли", "*.*")]
        )
        
        if file_path:
            self.input_file = file_path
            self.file_info.config(text=f"✓ Вибрано: {file_path.split('/')[-1]}")
            self.drop_label.config(text=f"Файл завантажено:\n{file_path.split('/')[-1]}")
    
    def extract_protocol_data(self, doc_path):
        """Витягує дані з протоколу"""
        doc = docx.Document(doc_path)
        full_text = '\n'.join([para.text for para in doc.paragraphs])
        
        # Знаходимо номер протоколу і дату
        protocol_match = re.search(r'ПРОТОКОЛ\s*№\s*(\d+)', full_text, re.IGNORECASE)
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', full_text)
        
        protocol_num = protocol_match.group(1) if protocol_match else "?"
        protocol_date = date_match.group(1) if date_match else "?"
        
        # Знаходимо блок "ВИРІШЕНО"
        decisions_start = -1
        for keyword in ['ВИРІШЕНО', 'вирішено', 'Вирішено']:
            pos = full_text.find(keyword)
            if pos != -1:
                decisions_start = pos
                break
        
        if decisions_start == -1:
            work_text = full_text
        else:
            work_text = full_text[decisions_start:]
        
        # Розділяємо на рішення за паттерном "Ім'я ПРІЗВИЩЕ"
        decisions = []
        lines = work_text.split('\n')
        
        current_decision = None
        current_text = []
        current_executor = None
        current_deadline = None
        decision_counter = 1
        
        # Список можливих виконавців
        executors = [
            (r'Сергію БОЧІНУ', 'Бочін Сергій'),
            (r'Ларисі ГАВРИЛЕНКО', 'Гавриленко Лариса'),
            (r'Вячеславу БІЛОГУБУ', 'Білогуб Вячеслав'),
            (r'Олексію БУКА', 'Бука Олексій'),
            (r'Юлії КОРЕНКОВІЙ', 'Коренкова Юлія'),
            (r'Михайлу ЛЕОНТЬЄВУ', 'Леонтьєв Михайло'),
            (r'Андрію ЗАЛУСЬКОМУ', 'Залуський Андрій'),
            (r'Андрію ІЩЕНКО', 'Іщенко Андрій'),
            (r'Ніні МОСКАЛЕНКО', 'Москаленко Ніна'),
            (r'Олегу ОЛІНКЕВИЧ', 'Олінкевич Олег'),
            (r'Всім присутнім', 'Всім присутнім')
        ]
        
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            
            if not line_stripped:
                continue
            
            # Перевіряємо, чи починається новий запис (є ім'я виконавця)
            is_new_decision = False
            found_executor = None
            executor_pattern_matched = None
            
            for pattern, executor_name in executors:
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    is_new_decision = True
                    found_executor = executor_name
                    executor_pattern_matched = pattern
                    break
            
            # Перевіряємо також на номерні записи (1., 2., тощо)
            numbered_match = re.match(r'^(\d+)\.?\s+', line_stripped)
            if numbered_match and len(numbered_match.group(1)) <= 3:
                is_new_decision = True
                decision_counter = int(numbered_match.group(1))
            
            if is_new_decision:
                # Зберігаємо попереднє рішення
                if current_text:
                    text_combined = ' '.join(current_text).strip()
                    text_combined = re.sub(r'\s+', ' ', text_combined)
                    
                    decisions.append({
                        'num': str(decision_counter - 1) if current_decision else str(len(decisions) + 1),
                        'text': text_combined,
                        'executor': current_executor or 'Не вказано',
                        'date_given': protocol_date,
                        'deadline': current_deadline or 'Не вказано'
                    })
                
                # Починаємо нове рішення
                current_decision = decision_counter
                
                # Видаляємо виконавця з тексту доручення
                clean_line = line_stripped
                if executor_pattern_matched:
                    clean_line = re.sub(executor_pattern_matched, '', clean_line, flags=re.IGNORECASE).strip()
                
                # Перша літера з великої
                if clean_line:
                    clean_line = clean_line[0].upper() + clean_line[1:] if len(clean_line) > 1 else clean_line.upper()
                
                current_text = [clean_line] if clean_line else []
                current_executor = found_executor
                current_deadline = None
                decision_counter += 1
                
            elif current_decision:
                # Шукаємо термін виконання
                if 'ермін' in line:
                    deadline_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                    if deadline_match:
                        current_deadline = deadline_match.group(1)
                    elif 'постійно' in line.lower():
                        current_deadline = 'постійно'
                    # Не додаємо рядок з терміном до тексту
                else:
                    current_text.append(line_stripped)
        
        # Додаємо останнє рішення
        if current_text:
            text_combined = ' '.join(current_text).strip()
            text_combined = re.sub(r'\s+', ' ', text_combined)
            
            decisions.append({
                'num': str(len(decisions) + 1),
                'text': text_combined,
                'executor': current_executor or 'Не вказано',
                'date_given': protocol_date,
                'deadline': current_deadline or 'Не вказано'
            })
        
        return decisions
    
    def sort_data(self, data):
        """Автоматичне сортування: спочатку за датою, потім за виконавцем
        + пункт про наступне засідання йде останнім для першої дати"""
        
        def sort_key(item):
            # Перетворюємо дату для сортування
            try:
                if item['deadline'] == 'постійно':
                    date = datetime.max  # "Постійно" - в кінець
                elif item['deadline'] != 'Не вказано':
                    date = datetime.strptime(item['deadline'], '%d.%m.%Y')
                else:
                    date = datetime.max  # Невідомі дати - в кінець
            except:
                date = datetime.max
            
            # Виконавець як другий ключ
            executor = item['executor'] if item['executor'] != 'Не вказано' else 'ЯЯЯ'
            
            # Перевіряємо, чи це пункт про засідання
            is_meeting_item = 'наступне засідання' in item['text'].lower() or 'призначити наступне' in item['text'].lower()
            
            # Якщо це засідання - додаємо маркер, щоб воно йшло останнім в межах своєї дати
            priority = 1 if is_meeting_item else 0
            
            return (date, priority, executor)
        
        return sorted(data, key=sort_key)
    
    def update_display(self):
        """Оновлює відображення даних з автоматичним сортуванням"""
        if not self.output_data:
            return
        
        sorted_data = self.sort_data(self.output_data.copy())
        
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, f"Знайдено доручень: {len(sorted_data)}\n")
        self.result_text.insert(tk.END, f"Сортування: за датою, потім за виконавцем\n\n")
        
        current_date = None
        for item in sorted_data:
            # Додаємо заголовок дати при зміні
            if current_date != item['deadline']:
                current_date = item['deadline']
                self.result_text.insert(tk.END, f"\n{'='*60}\n", 'header')
                self.result_text.insert(tk.END, f"📅 ТЕРМІН: {current_date}\n", 'header')
                self.result_text.insert(tk.END, f"{'='*60}\n\n", 'header')
            
            self.result_text.insert(tk.END, 
                f"№ {item['num']}: {item['text'][:80]}...\n"
                f"   👤 Виконавець: {item['executor']}\n"
                f"   📆 Термін: {item['deadline']}\n\n")
        
        # Налаштування тегів
        self.result_text.tag_config('header', foreground='blue', font=('Arial', 10, 'bold'))
    
    def process_file(self):
        """Обробка файлу протоколу"""
        if not self.input_file:
            messagebox.showwarning("Увага", "Спочатку виберіть файл протоколу!")
            return
        
        try:
            self.output_data = self.extract_protocol_data(self.input_file)
            self.update_display()
            messagebox.showinfo("Успіх", f"Оброблено {len(self.output_data)} доручень!")
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося обробити файл:\n{str(e)}")
    
    def create_table_document(self, output_path):
        """Створює документ з таблицею"""
        doc = docx.Document()
        
        # Налаштування сторінки (A4 landscape)
        section = doc.sections[0]
        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        
        # Сортуємо дані автоматично
        sorted_data = self.sort_data(self.output_data.copy())
        
        # Знаходимо дату наступного засідання (найближча дата з доручень)
        next_meeting_date = "Не вказано"
        if sorted_data:
            for item in sorted_data:
                if item['deadline'] != 'Не вказано' and item['deadline'] != 'постійно':
                    try:
                        datetime.strptime(item['deadline'], '%d.%m.%Y')
                        next_meeting_date = item['deadline']
                        break
                    except:
                        pass
        
        # Додаємо шапку з міжрядковим інтервалом 1
        header1 = doc.add_paragraph()
        run1 = header1.add_run("ІНФОРМАЦІЯ")
        run1.font.name = 'Times New Roman'
        run1.font.size = Pt(14)
        run1.bold = True
        header1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header1.paragraph_format.line_spacing = 1.0
        
        header2 = doc.add_paragraph()
        run2 = header2.add_run("з виконання протокольних доручень голови районної адміністрації Ільченко С.В.")
        run2.font.name = 'Times New Roman'
        run2.font.size = Pt(14)
        header2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header2.paragraph_format.line_spacing = 1.0
        
        header3 = doc.add_paragraph()
        run3 = header3.add_run(f"станом на {next_meeting_date}")
        run3.font.name = 'Times New Roman'
        run3.font.size = Pt(14)
        header3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header3.paragraph_format.line_spacing = 1.0
        
        doc.add_paragraph()  # Пустий рядок
        
        # Створюємо таблицю
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        
        # Встановлюємо ширину колонок
        # 1 см = 0.3937 дюйма, 5 см = 1.9685 дюйма
        # Загальна ширина: 10.69 дюймів (11.69 - 2*0.5 margins)
        col_widths = [
            Inches(0.3937),  # № з/п – 1 см
            Inches(5.2),     # Доручення – все інше місце
            Inches(1.3),     # Виконавці – по ширині прізвищ
            Inches(1.0),     # Термін надання – по ширині дати
            Inches(1.0),     # Термін виконання – по ширині дати
            Inches(1.9685)   # Примітки – 5 см
        ]
        
        for i, width in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = width
        
        # Заголовки
        headers = ['№ з/п', 'Доручення', 'Виконавці', 'Термін надання доручення', 
                   'Термін виконання доручення', 'Примітки']
        header_cells = table.rows[0].cells
        
        for i, header in enumerate(headers):
            header_cells[i].text = header
            # Форматування заголовків
            for paragraph in header_cells[i].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
                    run.font.bold = True
            # Вирівнювання по вертикалі по центру
            header_cells[i].vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Додаємо дані (відсортовані)
        for item in sorted_data:
            row_cells = table.add_row().cells
            
            # № з/п
            row_cells[0].text = item['num']
            self._format_cell(row_cells[0], WD_ALIGN_PARAGRAPH.CENTER)
            
            # Доручення – вирівнювання ПО ЛІВОМУ КРАЮ
            row_cells[1].text = item['text']
            self._format_cell(row_cells[1], WD_ALIGN_PARAGRAPH.LEFT)
            
            # Виконавці
            row_cells[2].text = item['executor']
            self._format_cell(row_cells[2], WD_ALIGN_PARAGRAPH.CENTER)
            
            # Термін надання
            row_cells[3].text = item['date_given']
            self._format_cell(row_cells[3], WD_ALIGN_PARAGRAPH.CENTER)
            
            # Термін виконання
            row_cells[4].text = item['deadline']
            self._format_cell(row_cells[4], WD_ALIGN_PARAGRAPH.CENTER)
            
            # Примітки
            row_cells[5].text = ''
            self._format_cell(row_cells[5], WD_ALIGN_PARAGRAPH.CENTER)
        
        # Зберігаємо
        doc.save(output_path)
    
    def _format_cell(self, cell, alignment):
        """Форматує комірку таблиці"""
        for paragraph in cell.paragraphs:
            paragraph.alignment = alignment
            paragraph.paragraph_format.line_spacing = 1.0
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)
        # Вирівнювання по вертикалі по центру
        cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def save_result(self):
        """Збереження результату"""
        if not self.output_data:
            messagebox.showwarning("Увага", "Спочатку обробіть файл!")
            return
        
        save_path = filedialog.asksaveasfilename(
            title="Зберегти результат",
            defaultextension=".docx",
            filetypes=[("Word документи", "*.docx"), ("Всі файли", "*.*")]
        )
        
        if save_path:
            try:
                self.create_table_document(save_path)
                messagebox.showinfo("Успіх", f"Таблицю збережено:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося зберегти файл:\n{str(e)}")
    
    def clear_all(self):
        """Очищення всього"""
        self.input_file = None
        self.output_data = []
        self.file_info.config(text="")
        self.drop_label.config(text="Перетягніть файл .docx сюди\nабо натисніть кнопку нижче")
        self.result_text.delete(1.0, tk.END)

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ProtocolProcessor(root)
    root.mainloop()
