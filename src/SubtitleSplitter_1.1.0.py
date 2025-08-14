import os
import sys
import subprocess
import logging
import re
import json
from tkinter import Tk, filedialog, messagebox, Frame, StringVar, IntVar, BooleanVar, Toplevel, Button, Label, Checkbutton, Entry, Text, Menu, PhotoImage, Scrollbar
from tkinter.ttk import Combobox
import keyboard
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError:
    logging.error("Библиотека tkinterdnd2 не установлена. Drag-and-drop не будет работать.")
    TkinterDnD = None
    DND_FILES = None

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Функция для получения пути к файлам ресурсов
def resource_path(relative_path):
    """Возвращает абсолютный путь к ресурсу, учитывая PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Функция для получения пути к файлу настроек
def settings_path():
    """Возвращает путь к settings.json в пользовательской директории."""
    if sys.platform.startswith('win'):
        base_path = os.getenv('APPDATA') or os.path.expanduser('~')
    else:
        base_path = os.path.expanduser('~/.config')
    config_dir = os.path.join(base_path, 'SubtitleSplitter')
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, 'settings.json')

class SubtitleSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Разделение субтитров")
        self.file_path_var = StringVar(value=os.path.expanduser("~/Desktop"))
        self.format_var = StringVar(value="ass")
        self.distribute_group_var = IntVar(value=1)
        self.distribute_multiple_var = IntVar(value=1)
        self.save_signs_ass_var = IntVar(value=1)  # Чекбокс для надписей
        self.show_group_option = BooleanVar(value=False)
        self.show_multiple_option = BooleanVar(value=False)
        self.show_signs_option = BooleanVar(value=False)  # Флаг для отображения чекбокса надписей
        self.settings_file = settings_path()
        self.show_update_var = BooleanVar(value=self.load_settings())
        logging.info(f"Инициализация: show_update_var={self.show_update_var.get()}, settings_file={self.settings_file}")

        self.actors = None
        self.group_lines = None
        self.multiple_actor_lines = None
        self.excluded_actor_lines = None
        self.sign_lines = None  # Список строк с надписями
        self.headers = None
        self.styles = None
        self.events = None
        self.all_actors = None

        # Создание меню
        self.menu_bar = Menu(self.root)
        self.root.config(menu=self.menu_bar)
        self.file_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Файл", menu=self.file_menu)
        self.file_menu.add_command(label="Открыть файл", command=self.choose_file)
        self.file_menu.add_command(label="Очистить поле", command=self.clear_field)
        self.file_menu.add_command(label="Настройки", command=self.show_settings)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Выход", command=self.on_closing)
        self.help_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Помощь", menu=self.help_menu)
        self.help_menu.add_command(label="Руководство", command=self.show_help)
        self.help_menu.add_command(label="О программе", command=self.show_about)
        self.help_menu.add_command(label="Информация об обновлениях", command=self.show_update_info)

        # Контекстное меню
        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Копировать", command=self.copy_text)
        self.context_menu.add_command(label="Вставить", command=self.paste_text)
        self.context_menu.add_command(label="Выделить все", command=self.select_all_text)

        # Горячие клавиши
        try:
            keyboard.add_hotkey('ctrl+c', self.copy_text)
            keyboard.add_hotkey('ctrl+v', self.paste_text)
            keyboard.add_hotkey('ctrl+a', self.select_all_text)
            keyboard.add_hotkey('ctrl+с', self.copy_text)
            keyboard.add_hotkey('ctrl+в', self.paste_text)
            keyboard.add_hotkey('ctrl+ф', self.select_all_text)
            logging.info("Горячие клавиши через keyboard установлены")
        except Exception as e:
            logging.error(f"Ошибка установки горячих клавиш через keyboard: {e}")
            self.root.bind("<Control-c>", self.copy_text)
            self.root.bind("<Control-v>", self.paste_text)
            self.root.bind("<Control-a>", self.select_all_text)

        # Настройка окна
        window_width = 450
        window_height = 300
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(False, False)
        self.root.configure(bg="#eceff1")

        try:
            self.root.iconbitmap(resource_path("favicon.ico"))
        except Exception as e:
            logging.warning(f"Не удалось установить иконку favicon.ico: {e}")

        # Основной фрейм
        self.main_frame = Frame(root, bg="#ffffff", bd=2, relief="flat", highlightbackground="#b0bec5", highlightcolor="#b0bec5", highlightthickness=2)
        self.main_frame.pack(padx=10, pady=10, fill="both", expand=True)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=0)

        Label(self.main_frame, text="Разделение субтитров по актерам", font=("Arial", 12, "bold"), bg="#ffffff", fg="black").grid(row=0, column=0, columnspan=2, pady=10)
        Label(self.main_frame, text="Путь к .ass файлу:", font=("Arial", 10), bg="#ffffff", fg="black").grid(row=1, column=0, sticky="w", padx=10)
        self.file_entry = Entry(self.main_frame, textvariable=self.file_path_var, font=("Arial", 9), bg="#f5f5f5", fg="black", relief="sunken", borderwidth=1)
        self.file_entry.grid(row=2, column=0, sticky="we", padx=(10, 5), pady=5)
        self.file_entry.bind("<Button-3>", self.show_context_menu)
        self.file_entry.bind("<Control-c>", self.copy_text)
        self.file_entry.bind("<Control-v>", self.paste_text)
        self.file_entry.bind("<Control-a>", self.select_all_text)
        self.setup_drag_and_drop()
        Button(self.main_frame, text="Выбрать", font=("Arial", 9), bg="#4CAF50", fg="white", activebackground="#45a049", activeforeground="white", relief="raised", borderwidth=2, command=self.choose_file).grid(row=2, column=1, sticky="w", padx=(5, 10), pady=5)
        self.group_check = Checkbutton(self.main_frame, text="Распределять строки 'гуры/все'", variable=self.distribute_group_var, font=("Arial", 9), bg="#ffffff", fg="black")
        self.group_check.grid_forget()
        self.multiple_check = Checkbutton(self.main_frame, text="Распределять множественные роли", variable=self.distribute_multiple_var, font=("Arial", 9), bg="#ffffff", fg="black")
        self.multiple_check.grid_forget()
        self.signs_check = Checkbutton(self.main_frame, text="Сохранять надписи в .ass", variable=self.save_signs_ass_var, font=("Arial", 9), bg="#ffffff", fg="black")
        self.signs_check.grid_forget()
        Label(self.main_frame, text="Формат сохранения:", font=("Arial", 10), bg="#ffffff", fg="black").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        format_menu = Combobox(self.main_frame, textvariable=self.format_var, values=["ass", "srt"], width=20, font=("Arial", 9), state="readonly")
        format_menu.grid(row=7, column=0, sticky="w", padx=10, pady=5)

        # Фрейм для кнопок
        button_frame = Frame(self.root, bg="#eceff1")
        button_frame.pack(side="bottom", fill="x", pady=5)
        close_button = Button(button_frame, text="Закрыть", font=("Arial", 10, "bold"), bg="#d32f2f", fg="white", activebackground="#c62828", activeforeground="white", relief="raised", borderwidth=2, command=self.close_app)
        close_button.pack(side="right", padx=5, ipadx=10)
        start_button = Button(button_frame, text="Запустить", font=("Arial", 10, "bold"), bg="#0288d1", fg="white", activebackground="#0277bd", activeforeground="white", relief="raised", borderwidth=2, command=self.start_processing)
        start_button.pack(side="right", padx=5, ipadx=10)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        if self.show_update_var.get():
            logging.info("Запланировано открытие окна 'Информация об обновлениях'")
            self.root.after(100, self.show_update_info)
        else:
            logging.info("Окно 'Информация об обновлениях' не будет показано")

    def setup_drag_and_drop(self):
        """Настраивает поддержку drag-and-drop для Entry."""
        if TkinterDnD is None:
            logging.warning("Drag-and-drop не поддерживается: tkinterdnd2 не установлена")
            messagebox.showwarning("Предупреждение", "Функция drag-and-drop недоступна. Установите библиотеку tkinterdnd2.")
            return
        try:
            self.file_entry.drop_target_register(DND_FILES)
            self.file_entry.dnd_bind('<<Drop>>', self.handle_drop)
            logging.info("Drag-and-drop настроен для поля ввода")
        except Exception as e:
            logging.error(f"Ошибка настройки drag-and-drop: {e}")
            messagebox.showerror("Ошибка", f"Не удалось настроить drag-and-drop: {e}")

    def handle_drop(self, event):
        """Обрабатывает событие drop для перетаскивания файла."""
        try:
            file_path = event.data
            if file_path.startswith('{') and file_path.endswith('}'):
                file_path = file_path[1:-1]
            file_path = file_path.strip()
            if not file_path:
                logging.error("Путь к файлу пустой")
                messagebox.showerror("Ошибка", "Перетаскиваемый файл не распознан.")
                return
            if not file_path.lower().endswith('.ass'):
                logging.error(f"Недопустимое расширение файла: {file_path}")
                messagebox.showerror("Ошибка", "Файл должен иметь расширение .ass.")
                return
            if not os.path.isfile(file_path):
                logging.error(f"Файл не существует: {file_path}")
                messagebox.showerror("Ошибка", "Указанный файл не существует.")
                return
            self.file_path_var.set(file_path)
            self.headers, self.styles, self.events = parse_ass_file(file_path)
            if self.headers is not None and self.styles is not None and self.events is not None:
                self.actors, self.group_lines, self.multiple_actor_lines, self.excluded_actor_lines, self.sign_lines, has_group_lines, has_multiple_actors, has_excluded_actors, has_sign_lines, self.all_actors = split_by_actor(self.events)
                self.show_group_option.set(has_group_lines)
                self.show_multiple_option.set(has_multiple_actors or has_excluded_actors)
                self.show_signs_option.set(has_sign_lines)
                self.group_check.grid_forget()
                self.multiple_check.grid_forget()
                self.signs_check.grid_forget()
                window_height = 260
                if has_group_lines:
                    self.group_check.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=2)
                    window_height += 25
                if has_multiple_actors or has_excluded_actors:
                    self.multiple_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=10, pady=2)
                    window_height += 25
                if has_sign_lines:
                    self.signs_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=2)
                    window_height += 25
                self.root.geometry(f"450x{window_height}")
                logging.info(f"Файл перетащен и обработан: {file_path}, окно установлено в размер 450x{window_height}")
        except Exception as e:
            logging.error(f"Ошибка при обработке перетаскивания файла: {e}")
            messagebox.showerror("Ошибка", f"Не удалось обработать перетаскиваемый файл: {e}")

    def load_settings(self):
        """Загружает настройки из settings.json."""
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                show_update = settings.get('show_update', True)
                logging.info(f"Настройки загружены: show_update={show_update}")
                return show_update
        except FileNotFoundError:
            logging.info("Файл настроек не найден, используется show_update=True")
            return True
        except json.JSONDecodeError:
            logging.error("Ошибка декодирования JSON, используется show_update=True")
            return True
        except Exception as e:
            logging.error(f"Ошибка загрузки настроек: {e}")
            return True

    def save_settings(self):
        """Сохраняет настройки в settings.json."""
        try:
            settings = {'show_update': self.show_update_var.get()}
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            logging.info(f"Настройки сохранены: {settings}")
        except PermissionError:
            logging.error(f"Ошибка: Нет прав для записи в {self.settings_file}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: нет прав для записи.")
        except Exception as e:
            logging.error(f"Ошибка сохранения настроек: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: {e}")

    def show_update_info(self):
        logging.info("Открытие окна 'Информация об обновлениях'")
        update_window = Toplevel(self.root)
        update_window.title("Информация об обновлениях")
        update_window.geometry("450x350")
        update_window.configure(bg="#eceff1")
        update_window.transient(self.root)
        update_window.grab_set()

        try:
            update_window.iconbitmap(resource_path("favicon.ico"))
        except Exception as e:
            logging.warning(f"Не удалось установить иконку: {e}")

        update_frame = Frame(update_window, bg="#ffffff", bd=2, relief="flat", highlightbackground="#b0bec5", highlightcolor="#b0bec5", highlightthickness=2)
        update_frame.pack(padx=10, pady=10, fill="both", expand=True)

        text_frame = Frame(update_frame, bg="#ffffff")
        text_frame.pack(fill="both", expand=True, padx=5, pady=5)

        update_text = Text(text_frame, height=15, width=50, wrap="word", font=("Arial", 9), bg="#f5f5f5", fg="black", relief="sunken", borderwidth=1)
        scrollbar = Scrollbar(text_frame, orient="vertical", command=update_text.yview)
        update_text.configure(yscrollcommand=scrollbar.set)

        update_text.tag_configure("bold", font=("Arial", 9, "bold"))
        update_text.tag_configure("header", font=("Arial", 12, "bold"))

        update_content = """Информация об обновлениях версии 1.1.0

    Изменения по сравнению с версией 1.0.2:

    Новые функции
    - Добавлен: полноценный графический интерфейс (GUI) с полем ввода, выбором формата и чекбоксами
    - Реализовано: контекстное меню и горячие клавиши (Ctrl+C, Ctrl+V, Ctrl+A)
    - Добавлено: меню приложения с разделами "Файл" и "Помощь"
    - Улучшена: обработка сложных случаев: строк "гуры/все", множественных ролей и исключений
    - Добавлены: информационные окна: "О программе", "Руководство", "Сохранение завершено"
    - Реализована: система настроек с возможностью отключения уведомлений об обновлениях
    - Добавлено: окно "Информация об обновлениях"
    - Добавлена: поддержка перетаскивания .ass файлов в поле ввода
    - Реализовано: автоматическая проверка расширения перетаскиваемых файлов
    - Добавлен: чекбокс "Сохранять надписи в .ass" для сохранения строк с метками "НАДПИСЬ", "Надпись", "надпись", "НАДПИСИ", "Надписи", "надписи", "SIGNS", "Signs", "signs", "SIGN", "Sign", "sign", "TEXT", "Text", "text", "ТЕКСТ", "Текст", "текст" в .ass файл
    - Реализовано: автоматическое определение строк с надписями при парсинге файла
    - Добавлено: сохранение надписей в отдельный .ass файл при выборе формата .srt

    Улучшения
    - Значительно улучшено: логирование и обработка ошибок
    - Все окна: теперь автоматически центрируются на экране
    - Реализована: безопасная обработка имен файлов с недопустимыми символами
    - Добавлено: динамическое изменение размера основного окна в зависимости от контента
    - Улучшена: работа с путями к файлам для кросс-платформенной совместимости
    - Добавлена: поддержка упаковки в exe с правильной обработкой ресурсов
    - Улучшена: обработка путей файлов при перетаскивании
    - Добавлена: поддержка кроссплатформенного drag-and-drop
    - Улучшена: обработка строк с надписями для совместимости с вшиванием в видео
    - Добавлено: динамическое отображение чекбокса для надписей

    Исправления
    - Исправлена: обработка пустых имен актеров
    - Улучшена: работа с Unicode для русских символов
    - Исправлены: различные мелкие баги в обработке субтитров
    - Улучшена: совместимость с разными форматами ASS файлов
    - Исправлена: обработка некорректных путей при перетаскивании
    - Улучшена: совместимость с различными форматами путей файлов
    - Исправлена: обработка некорректных меток надписей

    Поддержка
    Обратитесь к автору: https://t.me/itsptashka"""

        update_text.insert("end", update_content)
        logging.info("Текст обновлений успешно вставлен в виджет")
        update_text.tag_add("header", "1.0", "1.end")
        for section in ["Новые функции", "Улучшения", "Исправления", "Поддержка"]:
            start = update_text.search(section, "1.0", stopindex="end")
            if start:
                end = f"{start}+{len(section)}c"
                update_text.tag_add("bold", start, end)
                logging.debug(f"Применен жирный шрифт для раздела: {section}")
        for line in update_content.split('\n'):
            if line.startswith('- ') and ':' in line:
                start = update_text.search(line, "1.0", stopindex="end")
                if start:
                    colon_pos = update_text.search(":", start, stopindex=f"{start}+{len(line)}c")
                    if colon_pos:
                        end = f"{colon_pos}+1c"
                        update_text.tag_add("bold", start, end)
                        logging.debug(f"Применен жирный шрифт для строки: {line[:line.index(':')+1]}")

        update_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        update_text.configure(state="disabled")
        logging.info("Виджет Text и Scrollbar добавлены в окно")

        button_frame = Frame(update_frame, bg="#ffffff")
        button_frame.pack(fill="x", pady=5, side="bottom")

        dont_show_var = BooleanVar(value=not self.show_update_var.get())
        def save_and_close():
            self.show_update_var.set(not dont_show_var.get())
            self.save_settings()
            update_window.destroy()
            logging.info("Окно обновлений закрыто, настройки сохранены")

        Checkbutton(button_frame, text="Больше не показывать", variable=dont_show_var, font=("Arial", 9), bg="#ffffff", fg="black").pack(side="left", padx=5)
        close_button = Button(button_frame, text="Закрыть", font=("Arial", 9), bg="#d32f2f", fg="white", activebackground="#c62828", activeforeground="white", relief="raised", borderwidth=2, command=save_and_close)
        close_button.pack(side="right", padx=5, ipadx=10)

        update_window.update_idletasks()
        width = update_window.winfo_reqwidth()
        height = update_window.winfo_reqheight()
        x = (update_window.winfo_screenwidth() // 2) - (width // 2)
        y = (update_window.winfo_screenheight() // 2) - (height // 2)
        update_window.geometry(f"{width}x{height}+{x}+{y}")
        update_window.lift()
        logging.info(f"Окно 'Информация об обновлениях' центрировано: {width}x{height}+{x}+{y}")

    def show_settings(self):
        logging.info("Открытие окна 'Настройки'")
        settings_window = Toplevel(self.root)
        settings_window.title("Настройки")
        settings_window.geometry("350x150")
        settings_window.configure(bg="#eceff1")
        settings_window.transient(self.root)
        settings_window.grab_set()

        try:
            settings_window.iconbitmap(resource_path("favicon.ico"))
        except Exception as e:
            logging.warning(f"Не удалось установить иконку: {e}")

        settings_frame = Frame(settings_window, bg="#ffffff", bd=2, relief="flat", highlightbackground="#b0bec5", highlightcolor="#b0bec5", highlightthickness=2)
        settings_frame.pack(padx=10, pady=10, fill="both", expand=True)

        Label(settings_frame, text="Настройки программы", font=("Arial", 12, "bold"), bg="#ffffff", fg="black").pack(pady=5)
        Checkbutton(settings_frame, text="Показывать информацию об обновлениях при запуске", variable=self.show_update_var, font=("Arial", 9), bg="#ffffff", fg="black", command=lambda: [self.show_update_var.set(not self.show_update_var.get()), self.save_settings()]).pack(anchor="w", padx=10, pady=5)

        button_frame = Frame(settings_frame, bg="#ffffff")
        button_frame.pack(fill="x", pady=5, side="bottom")
        close_button = Button(button_frame, text="Закрыть", font=("Arial", 9), bg="#d32f2f", fg="white", activebackground="#c62828", activeforeground="white", relief="raised", borderwidth=2, command=settings_window.destroy)
        close_button.pack(side="right", padx=5, ipadx=10)

        settings_window.update_idletasks()
        width = settings_window.winfo_reqwidth()
        height = settings_window.winfo_reqheight()
        x = (settings_window.winfo_screenwidth() // 2) - (width // 2)
        y = (settings_window.winfo_screenheight() // 2) - (height // 2)
        settings_window.geometry(f"{width}x{height}+{x}+{y}")
        settings_window.lift()
        logging.info(f"Окно 'Настройки' центрировано: {width}x{height}+{x}+{y}")

    def show_context_menu(self, event):
        self.context_menu.post(event.x_root, event.y_root)
        logging.info("Контекстное меню отображено")

    def copy_text(self, event=None):
        try:
            focused_widget = self.root.focus_get()
            if isinstance(focused_widget, Entry):
                text = focused_widget.selection_get() if focused_widget.selection_present() else focused_widget.get()
                if text:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(text)
                    logging.info(f"Текст скопирован (Entry): {text}")
            elif isinstance(focused_widget, Text):
                if focused_widget.tag_ranges("sel"):
                    text = focused_widget.get("sel.first", "sel.last")
                else:
                    text = focused_widget.get("1.0", "end-1c")
                if text:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(text)
                    logging.info(f"Текст скопирован (Text): {text}")
        except Exception as e:
            logging.error(f"Ошибка копирования: {e}")

    def paste_text(self, event=None):
        try:
            focused_widget = self.root.focus_get()
            if isinstance(focused_widget, Entry):
                cursor_pos = focused_widget.index("insert")
                text = self.root.clipboard_get()
                focused_widget.insert(cursor_pos, text)
                logging.info(f"Текст вставлен (Entry): {text}")
            elif isinstance(focused_widget, Text):
                cursor_pos = focused_widget.index("insert")
                text = self.root.clipboard_get()
                focused_widget.insert(cursor_pos, text)
                logging.info(f"Текст вставлен (Text): {text}")
        except Exception as e:
            logging.error(f"Ошибка вставки: {e}")

    def select_all_text(self, event=None):
        try:
            focused_widget = self.root.focus_get()
            if isinstance(focused_widget, Entry):
                focused_widget.select_range(0, "end")
                focused_widget.icursor("end")
                logging.info("Текст выделен (Entry)")
            elif isinstance(focused_widget, Text):
                focused_widget.tag_add("sel", "1.0", "end-1c")
                focused_widget.mark_set("insert", "end")
                logging.info("Текст выделен (Text)")
        except Exception as e:
            logging.error(f"Ошибка выделения текста: {e}")

    def show_help(self):
        logging.info("Открытие окна 'Руководство'")
        help_window = Toplevel(self.root)
        help_window.title("Руководство")
        help_window.geometry("450x350")
        help_window.configure(bg="#eceff1")
        help_window.transient(self.root)
        help_window.grab_set()

        try:
            help_window.iconbitmap(resource_path("favicon.ico"))
        except Exception as e:
            logging.warning(f"Не удалось установить иконку: {e}")

        help_frame = Frame(help_window, bg="#ffffff", bd=2, relief="flat", highlightbackground="#b0bec5", highlightcolor="#b0bec5", highlightthickness=2)
        help_frame.pack(padx=10, pady=10, fill="both", expand=True)

        text_frame = Frame(help_frame, bg="#ffffff")
        text_frame.pack(fill="both", expand=True, padx=5, pady=5)

        help_text = Text(text_frame, height=15, width=50, wrap="word", font=("Arial", 9), bg="#f5f5f5", fg="black", relief="sunken", borderwidth=1)
        scrollbar = Scrollbar(text_frame, orient="vertical", command=help_text.yview)
        help_text.configure(yscrollcommand=scrollbar.set)

        help_text.tag_configure("bold", font=("Arial", 9, "bold"))
        help_text.tag_configure("header", font=("Arial", 12, "bold"))

        help_content = """Руководство по программе

    Назначение программы
    Программа Разделение субтитров предназначена для автоматического разделения субтитров в формате .ass на отдельные файлы для каждого актера. Поддерживаются форматы сохранения .ass и .srt.

    Подготовка файла субтитров
    Файл .ass должен содержать секцию [Events] с диалогами в формате:
    Dialogue: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text

    Где поле Name содержит имена актеров или метки. Поддерживаются следующие форматы:
    1. Один актер:          Name: Актер 1
    2. Несколько актеров:   Name: Актер 1, Актер 2  или  Name: Актер 1; Актер 2
    3. Исключения:          Name: !Актер 1, Актер 2 (диалог для всех КРОМЕ указанных актеров)
    4. Групповые диалоги:   Name: гуры  или  Name: все.
    5. Надписи:             Name: НАДПИСЬ, Надпись, надпись, НАДПИСИ, Надписи, надписи, SIGNS, Signs, signs, SIGN, Sign, sign, TEXT, Text, text, ТЕКСТ, Текст, "текст"

    Специальные символы:
    - Запятая или точка с запятой - разделители для нескольких актеров
    - Восклицательный знак в начале - индикатор исключения

    Как использовать
    1. Перетащите .ass файл в поле ввода или нажмите кнопку "Выбрать" либо используйте меню "Файл → Открыть файл"
    2. Настройте параметры:
       - "Распределять строки 'гуры/все'" - включает общие строки в файлы всех актеров
       - "Распределять множественные роли" - добавляет строки с несколькими актерами/исключениями
       - "Сохранять надписи в .ass" - сохраняет строки с метками надписей в .ass файл (даже при выборе .srt)
    3. Выберите формат сохранения: .ass или .srt
    4. Нажмите "Запустить" для обработки файла

    Результат
    - Файлы сохраняются в папку "Subtitles_by_Actor" рядом с исходным файлом
    - Имена файлов: "<Исходное_имя> - <Имя_актера> - (<Количество_строк>).ass/srt"
    - При выборе .srt и включенном "Сохранять надписи в .ass" создается дополнительный файл "<Исходное_имя> - Надписи.ass"
    - В .srt добавляется защитный субтитр для совместимости с REAPER

    Примеры обработки:
    1. "Актер 1, Актер 2" - строка попадет в файлы обоих актеров
    2. "!Актер 1" - строка попадет во все файлы, КРОМЕ файла Актер 1
    3. "гуры" или "все" - строка попадет во все файлы (если включена опция)
    4. "Надпись" - строка попадет в отдельный .ass файл (если включена опция)

    Преимущества
    - Экономия времени: автоматизация рутинной работы
    - Гибкость: поддержка сложных случаев с несколькими актерами и надписями
    - Простота: интуитивный интерфейс, горячие клавиши и перетаскивание файлов
    - Надежность: подробное логирование и обработка ошибок

    Поддержка
    При возникновении вопросов обратитесь к автору программы: https://t.me/itsptashka

    Разделение субтитров — ваш удобный инструмент для работы с субтитрами!"""

        help_text.insert("end", help_content)
        logging.info("Текст руководства успешно вставлен в виджет")
        help_text.tag_add("header", "1.0", "1.end")
        for section in ["Назначение программы", "Как использовать", "Результат", "Преимущества", "Поддержка"]:
            start = help_text.search(section, "1.0", stopindex="end")
            if start:
                end = f"{start}+{len(section)}c"
                help_text.tag_add("bold", start, end)
                logging.debug(f"Применен жирный шрифт для раздела: {section}")

        help_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        help_text.configure(state="disabled")
        logging.info("Виджет Text и Scrollbar добавлены в окно руководства")

        button_frame = Frame(help_frame, bg="#ffffff")
        button_frame.pack(fill="x", pady=5, side="bottom")
        close_button = Button(button_frame, text="Закрыть", font=("Arial", 9), bg="#d32f2f", fg="white", activebackground="#c62828", activeforeground="white", relief="raised", borderwidth=2, command=help_window.destroy)
        close_button.pack(side="right", padx=5, ipadx=10)

        help_window.update_idletasks()
        width = help_window.winfo_reqwidth()
        height = help_window.winfo_reqheight()
        x = (help_window.winfo_screenwidth() // 2) - (width // 2)
        y = (help_window.winfo_screenheight() // 2) - (height // 2)
        help_window.geometry(f"{width}x{height}+{x}+{y}")
        help_window.lift()
        logging.info(f"Окно 'Руководство' центрировано: {width}x{height}+{x}+{y}")

    def show_about(self):
        logging.info("Открытие окна 'О программе'")
        about_window = Toplevel(self.root)
        about_window.title("О программе")
        about_window.geometry("350x200")
        about_window.configure(bg="#eceff1")
        about_window.transient(self.root)
        about_window.grab_set()

        try:
            about_window.iconbitmap(resource_path("favicon.ico"))
        except Exception as e:
            logging.warning(f"Не удалось установить иконку: {e}")

        about_frame = Frame(about_window, bg="#ffffff", bd=2, relief="flat", highlightbackground="#b0bec5", highlightcolor="#b0bec5", highlightthickness=2)
        about_frame.pack(padx=10, pady=10, fill="both", expand=True)
        try:
            favicon = PhotoImage(file=resource_path("favicon.png"))
            Label(about_frame, image=favicon, bg="#ffffff").pack(anchor="nw", padx=5, pady=5)
            about_frame.image = favicon
        except Exception as e:
            logging.warning(f"Не удалось загрузить favicon.png: {e}")
        Label(about_frame, text="Распределитель субтитров", font=("Arial", 10, "bold"), bg="#ffffff", fg="black").pack(anchor="w", padx=5)
        Label(about_frame, text="Версия: 1.1.0", font=("Arial", 9), bg="#ffffff", fg="black").pack(anchor="w", padx=5)
        Label(about_frame, text="Издатель: Объект всеобщей ненависти", font=("Arial", 9), bg="#ffffff", fg="black").pack(anchor="w", padx=5)
        Label(about_frame, text="Разработчик: Феофилакт Птахен", font=("Arial", 9), bg="#ffffff", fg="black").pack(anchor="w", padx=5)
        Label(about_frame, text="Описание: Разделение субтитров по актерам", font=("Arial", 9), bg="#ffffff", fg="black").pack(anchor="w", padx=5)
        Label(about_frame, text="© Объект всеобщей ненависти. Все права не защищены.", font=("Arial", 9), bg="#ffffff", fg="black", wraplength=320, justify="left").pack(anchor="w", padx=5)
        button_frame = Frame(about_frame, bg="#ffffff")
        button_frame.pack(fill="x", pady=5, side="bottom")
        close_button = Button(button_frame, text="Закрыть", font=("Arial", 9), bg="#d32f2f", fg="white", activebackground="#c62828", activeforeground="white", relief="raised", borderwidth=2, command=about_window.destroy)
        close_button.pack(side="right", padx=5, ipadx=10)

        about_window.update_idletasks()
        width = about_window.winfo_reqwidth()
        height = about_window.winfo_reqheight()
        x = (about_window.winfo_screenwidth() // 2) - (width // 2)
        y = (about_window.winfo_screenheight() // 2) - (height // 2)
        about_window.geometry(f"{width}x{height}+{x}+{y}")
        about_window.lift()
        logging.info(f"Окно 'О программе' центрировано: {width}x{height}+{x}+{y}")

    def clear_field(self):
        self.file_path_var.set(os.path.expanduser("~/Desktop"))
        self.show_group_option.set(False)
        self.show_multiple_option.set(False)
        self.show_signs_option.set(False)
        self.group_check.grid_forget()
        self.multiple_check.grid_forget()
        self.signs_check.grid_forget()
        self.actors = None
        self.group_lines = None
        self.multiple_actor_lines = None
        self.excluded_actor_lines = None
        self.sign_lines = None
        self.all_actors = None
        self.root.geometry("450x300")
        logging.info("Поле ввода и чекбоксы очищены")

    def on_closing(self):
        try:
            keyboard.unhook_all()
            self.root.destroy()
            logging.info("Программа закрыта")
            sys.exit(0)
        except Exception as e:
            logging.error(f"Ошибка при закрытии программы: {e}")
            sys.exit(1)

    def close_app(self):
        logging.info("Пользователь закрыл программу")
        self.on_closing()

    def choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("ASS files", "*.ass"), ("All files", "*.*")], initialdir=os.path.expanduser("~/Desktop"))
        if file_path:
            self.file_path_var.set(file_path)
            self.headers, self.styles, self.events = parse_ass_file(file_path)
            if self.headers is not None and self.styles is not None and self.events is not None:
                self.actors, self.group_lines, self.multiple_actor_lines, self.excluded_actor_lines, self.sign_lines, has_group_lines, has_multiple_actors, has_excluded_actors, has_sign_lines, self.all_actors = split_by_actor(self.events)
                self.show_group_option.set(has_group_lines)
                self.show_multiple_option.set(has_multiple_actors or has_excluded_actors)
                self.show_signs_option.set(has_sign_lines)
                self.group_check.grid_forget()
                self.multiple_check.grid_forget()
                self.signs_check.grid_forget()
                window_height = 260
                if has_group_lines:
                    self.group_check.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=2)
                    window_height += 25
                if has_multiple_actors or has_excluded_actors:
                    self.multiple_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=10, pady=2)
                    window_height += 25
                if has_sign_lines:
                    self.signs_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=2)
                    window_height += 25
                self.root.geometry(f"450x{window_height}")
                logging.info(f"Окно установлено в размер 450x{window_height}")
        else:
            self.clear_field()

    def start_processing(self):
        file_path = self.file_path_var.get()
        if not file_path or file_path == os.path.expanduser("~/Desktop"):
            messagebox.showerror("Ошибка", "Укажите путь к .ass файлу.")
            return
        if not os.path.isfile(file_path):
            messagebox.showerror("Ошибка", "Указанный .ass файл не существует.")
            return
        if not file_path.lower().endswith('.ass'):
            messagebox.showerror("Ошибка", "Файл должен иметь расширение .ass.")
            return
        if self.actors is None:
            messagebox.showerror("Ошибка", "Сначала выберите .ass файл.")
            return
        original_filename = os.path.splitext(os.path.basename(file_path))[0]
        output_dir = os.path.join(os.path.dirname(file_path), 'Subtitles_by_Actor')
        export_format = self.format_var.get()
        distribute_group = bool(self.distribute_group_var.get()) if self.show_group_option.get() else False
        distribute_multiple = bool(self.distribute_multiple_var.get()) if self.show_multiple_option.get() else False
        save_signs_ass = bool(self.save_signs_ass_var.get()) if self.show_signs_option.get() else False
        logging.info(f"Запуск обработки: файл {file_path}, формат {export_format}, папка {output_dir}, save_signs_ass={save_signs_ass}")
        save_actor_files(self.headers, self.styles, self.actors, self.group_lines, self.multiple_actor_lines, self.excluded_actor_lines, self.sign_lines, output_dir, original_filename, export_format, distribute_group, distribute_multiple, save_signs_ass, self.all_actors)
        self.show_completion_dialog(output_dir)

    def show_completion_dialog(self, output_dir):
        logging.info("Открытие окна 'Сохранение завершено'")
        dialog = Toplevel(self.root)
        dialog.title("Сохранение завершено")
        dialog.geometry("400x250")
        dialog.configure(bg="#eceff1")
        try:
            dialog.iconbitmap(resource_path("favicon.ico"))
        except Exception as e:
            logging.warning(f"Не удалось установить иконку: {e}")
        dialog_frame = Frame(dialog, bg="#ffffff", bd=2, relief="flat", highlightbackground="#b0bec5", highlightcolor="#b0bec5", highlightthickness=2)
        dialog_frame.pack(padx=10, pady=10, fill="both", expand=True)
        Label(dialog_frame, text="Сохранение завершено", font=("Arial", 12, "bold"), bg="#ffffff", fg="black").pack(pady=5)
        Label(dialog_frame, text="Субтитры успешно сохранены:", font=("Arial", 10), bg="#ffffff", fg="black").pack(pady=5)
        path_entry = Text(dialog_frame, height=3, width=40, wrap="word", font=("Arial", 9), bg="#f5f5f5", fg="black", relief="sunken", borderwidth=1)
        path_entry.insert("end", output_dir)
        path_entry.bind("<Key>", lambda e: "break")
        path_entry.pack(pady=5)
        path_entry.bind("<Button-3>", self.show_context_menu)
        path_entry.bind("<Control-c>", self.copy_text)
        path_entry.bind("<Control-v>", self.paste_text)
        path_entry.bind("<Control-a>", self.select_all_text)
        copy_button = Button(dialog_frame, text="Копировать путь", font=("Arial", 9), bg="#4CAF50", fg="white", activebackground="#45a049", activeforeground="white", relief="raised", borderwidth=2, command=lambda: [self.root.clipboard_clear(), self.root.clipboard_append(output_dir), logging.info(f"Путь скопирован: {output_dir}")])
        copy_button.pack(pady=5)
        button_frame = Frame(dialog_frame, bg="#ffffff")
        button_frame.pack(fill="x", pady=5, side="bottom")
        close_button = Button(button_frame, text="Закрыть", font=("Arial", 9), bg="#d32f2f", fg="white", activebackground="#c62828", activeforeground="white", relief="raised", borderwidth=2, command=dialog.destroy)
        close_button.pack(side="right", padx=5, ipadx=10)
        new_file_button = Button(button_frame, text="Новый файл", font=("Arial", 9), bg="#0288d1", fg="white", activebackground="#0277bd", activeforeground="white", relief="raised", borderwidth=2, command=lambda: [dialog.destroy(), self.choose_file()])
        new_file_button.pack(side="right", padx=5, ipadx=10)

        dialog.update_idletasks()
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.lift()
        logging.info(f"Окно 'Сохранение завершено' центрировано: {width}x{height}+{x}+{y}")

def parse_ass_file(file_path):
    logging.info(f"Начало парсинга файла: {file_path}")
    headers = []
    styles = []
    events = []
    current_section = None
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as file:
            for line_number, line in enumerate(file, 1):
                line = line.strip()
                if not line:
                    continue
                if line.startswith('['):
                    current_section = line
                    continue
                if current_section == '[Script Info]' or current_section == '[V4+ Styles]':
                    if line:
                        if current_section == '[Script Info]':
                            headers.append(line)
                        elif current_section == '[V4+ Styles]':
                            styles.append(line)
                elif current_section == '[Events]':
                    if line.startswith('Dialogue:'):
                        events.append(line)
        if not events:
            logging.warning("Не найдено событий в секции [Events]")
            messagebox.showerror("Ошибка", "В файле не найдено строк Dialogue.")
            return None, None, None
        logging.info(f"Успешно распарсено: {len(events)} событий")
        return headers, styles, events
    except Exception as e:
        logging.error(f"Ошибка при парсинге файла {file_path}: {e}")
        messagebox.showerror("Ошибка", f"Не удалось распарсить файл {file_path}: {e}")
        return None, None, None

def split_by_actor(events):
    logging.info("Начало разделения событий по актерам")
    actors = {}
    group_lines = []
    multiple_actor_lines = []
    excluded_actor_groups = {}
    sign_lines = []  # Список для строк с надписями
    has_group_lines = False
    has_multiple_actors = False
    has_excluded_actors = False
    has_sign_lines = False
    all_actors = set()

    sign_variants = {'НАДПИСЬ', 'Надпись', 'надпись', 'НАДПИСИ', 'Надписи', 'надписи', 'ТЕКСТ', 'Текст', 'текст', 'SIGN', 'Sign', 'sign', 'SIGNS', 'Signs', 'signs', 'TEXT', 'Text', 'text'}  # Варианты меток надписей

    for event in events:
        parts = event.split(',', 9)
        if len(parts) < 10:
            logging.warning(f"Пропущена некорректная строка: {event}")
            continue
        actor_field = parts[4].strip()
        if actor_field.lower() in ['гуры', 'все'] or actor_field in sign_variants:
            continue
        if actor_field.startswith('!'):
            actors_list = [a.strip() for a in actor_field[1:].replace(';', ',').split(',') if a.strip()]
            for actor in actors_list:
                all_actors.add(actor)
        else:
            actors_list = [a.strip() for a in actor_field.replace(';', ',').replace('{', '').replace('}', '').split(',') if a.strip()]
            for actor in actors_list:
                all_actors.add(actor)

    for event in events:
        try:
            parts = event.split(',', 9)
            if len(parts) < 10:
                logging.warning(f"Пропущена некорректная строка: {event}")
                continue
            actor_field = parts[4].strip()
            if actor_field.lower() in ['гуры', 'все']:
                group_lines.append(event)
                has_group_lines = True
                logging.debug(f"Найдена строка 'гуры/все': {event}")
                continue
            if actor_field in sign_variants:
                sign_lines.append(event)
                has_sign_lines = True
                logging.debug(f"Найдена строка с надписью '{actor_field}': {event}")
                continue
            if actor_field.startswith('!'):
                has_excluded_actors = True
                actor_field = actor_field[1:].replace('{', '').replace('}', '')
                excluded_actors = [a.strip() for a in actor_field.replace(';', ',').split(',') if a.strip()]
                if not excluded_actors:
                    excluded_actors = ["unknown"]
                excluded_key = tuple(sorted(excluded_actors))
                if excluded_key not in excluded_actor_groups:
                    excluded_actor_groups[excluded_key] = []
                excluded_actor_groups[excluded_key].append(event)
                logging.debug(f"Добавлена строка с исключениями {excluded_actors}: {event}")
                continue
            actor_field = actor_field.replace('{', '').replace('}', '')
            actors_list = [a.strip() for a in actor_field.replace(';', ',').split(',') if a.strip()]
            if not actors_list:
                actors_list = ["unknown"]
            if len(actors_list) > 1:
                has_multiple_actors = True
                multiple_actor_lines.append((event, actors_list))
                logging.debug(f"Найдена множественная роль {actors_list}: {event}")
                continue
            for actor in actors_list:
                if actor not in actors:
                    actors[actor] = []
                actors[actor].append(event)
                logging.debug(f"Добавлено событие для актера {actor}: {event}")
        except Exception as e:
            logging.error(f"Ошибка при обработке строки: {event}, ошибка: {e}")
            continue
    if not actors and not group_lines and not multiple_actor_lines and not excluded_actor_groups and not sign_lines:
        logging.warning("Не найдено актеров, событий или надписей")
        messagebox.showerror("Ошибка", "Не найдено актеров, событий или надписей.")
        return None, None, None, None, None, False, False, False, False, None
    logging.info(f"Найдено актеров: {len(actors)}, строк 'гуры/все': {len(group_lines)}, строк с множественными ролями: {len(multiple_actor_lines)}, групп исключений: {len(excluded_actor_groups)}, строк с надписями: {len(sign_lines)}")
    return actors, group_lines, multiple_actor_lines, excluded_actor_groups, sign_lines, has_group_lines, has_multiple_actors, has_excluded_actors, has_sign_lines, all_actors

def format_srt_time(ass_time):
    try:
        parts = ass_time.split(':')
        if len(parts) != 3:
            return ass_time.replace('.', ',')
        hours = parts[0].zfill(2)
        minutes = parts[1]
        seconds, centiseconds = parts[2].split('.')
        milliseconds = centiseconds.ljust(3, '0')[:3]
        return f"{hours}:{minutes}:{seconds},{milliseconds}"
    except Exception as e:
        logging.error(f"Ошибка при преобразовании времени {ass_time}: {e}")
        return ass_time.replace('.', ',')

def save_ass_file(headers, styles, events, output_file):
    logging.info(f"Попытка сохранения .ass файла: {output_file}")
    try:
        with open(output_file, 'w', encoding='utf-8') as file:
            file.write('[Script Info]\n')
            for header in headers:
                file.write(header + '\n')
            file.write('\n[V4+ Styles]\n')
            file.write('Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding\n')
            for style in styles:
                file.write(style + '\n')
            file.write('\n[Events]\n')
            file.write('Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n')
            for event in events:
                file.write(event + '\n')
        logging.info(f"Успешно сохранен файл: {output_file}")
    except Exception as e:
        logging.error(f"Ошибка при сохранении .ass файла {output_file}: {e}")
        raise

def save_srt_file(events, output_file):
    logging.info(f"Попытка сохранения .srt файла: {output_file}")
    try:
        with open(output_file, 'w', encoding='utf-8') as file:
            if events:
                first_event = events[0].split(',', 9)
                if len(first_event) >= 3:
                    start_time = format_srt_time(first_event[1])
                else:
                    start_time = "00:00:00,000"
            else:
                start_time = "00:00:00,000"
            file.write("1\n00:00:00,000 --> " + start_time + "\n(Защита от удаления первого саба REAPER'ом!)\n\n")
            for index, event in enumerate(events, 2):
                parts = event.split(',', 9)
                if len(parts) < 10:
                    logging.warning(f"Пропущена некорректная строка: {event}")
                    continue
                start_time = format_srt_time(parts[1])
                end_time = format_srt_time(parts[2])
                text = parts[9].replace('\\N', ' ').replace('{i}', '')
                text = ''.join(c for c in text if c not in '{}')
                file.write(f"{index}\n{start_time} --> {end_time}\n{text}\n\n")
        logging.info(f"Успешно сохранен файл: {output_file}")
    except Exception as e:
        logging.error(f"Ошибка при сохранении .srt файла {output_file}: {e}")
        raise

def save_actor_files(headers, styles, actors, group_lines, multiple_actor_lines, excluded_actor_groups, sign_lines, output_dir, original_filename, export_format, distribute_group, distribute_multiple, save_signs_ass, all_actors):
    logging.info(f"Проверка прав доступа для папки: {output_dir}")
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            logging.info(f"Создана папка: {output_dir}")
        else:
            test_file = os.path.join(output_dir, "test_write.txt")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            logging.info(f"Права доступа для записи в {output_dir} подтверждены")
    except Exception as e:
        logging.error(f"Ошибка доступа к папке {output_dir}: {e}")
        messagebox.showerror("Ошибка", f"Нет прав доступа для записи в папку {output_dir}: {e}")
        return

    if not actors and not group_lines and not multiple_actor_lines and not excluded_actor_groups and not sign_lines:
        logging.error("Нет актеров, событий или надписей для сохранения файлов")
        messagebox.showerror("Ошибка", "Не найдено актеров, событий или надписей в файле субтитров.")
        return

    for actor, events in actors.items():
        safe_actor_name = re.sub(r'[<>:"/\\|?*]', '', actor).strip()
        if distribute_group and group_lines:
            events = events + group_lines
            logging.debug(f"Добавлены строки 'гуры/все' для актера {actor}: {len(group_lines)} строк")
        if distribute_multiple and multiple_actor_lines:
            for event, actors_list in multiple_actor_lines:
                if actor in actors_list:
                    events.append(event)
                    logging.debug(f"Добавлена строка с множественными ролями для актера {actor}: {event}")
        if distribute_multiple and excluded_actor_groups:
            for excluded_actors, excl_events in excluded_actor_groups.items():
                if actor not in excluded_actors:
                    events.extend(excl_events)
                    logging.debug(f"Добавлены строки с исключениями для актера {actor} (не в {excluded_actors}): {len(excl_events)} строк")
        line_count = len(events)
        output_file = os.path.join(output_dir, f"{original_filename} - {safe_actor_name} - ({line_count}).{export_format}")
        logging.info(f"Сохранение файла для актера {actor}: {output_file} (строк: {line_count})")
        try:
            if export_format == 'srt':
                save_srt_file(events, output_file)
            elif export_format == 'ass':
                save_ass_file(headers, styles, events, output_file)
            else:
                logging.error(f"Недопустимый формат: {export_format}")
                messagebox.showerror("Ошибка", f"Недопустимый формат: {export_format}")
                return
        except Exception as e:
            logging.error(f"Ошибка при сохранении файла {output_file}: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл {output_file}: {e}")

    if not distribute_group and group_lines:
        safe_actor_name = "Гуры"
        line_count = len(group_lines)
        output_file = os.path.join(output_dir, f"{original_filename} - {safe_actor_name} - ({line_count}).{export_format}")
        logging.info(f"Сохранение файла для гуры/все: {output_file} (строк: {line_count})")
        try:
            if export_format == 'srt':
                save_srt_file(group_lines, output_file)
            elif export_format == 'ass':
                save_ass_file(headers, styles, group_lines, output_file)
        except Exception as e:
            logging.error(f"Ошибка при сохранении файла {output_file}: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл {output_file}: {e}")

    if not distribute_multiple and multiple_actor_lines:
        for event, actors_list in multiple_actor_lines:
            safe_actor_name = " ".join(re.sub(r'[<>:"/\\|?*]', '', actor).strip() for actor in actors_list)
            output_file = os.path.join(output_dir, f"{original_filename} - {safe_actor_name} - (1).{export_format}")
            logging.info(f"Сохранение файла для множественных ролей {actors_list}: {output_file}")
            try:
                if export_format == 'srt':
                    save_srt_file([event], output_file)
                elif export_format == 'ass':
                    save_ass_file(headers, styles, [event], output_file)
            except Exception as e:
                logging.error(f"Ошибка при сохранении файла {output_file}: {e}")
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл {output_file}: {e}")

    if not distribute_multiple and excluded_actor_groups:
        for excluded_actors, events in excluded_actor_groups.items():
            safe_actor_name = "Без " + " ".join(re.sub(r'[<>:"/\\|?*]', '', actor).strip() for actor in excluded_actors)
            line_count = len(events)
            output_file = os.path.join(output_dir, f"{original_filename} - {safe_actor_name} - ({line_count}).{export_format}")
            logging.info(f"Сохранение файла для исключённых актёров {excluded_actors}: {output_file} (строк: {line_count})")
            try:
                if export_format == 'srt':
                    save_srt_file(events, output_file)
                elif export_format == 'ass':
                    save_ass_file(headers, styles, events, output_file)
            except Exception as e:
                logging.error(f"Ошибка при сохранении файла {output_file}: {e}")
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл {output_file}: {e}")

    if save_signs_ass and sign_lines:
        safe_actor_name = "Надписи"
        line_count = len(sign_lines)
        output_file = os.path.join(output_dir, f"{original_filename} - {safe_actor_name} - ({line_count}).ass")
        logging.info(f"Сохранение файла для надписей: {output_file} (строк: {line_count})")
        try:
            save_ass_file(headers, styles, sign_lines, output_file)
        except Exception as e:
            logging.error(f"Ошибка при сохранении файла {output_file}: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл {output_file}: {e}")

def main():
    logging.info("Запуск программы")
    if TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        root = Tk()
    app = SubtitleSplitterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()