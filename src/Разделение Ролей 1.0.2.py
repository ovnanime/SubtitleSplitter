import os
import sys
import subprocess
import logging
from tkinter import Tk, filedialog, messagebox

# Настройка логирования (только в консоль)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Проверяем и устанавливаем tkinter
try:
    from tkinter import Tk, filedialog, messagebox
    logging.info("tkinter успешно импортирован")
except ImportError:
    logging.error("tkinter отсутствует, пытаюсь установить...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "tk"])
        from tkinter import Tk, filedialog, messagebox
        logging.info("tkinter успешно установлен и импортирован")
    except Exception as e:
        logging.error(f"Ошибка при установке tkinter: {e}")
        print("Попробуйте установить вручную: 'pip install tk' или 'sudo apt-get install python3-tk' (для Linux).")
        sys.exit(1)

def parse_ass_file(file_path):
    """Читает .ass файл и возвращает заголовки, стили и события."""
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
                    logging.debug(f"Смена секции на: {current_section}")
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
            messagebox.showerror("Ошибка", "В файле не найдено строк Dialogue в секции [Events].")
            return None, None, None
        logging.info(f"Успешно распарсено: {len(events)} событий, {len(headers)} заголовков, {len(styles)} стилей")
        return headers, styles, events
    except Exception as e:
        logging.error(f"Ошибка при парсинге файла {file_path} на строке {line_number}: {e}")
        messagebox.showerror("Ошибка", f"Не удалось распарсить файл {file_path}: {e}")
        return None, None, None

def split_by_actor(events):
    """Разделяет события по актерам, включая неизвестных (unknown)."""
    logging.info("Начало разделения событий по актерам")
    actors = {}
    for event in events:
        try:
            parts = event.split(',', 9)  # Dialogue: Layer,Start,End,Style,Name,MarginL,MarginR,MarginV,Effect,Text
            if len(parts) < 10:
                logging.warning(f"Пропущена некорректная строка: {event}")
                continue
            actor = parts[4].strip()
            if not actor:  # Если имя актера пустое, используем "unknown"
                actor = "unknown"
            if actor not in actors:
                actors[actor] = []
            actors[actor].append(event)
        except Exception as e:
            logging.error(f"Ошибка при обработке строки: {event}, ошибка: {e}")
            continue
    if not actors:
        logging.warning("Не найдено актеров с непустыми именами или неизвестных")
        messagebox.showerror("Ошибка", "Не найдено актеров или событий в файле субтитров.")
        return None
    logging.info(f"Найдено актеров: {len(actors)}")
    return actors

def format_srt_time(ass_time):
    """Преобразует время из формата .ass (0:00:00.00) в .srt (00:00:00,000)."""
    try:
        parts = ass_time.split(':')
        if len(parts) != 3:
            logging.warning(f"Некорректный формат времени: {ass_time}")
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
    """Сохраняет субтитры в формате .ass без защитного субтитра."""
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
    """Сохраняет субтитры в формате .srt с защитным субтитром, убирая форматирование."""
    logging.info(f"Попытка сохранения .srt файла: {output_file}")
    try:
        with open(output_file, 'w', encoding='utf-8') as file:
            # Извлекаем время начала первого события для защитного субтитра
            if events:
                first_event = events[0].split(',', 9)
                if len(first_event) >= 3:
                    start_time = format_srt_time(first_event[1])
                else:
                    start_time = "00:00:00,000"
            else:
                start_time = "00:00:00,000"

            # Добавляем защитный субтитр
            file.write("1\n00:00:00,000 --> " + start_time + "\n(Защита от удаления первого саба REAPER'ом!)\n\n")
            # Основные субтитры начинаются с индекса 2
            for index, event in enumerate(events, 2):
                parts = event.split(',', 9)
                if len(parts) < 10:
                    logging.warning(f"Пропущена некорректная строка: {event}")
                    continue
                start_time = format_srt_time(parts[1])
                end_time = format_srt_time(parts[2])
                text = parts[9].replace('\\N', ' ').replace('{i}', '')  # Убираем \N и {i}
                text = ''.join(c for c in text if c not in '{}')  # Удаляем все фигурные скобки и их содержимое
                file.write(f"{index}\n{start_time} --> {end_time}\n{text}\n\n")
        logging.info(f"Успешно сохранен файл: {output_file}")
    except Exception as e:
        logging.error(f"Ошибка при сохранении .srt файла {output_file}: {e}")
        raise

def save_actor_files(headers, styles, actors, output_dir, original_filename, export_format):
    """Сохраняет субтитры для каждого актера в выбранном формате."""
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

    if not actors:
        logging.error("Нет актеров для сохранения файлов")
        messagebox.showerror("Ошибка", "Не найдено актеров или событий в файле субтитров.")
        return

    for actor, events in actors.items():
        safe_actor_name = "".join(c for c in actor if c.isalnum() or c in ('-', ' ')).strip().replace(' ', '-')
        line_count = len(events)
        output_file = os.path.join(output_dir, f"{original_filename} - {safe_actor_name} - ({line_count}).{export_format}")
        logging.info(f"Сохранение файла для актера {actor}: {output_file} (формат: {export_format})")

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

def ask_convert_to_srt():
    """Создает стандартное окно Windows с вопросом о конвертации в .srt."""
    logging.info("Открытие окна вопроса о конвертации в .srt")
    try:
        result = messagebox.askyesno(
            title="Конвертация субтитров",
            message="Конвертировать субтитры в формат .srt?\n(Да - SRT, Нет - сохранить как ASS)"
        )
        if result:
            logging.info("Выбрано 'Да' (конвертировать в .srt)")
            return 'srt'
        else:
            logging.info("Выбрано 'Нет' (оставить в .ass)")
            return 'ass'
    except Exception as e:
        logging.error(f"Ошибка в окне вопроса о конвертации: {e}")
        messagebox.showerror("Ошибка", f"Не удалось открыть окно конвертации: {e}")
        return None

def main():
    logging.info("Запуск программы")
    root = Tk()
    # Устанавливаем иконку из файла favicon.ico в той же папке
    try:
        root.iconbitmap("favicon.ico")
    except Exception as e:
        logging.warning(f"Не удалось установить иконку favicon.ico: {e}")
    root.withdraw()
    try:
        logging.info("Открытие диалога выбора файла")
        file_path = filedialog.askopenfilename(filetypes=[("ASS files", "*.ass"), ("All files", "*.*")])
        
        if not file_path:
            logging.warning("Файл не выбран")
            messagebox.showinfo("Информация", "Файл не выбран.")
            return

        original_filename = os.path.splitext(os.path.basename(file_path))[0]
        output_dir = os.path.join(os.path.dirname(file_path), 'Subtitles_by_Actor')
        headers, styles, events = parse_ass_file(file_path)
        
        if headers is None or styles is None or events is None:
            logging.error("Парсинг файла не удался, прерывание")
            return

        actors = split_by_actor(events)
        if actors is None:
            logging.error("Разделение по актерам не удалось, прерывание")
            return

        export_format = ask_convert_to_srt()
        if export_format is None:
            logging.error("Формат не выбран, прерывание")
            return
        logging.info(f"Перед сохранением файлов: выбран формат {export_format}")
        save_actor_files(headers, styles, actors, output_dir, original_filename, export_format)
        messagebox.showinfo("Успех", f"Субтитры разделены и сохранены в папке: {output_dir}")
        logging.info("Программа успешно завершена")
    except Exception as e:
        logging.error(f"Ошибка в main: {e}")
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
    finally:
        root.destroy()
        logging.info("Закрытие главного окна tkinter")

if __name__ == "__main__":
    main()
