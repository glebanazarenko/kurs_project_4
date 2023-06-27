import os
import io
import argparse
import math
import fitz
import re
import ebooklib #для работы с ePub-файлами
from ebooklib import epub
import sqlite3
import matplotlib.pyplot as plt
from typing import List, Tuple
from PyPDF2 import PdfReader
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import docx2txt
from odf.opendocument import load
from odf.namespaces import OFFICENS
from odf import meta
from odf import text
from odf import teletype
import aspose.words as aw

def yes_no_indicator(value):
    if value == 0:
        return "Нет"
    elif value == 1:
        return "Да"
    else:
        return "Недопустимое значение"

def pretty_size(size_bytes: int) -> str:
    """
    Конвертирует размер файла в формат, понятный человеку.\n
    Аргументы:
    size_bytes -- размер файла в байтах\n
    Возвращает:
    Строку, содержащую размер файла в формате, понятном человеку.
    Например: "10.5 МБ", "678 КБ", "5 ГБ" и т.д.
    """
    units = ['байт', 'КБ', 'МБ', 'ГБ', 'ТБ', 'ПБ']
    if size_bytes < 1024:
        return f"{size_bytes:.0f} {units[0]}"
    exp = int(math.log(size_bytes, 1024))
    size = size_bytes / 1024 ** exp
    size_str = f"{size:.1f}"
    return f"{size_str} {units[exp]}"

class BookAnalyzer:
    def __init__(self, db_path: str, reset=False, convert_docx_to_pdf=False, convert_odt_to_pdf=False):
        self.db_path = db_path
        self.reset = reset
        self.convert_docx_to_pdf = convert_docx_to_pdf
        self.convert_odt_to_pdf = convert_odt_to_pdf
        self.init_database()

    def open_db(self):
        self.conn = sqlite3.connect(self.db_path)
        return self.conn.cursor()
    
    def close_db(self):
        self.conn.commit()
        self.conn.close()

    def init_database(self):
        # Создайте БД (если не существует) и определите таблицы для хранения метаданных книг и превью
        # Создаем соединение с БД, если файла нет, он будет создан
        conn = sqlite3.connect(self.db_path)

        # Создаем курсор для выполнения SQL-запросов
        cursor = conn.cursor()

        # Удаление таблицы в бд, если reset = True
        if self.reset:
            cursor.execute('''
                DROP TABLE IF EXISTS books
            ''')

        # Создаем таблицу, если она не существует
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                author TEXT,
                file_ext TEXT,
                file_path TEXT UNIQUE,
                file_size INTEGER,
                num_pages INTEGER,
                preview BLOB,
                metadata TEXT,
                favorite INTEGER
            )
        ''')

        # Сохраняем изменения и закрываем соединение
        conn.commit()
        conn.close()

    def update_book_data(self, file_path):
        # Создаем соединение с БД
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            if file_ext == '.pdf':
                with open(file_path, 'rb') as file:
                    reader = PdfReader(file)
                    metadata = reader.metadata
                    num_pages = len(reader.pages)
                    preview = self.__get_preview(file_path)

                    # Извлечение данных о названии и авторе
                    if metadata != None:
                        title = metadata.get('/Title')
                        author = metadata.get('/Author')
                    else:
                        title = None
                        author = None

            elif file_ext == '.epub':
                book = epub.read_epub(file_path)
                metadata = book.metadata
                num_pages = self.__count_generator_items(book.get_items_of_type(ebooklib.ITEM_DOCUMENT))
                preview = self.__get_epub_cover(book)

                # Извлечение данных о названии и авторе
                dc_metadata = metadata.get('http://purl.org/dc/elements/1.1/')
                title = dc_metadata['title'][0][0] if dc_metadata and 'title' in dc_metadata else None
                author = dc_metadata['creator'][0][0] if dc_metadata and 'creator' in dc_metadata else None
            
            elif file_ext == '.docx':
                if self.convert_docx_to_pdf:
                    metadata, num_pages, preview, title, author = self.__convert_to_pdf(file_path)
                else:
                    document = Document(file_path)
                    # Заголовок файла будет использоваться как название
                    title = document.core_properties.title or os.path.splitext(os.path.basename(file_path))[0]
                    # Информация об авторе
                    author = document.core_properties.author
                    # Количество страниц в docx файлах обычно не доступно
                    num_pages = self.__count_pages_docx(file_path)
                    # Метаданные из core_properties
                    metadata = {
                        'author': document.core_properties.author,
                        'title': document.core_properties.title,
                        'subject': document.core_properties.subject,
                        'keywords': document.core_properties.keywords,
                        'last_modified_by': document.core_properties.last_modified_by,
                        'created': document.core_properties.created,
                        'modified': document.core_properties.modified,
                        'category': document.core_properties.category,
                        'comments': document.core_properties.comments,
                        'content_status': document.core_properties.content_status,
                        'identifier': document.core_properties.identifier,
                        'language': document.core_properties.language,
                        'version': document.core_properties.version,
                        'last_printed': document.core_properties.last_printed,
                        'revision': document.core_properties.revision,
                    }
                    preview = self.__get_docx_preview(file_path)
            elif file_ext == '.odt':
                if self.convert_odt_to_pdf:
                    metadata, num_pages, preview, title, author = self.__convert_to_pdf(file_path)
                else:
                    # Извлечение метаданных из файла odt без преобразования в pdf
                    metadata = None
                    num_pages = None
                    title = None
                    author = None
                    preview = self.__get_odt_preview(file_path)
                

            # Если в метаданных нет названия, используем имя файла без расширения
            if not title:
                title = os.path.splitext(os.path.basename(file_path))[0]

            # Извлечение размера файла
            file_size = os.path.getsize(file_path)

            # Проверяем, есть ли уже книга в БД
            cursor.execute('SELECT * FROM books WHERE file_path = ?', (file_path,))
            row = cursor.fetchone()

            if row is None:
                # Если книги нет в БД, добавляем ее
                # Выполняем запрос к базе данных

                # Подготавливаем данные для вставки
                data = (file_path, title, author, file_size, str(metadata), num_pages, preview, file_ext, 0)

                cursor.execute("""
                    INSERT OR REPLACE INTO books (file_path, title, author, file_size, metadata, num_pages, preview, file_ext, favorite)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, data)
            else:
                # Если книга уже есть в БД, обновляем ее

                # Подготавливаем данные для вставки
                data = (title, author, file_size, str(metadata), num_pages, preview, file_ext, 0, file_path)

                cursor.execute('''
                    UPDATE books SET title = ?, author = ?, file_size = ?, metadata = ?, num_pages = ?, preview = ?, file_ext = ?, favorite = ?
                    WHERE file_path = ?
                ''', data)
        except Exception as e:
            print(f"Ошибка в работе с файлом {file_path}. Причина: {e}")
        
        # Сохраняем изменения и закрываем соединение
        conn.commit()
        conn.close()

    def __get_preview(self, book_path: str) -> bytes:
        # Возвращает изображение превью (скриншот 1-й страницы) книги в виде байтов
        doc = fitz.open(book_path)
        page = doc[0]  # Возьмем первую страницу

        # Рендерим страницу в изображение
        pix = page.get_pixmap()
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Конвертируем PIL Image в bytes
        byte_arr = io.BytesIO()
        image.save(byte_arr, format='PNG')
        return byte_arr.getvalue()
    
    def __get_epub_cover(self, book: ebooklib.epub.EpubBook) -> bytes:
        # Попытка извлечь обложку
        cover_item_id = None
        cover_metadata = book.get_metadata('OPF', 'cover')

        if cover_metadata:
            cover_item_id = cover_metadata[0][0]

        cover_item = book.get_item_with_id(cover_item_id) if cover_item_id else None

        # Если обложка не найдена, ищем первое изображение в книге
        if cover_item is None:
            for item in book.get_items_of_type(ebooklib.ITEM_IMAGE):
                cover_item = item
                break

        # Если изображение так и не найдено, возвращаем None
        if cover_item is None:
            return None

        # Преобразование обложки в байты
        cover_image = Image.open(io.BytesIO(cover_item.get_content()))
        byte_arr = io.BytesIO()
        cover_image.save(byte_arr, format='PNG')
        return byte_arr.getvalue()
    
    @staticmethod
    def __count_pages_docx(docx_file_path):
        text = docx2txt.process(docx_file_path)
        return text.count('\f')  # '\f' является символом подачи формы, представляющим разрывы страниц
    
    def __get_docx_preview(self, file_path: str) -> bytes:
        # Возвращает изображение превью (первые несколько параграфов) документа в виде байтов
        text = docx2txt.process(file_path)
        lines = text.split('\n')[:10]  # Берем первые 10 строк текста
        
        # Генерируем изображение из текста
        font = ImageFont.truetype('arial', 15)
        img = Image.new('RGB', (500, 200), color=(73, 109, 137))
        d = ImageDraw.Draw(img)
        for i, line in enumerate(lines):
            d.text((10, 10 + i*15), line, fill=(255, 255, 0), font=font)
        
        byte_arr = io.BytesIO()
        img.save(byte_arr, format='PNG')
        return byte_arr.getvalue()
    
    @staticmethod
    def extract_odt_metadata(file_path):
        doc = load(file_path)
        meta_info = doc.getElementsByType(meta.Meta)
        
        if meta_info:
            meta_info = teletype.extractText(meta_info[0])
            # Делаем предположение о том, как данные разделены, и разделяем строку
            meta_info_parts = re.split(r"(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z)", meta_info)
            software_used, title, author, username, *dates, edit_cycles, duration = meta_info_parts

            metadata = {
                'software_used': software_used,
                'title': title,
                'author': author,
                'username': username,
                'dates': dates,
                'edit_cycles': edit_cycles,
                'edit_duration': duration
            }
            
            return metadata
        else:
            return None

    @staticmethod
    def __get_odt_metadata(file_path):
        doc = load(file_path)
        meta = doc.meta

        metadata = {
            'title': teletype.extractText(meta.title) if meta.title else None,
            'description': teletype.extractText(meta.description) if meta.description else None,
            'subject': teletype.extractText(meta.subject) if meta.subject else None,
            'language': teletype.extractText(meta.language) if meta.language else None,
            'editing_cycles': teletype.extractText(meta.editingcycles) if meta.editingcycles else None,
            'editing_duration': teletype.extractText(meta.editingduration) if meta.editingduration else None,
            'date': teletype.extractText(meta.date) if meta.date else None,
            'creator': teletype.extractText(meta.creator) if meta.creator else None,
            'keyword': teletype.extractText(meta.keyword) if meta.keyword else None,
        }
        
        return metadata
    
    @staticmethod
    def __get_odt_preview(file_path):
        # Загружаем документ
        doc = load(file_path)
        # Извлекаем текст из каждого элемента 'P' и объединяем их с новыми строками
        all_text = "\n".join(teletype.extractText(p) for p in doc.getElementsByType(text.P))
        # Берем первые 10 строк текста
        lines = all_text.split('\n')[:10]
        
        # Генерируем изображение из текста
        font = ImageFont.truetype('arial', 15)
        img = Image.new('RGB', (500, 200), color=(73, 109, 137))
        d = ImageDraw.Draw(img)
        for i, line in enumerate(lines):
            d.text((10, 10 + i*15), line, fill=(255, 255, 0), font=font)
        
        # Преобразуем изображение в байты
        byte_arr = io.BytesIO()
        img.save(byte_arr, format='PNG')
        return byte_arr.getvalue()
    
    @staticmethod
    def __count_generator_items(generator):
            return sum(1 for _ in generator)
    
    def __convert_to_pdf(self, file_path):
        # создаем объект Document и загружаем файл
        doc = aw.Document(file_path)
        file_pdf = "Output.pdf"
        # сохраняем документ в формате pdf
        doc.save(file_pdf, aw.SaveFormat.PDF)
        reader = PdfReader(file_pdf)
        metadata = reader.metadata
        num_pages = len(reader.pages)
        preview = self.__get_preview(file_pdf)
        # Извлечение данных о названии и авторе
        try:
            title = metadata.get('/Title')
            author = metadata.get('/Author')
        except Exception as e:
            title = os.path.splitext(os.path.basename(file_pdf))[0]
            author = None
            
        os.remove(file_pdf)
        return metadata, num_pages, preview, title, author

    
    def __get_all_previews(self):
        # Создаем соединение с БД
        cursor = self.open_db()

        # Получаем все превью из БД
        cursor.execute('SELECT title, preview FROM books')
        previews = cursor.fetchall()

        # Закрываем соединение
        self.close_db()

        # Возвращаем список кортежей вида (название книги, превью)
        return previews
    
    def display_previews(self):
        # Получаем все превью
        previews = self.__get_all_previews()

        # Выводим превью каждой книги
        for title, preview in previews:
            # Конвертируем bytes в PIL Image
            image = Image.open(io.BytesIO(preview))

            # Выводим изображение
            plt.figure()
            plt.imshow(image)
            plt.title(title)
            plt.axis('off')
            plt.show()
    
    def process_directory(self, directory, file_types, exclude, max_depth = 5, current_depth=0, convert_odt_to_pdf=None, convert_docx_to_pdf=None):
        if convert_odt_to_pdf is not None:
            self.convert_odt_to_pdf = convert_odt_to_pdf
        if convert_docx_to_pdf is not None:
            self.convert_docx_to_pdf = convert_docx_to_pdf
        # Обработка каталога (рекурсивно), обновление информации о книгах в БД
        if current_depth > max_depth:
            return

        try:
            with os.scandir(directory) as entries:
                for entry in entries:
                    if entry.is_file() and any(entry.name.lower().endswith(ft) for ft in file_types):
                        # Если это файл и его тип в списке разрешенных типов файлов, обрабатываем его
                        self.update_book_data(str(entry.path))
                    elif entry.is_dir() and entry.name not in exclude:
                        # Если это каталог и его имя не в списке исключений, рекурсивно обрабатываем его
                        self.process_directory(entry.path, file_types, exclude, max_depth, current_depth + 1)
        except PermissionError:
            print(f"Permission denied for directory: {directory}")

    # ЗАПРОСЫ К БД

    def get_book_metadata(self, file_path):
        cursor = self.open_db()
        query = f"SELECT metadata FROM books WHERE file_path = '{file_path}'"

        cursor.execute(query)
        row = cursor.fetchone()

        self.close_db()

        # Возвращаем метаданные, если они найдены, иначе None
        return eval(row[0]) if row else None

    def update_book_favorite_status(self, file_path):
        try:
            cursor = self.open_db()

            # Получение текущего статуса избранного для книги
            query = f"SELECT favorite FROM books WHERE file_path = '{file_path}'"

            cursor.execute(query)
            rows = cursor.fetchall()

            if rows is None:
                raise ValueError("Книга не найдена")

            # Инвертирование статуса избранного
            new_status = 0 if rows[0][0] == 1 else 1

            # Подготавливаем данные для вставки
            data = (new_status, file_path)
            #data = (file_path,)

            # Обновление статуса избранного в базе данных
            cursor.execute('''
                UPDATE books SET favorite = ?
                WHERE file_path = ?
            ''', data)

            self.close_db()

        except Exception as e:
            print("Ошибка обновления статуса избранное у книги:", str(e))

    def get_all_books(self, only_favorites=False, limit=None, offset=None):
        cursor = self.open_db()
        query = "SELECT favorite, title, author, file_ext, file_path, file_size, num_pages, metadata FROM books"

        if only_favorites:
            query += " WHERE favorite = 1"

        if limit is not None and offset is not None:
            query += f" LIMIT {limit} OFFSET {offset}"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        # Преобразуем размер файла в человеко-читаемый формат
        rows = [(yes_no_indicator(favorite), title, author, file_ext, file_path, pretty_size(file_size), num_pages, metadata) for favorite, title, author, file_ext, file_path, file_size, num_pages, metadata in rows]

        return rows
    
    def get_book_preview(self, book_id):
        cursor = self.open_db()
        query = f"SELECT preview FROM books WHERE id = ?"

        cursor.execute(query, (book_id,))
        row = cursor.fetchone()

        self.close_db()

        return row[0] if row else None
    
    def get_book_preview_path(self, file_path):
        cursor = self.open_db()
        query = f"SELECT preview FROM books WHERE file_path = ?"

        cursor.execute(query, (file_path,))
        row = cursor.fetchone()

        self.close_db()

        return row[0] if row else None

    # Поиск книг по названию
    def search_books_by_title(self, title, only_favorites=False):
        cursor = self.open_db()
        query = f"SELECT favorite, title, author, num_pages, file_path FROM books WHERE title LIKE '%{title}%'"

        if only_favorites:
            query += " AND favorite = 1"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return [(yes_no_indicator(favorite), title, author, num_pages, file_path) for favorite, title, author, num_pages, file_path in rows]

    # Поиск книг по автору
    def search_books_by_author(self, author, only_favorites=False):
        cursor = self.open_db()
        query = f"SELECT favorite, title, author, num_pages, file_path FROM books WHERE author LIKE '%{author}%'"

        if only_favorites:
            query += " AND favorite = 1"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return [(yes_no_indicator(favorite), title, author, num_pages, file_path) for favorite, title, author, num_pages, file_path in rows]

    # Поиск книг по расширению файла
    def search_books_by_extension(self, file_ext, only_favorites=False):
        cursor = self.open_db()
        query = f"SELECT favorite, title, author, file_ext, file_path FROM books WHERE file_ext LIKE '%{file_ext}%'"

        if only_favorites:
            query += " AND favorite = 1"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return [(yes_no_indicator(favorite), title, author, file_ext, file_path) for favorite, title, author, file_ext, file_path in rows]

    # Получить самые большие книги
    def get_largest_books(self, limit=5, offset=0, only_favorites=False):
        cursor = self.open_db()
        query = f"SELECT favorite, title, author, file_size, file_path FROM books ORDER BY file_size DESC LIMIT {limit} OFFSET {offset}"

        if only_favorites:
            query = f"SELECT favorite, title, author, file_size, file_path FROM books WHERE favorite = 1 ORDER BY file_size DESC LIMIT {limit} OFFSET {offset}"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        # Преобразуем размер файла в человеко-читаемый формат
        return [(yes_no_indicator(favorite), title, author, pretty_size(file_size), file_path) for favorite, title, author, file_size, file_path in rows]

    # Получить книги с наибольшим количеством страниц
    def get_books_with_most_pages(self, limit=5, offset=0, only_favorites=False): 
        cursor = self.open_db()
        query = f"SELECT favorite, title, author, num_pages, file_path FROM books ORDER BY num_pages DESC LIMIT {limit} OFFSET {offset}"

        if only_favorites:
            query = f"SELECT favorite, title, author, num_pages, file_path FROM books WHERE favorite = 1 ORDER BY num_pages DESC LIMIT {limit} OFFSET {offset}"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return [(yes_no_indicator(favorite), title, author, num_pages, file_path) for favorite, title, author, num_pages, file_path in rows]

    # Получить книги, добавленные последними
    def get_recently_added_books(self, limit=5, offset=0, only_favorites=False):
        cursor = self.open_db()
        query = f"SELECT favorite, title, author, file_path FROM books ORDER BY id DESC LIMIT {limit} OFFSET {offset}"

        if only_favorites:
            query = f"SELECT favorite, title, author, file_path FROM books WHERE favorite = 1 ORDER BY id DESC LIMIT {limit} OFFSET {offset}"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return [(yes_no_indicator(favorite), title, author, file_path) for favorite, title, author, file_path in rows]

    # Получить книги без автора
    def get_books_without_author(self, only_favorites=False):
        cursor = self.open_db()
        query = f"SELECT favorite, title, num_pages, file_path FROM books WHERE author IS NULL"

        if only_favorites:
            query += " AND favorite = 1"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return [(yes_no_indicator(favorite), title, num_pages, file_path) for favorite, title, num_pages, file_path in rows]

    # Получить книги без метаданных
    def get_books_without_metadata(self, only_favorites=False):
        cursor = self.open_db()
        query = f"SELECT favorite, title, file_ext, file_size, file_path FROM books WHERE metadata like 'None'"

        if only_favorites:
            query += " AND favorite = 1"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        # Преобразуем размер файла в человеко-читаемый формат
        return [(yes_no_indicator(favorite), title, file_ext, pretty_size(file_size), file_path) for favorite, title, file_ext, file_size, file_path in rows]

    # Получить статистику по расширениям файлов
    def get_file_extension_statistics(self):
        cursor = self.open_db()
        query = f"SELECT file_ext, COUNT(*) as count FROM books GROUP BY file_ext"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return rows
    
    # Поиск книг по части метаданных
    def search_books_by_metadata(self, metadata):
        cursor = self.open_db()
        query = f"SELECT title, author, metadata FROM books WHERE metadata LIKE '%{metadata}%'"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()

        return rows
    
    def plot_books_pages(self):
        cursor = self.open_db()
        query = "SELECT title, num_pages FROM books WHERE num_pages IS NOT NULL"

        cursor.execute(query)
        rows = cursor.fetchall()

        self.close_db()
        
        # Разделяем данные на названия книг и количество страниц
        titles = [row[0] for row in rows]
        num_pages = [row[1] for row in rows]

        # Создаем график
        plt.figure(figsize=(10, 6))
        plt.barh(titles, num_pages, color='skyblue')
        plt.xlabel('Number of Pages')
        plt.ylabel('Book Titles')
        plt.title('Number of Pages in Each Book')
        plt.tight_layout()
        plt.show()

def main():
    # Создаем парсер аргументов командной строки
    parser = argparse.ArgumentParser(description='Book Analyzer')
    parser.add_argument('--db_path', default='books.db', help='Path to the database file')
    parser.add_argument('--dir_path', required=True, help='Path to the directory to analyze')
    parser.add_argument('--file_types', nargs='+', default=['pdf'], help='File types to process')
    parser.add_argument('--exclude', nargs='+', default=[], help='Directories to exclude')
    parser.add_argument('--max_depth', type=int, default=5, help='Maximum directory depth to process')
    parser.add_argument('--web_page', help='Path to the generated web page')

    # Анализируем аргументы командной строки
    args = parser.parse_args()

    # Создаем экземпляр BookAnalyzer
    analyzer = BookAnalyzer(args.db_path)

    # Обрабатываем указанный каталог
    analyzer.process_directory(args.dir_path, args.file_types, args.exclude, args.max_depth)

    # Генерируем веб-страницу, если указан соответствующий аргумент
    if args.web_page is not None:
        analyzer.generate_web_page(args.web_page)

if __name__ == '__main__':
    main()
