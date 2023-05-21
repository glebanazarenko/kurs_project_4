import os
import io
import sys
import argparse
import json
import math
import fitz
import ebooklib #для работы с ePub-файлами
from ebooklib import epub
from bs4 import BeautifulSoup #для обработки HTML-кода в ePub-файлах
from pathlib import Path
import sqlite3
import matplotlib.pyplot as plt
from typing import List, Tuple
from PyPDF2 import PdfReader
from PIL import Image
from pdf2image import convert_from_path
from jinja2 import Environment, FileSystemLoader
import base64

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
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        # Создайте БД (если не существует) и определите таблицы для хранения метаданных книг и превью
        # Создаем соединение с БД, если файла нет, он будет создан
        conn = sqlite3.connect(self.db_path)

        # Создаем курсор для выполнения SQL-запросов
        cursor = conn.cursor()

        #удаление таблицы в бд
        cursor.execute('''
            DROP TABLE IF EXISTS books
        ''')

        # Создаем таблицу, если она не существует
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                author TEXT,
                publication_year INTEGER,
                file_path TEXT UNIQUE,
                file_size INTEGER,
                num_pages INTEGER,
                preview BLOB,
                metadata TEXT
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
                    title = metadata.get('/Title')
                    author = metadata.get('/Author')

            elif file_ext == '.epub':
                book = epub.read_epub(file_path)
                metadata = book.metadata
                num_pages = self.__count_generator_items(book.get_items_of_type(ebooklib.ITEM_DOCUMENT))
                preview = self.__get_epub_cover(book)

                # Извлечение данных о названии и авторе
                dc_metadata = metadata.get('http://purl.org/dc/elements/1.1/')
                title = dc_metadata['title'][0][0] if dc_metadata and 'title' in dc_metadata else None
                author = dc_metadata['creator'][0][0] if dc_metadata and 'creator' in dc_metadata else None

            print(metadata)

            # Если в метаданных нет названия, используем имя файла без расширения
            if not title:
                title = os.path.splitext(os.path.basename(file_path))[0]

            # Извлечение размера файла
            file_size = pretty_size(os.path.getsize(file_path))

            # Проверяем, есть ли уже книга в БД
            cursor.execute('SELECT * FROM books WHERE file_path = ?', (file_path,))
            row = cursor.fetchone()

            if row is None:
                # Если книги нет в БД, добавляем ее
                # Выполняем запрос к базе данных

                # Подготавливаем данные для вставки
                data = (file_path, title, author, file_size, str(metadata), num_pages, preview)

                cursor.execute("""
                    INSERT OR REPLACE INTO books (file_path, title, author, file_size, metadata, num_pages, preview)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, data)
            else:
                # Если книга уже есть в БД, обновляем ее

                # Подготавливаем данные для вставки
                data = (title, author, file_size, str(metadata), num_pages, preview, file_path)

                cursor.execute('''
                    UPDATE books SET title = ?, author = ?, file_size = ?, metadata = ?, num_pages = ?, preview = ?
                    WHERE file_path = ?
                ''', data)
        except Exception as e:
            print(f"Failed to process file {file_path}. Reason: {e}")
        
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
    def __count_generator_items(generator):
            return sum(1 for _ in generator)

    
    def __get_all_previews(self):
        # Создаем соединение с БД
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Получаем все превью из БД
        cursor.execute('SELECT title, preview FROM books')
        previews = cursor.fetchall()

        # Закрываем соединение
        conn.close()

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
    
    def process_directory(self, directory, file_types, exclude, max_depth, current_depth=0):
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

    def generate_web_page(self, output_path: str):
        # Создаем соединение с БД
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Получаем все книги из БД
        cursor.execute('SELECT * FROM books')
        books = cursor.fetchall()

        # Закрываем соединение
        conn.close()

        # Конвертируем превью из BLOB в base64 для отображения на веб-странице
        for book in books:
            book['preview'] = base64.b64encode(book['preview']).decode('utf-8')

        # Загружаем шаблон из файла
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template('template.html')

        # Генерируем HTML на основе шаблона и данных книг
        html = template.render(books=books)

        # Сохраняем HTML в файл
        with open(output_path, 'w') as file:
            file.write(html)

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
