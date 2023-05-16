import os
import io
import sys
import argparse
import json
from pathlib import Path
import sqlite3
from typing import List, Tuple
from PyPDF2 import PdfReader
from PIL import Image
from pdf2image import convert_from_path
from jinja2 import Environment, FileSystemLoader
import base64

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

        try:
            with open(file_path, 'rb') as file:
                reader = PdfReader(file)
                metadata = reader.metadata
                num_pages = len(reader.pages)

                # Подготавливаем данные для вставки
                data = (file_path, str(metadata), num_pages)

                # Выполняем запрос к базе данных
                cursor.execute("""
                    INSERT OR REPLACE INTO books (file_path, metadata, num_pages)
                    VALUES (?, ?, ?)
                """, data)
        except Exception as e:
            print(f"Failed to process file {file_path}. Reason: {e}")
        
        # Сохраняем изменения и закрываем соединение
        conn.commit()
        conn.close()
    
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
