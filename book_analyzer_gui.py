import os
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, Menu, ttk
from bookAnalyzer import BookAnalyzer
import matplotlib.pyplot as plt
from PIL import Image, ImageTk
import numpy as np
from io import BytesIO


class App:
    def __init__(self, root, analyzer):
        self.root = root
        self.analyzer = analyzer

        self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'Num Pages'), show='headings')
        self.tree.heading('Title', text='Title')
        self.tree.heading('Author', text='Author')
        self.tree.heading('Num Pages', text='Num Pages')
        self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

        self.text_widget = tk.Text(self.root)
        self.text_widget.grid(row=1, column=0, columnspan=4, sticky="nsew")
        self.text_widget.configure(state='disabled')  # делаем текстовое поле read-only

        self.favorites_var = tk.IntVar()
        self.last_method = self.display_all_books
        self.last_args = dict(limit=30, offset=1)

        self.favorites_checkbox = tk.Checkbutton(self.root, text="Показать только избранные", variable=self.favorites_var, command=self.update_table)
        self.favorites_checkbox.grid(row=0, column=3)

        self.display_button = tk.Button(self.root, text="Показать все книги", command=self.display_all_books)
        self.display_button.grid(row=0, column=0)

        self.process_directory_button = tk.Button(self.root, text="Обработать директорию", command=self.process_directory)
        self.process_directory_button.grid(row=0, column=1)

        self.reset_database_button = tk.Button(self.root, text="Перезапустить базу данных", command=self.reset_database)
        self.reset_database_button.grid(row=0, column=2)

        # Создание выпадающего меню
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Основные функции", menu=file_menu)
        file_menu.add_command(label="Поиск книг по названию", command=self.search_books_by_title)
        file_menu.add_command(label="Поиск книг по автору", command=self.search_books_by_author)
        file_menu.add_command(label="Поиск книг по расширению", command=self.search_books_by_extension)
        file_menu.add_command(label="Самые большие книги", command=self.display_largest_books)
        file_menu.add_command(label="Книги с наибольшим количеством страниц", command=self.display_books_with_most_pages)
        file_menu.add_command(label="Недавно добавленные книги", command=self.display_recently_added_books)
        file_menu.add_command(label="Книги без автора", command=self.display_books_without_author)
        file_menu.add_command(label="Книги без метаданных", command=self.display_books_without_metadata)
        file_menu.add_command(label="Показать статистику расширений файлов", command=self.display_file_extension_statistics)
        file_menu.add_command(label="Показать график количества страниц", command=self.display_books_pages_chart)


        # Задаём растягиваемость строк и столбцов
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)  # Изменим номер строки для растягиваемости, т.к. текстовое поле теперь во второй строке (нумерация с 0)

        # Показываем все книги при старте приложения
        self.display_all_books()

    def update_table(self):
        # Обновляем таблицу с учетом текущего значения чекбокса
        self.last_method(**self.last_args)

    def display_all_books(self, limit=None, offset=None):
        try:
            # Создаем новую таблицу с нужными столбцами
            self.tree = ttk.Treeview(self.root, columns=('Favorite', 'Title', 'Author', 'File Ext', 'File Path', 'File Size', 'Num Pages', 'Metadata'), show='headings')
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Author', text='Автор', command=lambda: self.treeview_sort_column(self.tree, 'Author', False))
            self.tree.heading('File Ext', text='Расширение файла', command=lambda: self.treeview_sort_column(self.tree, 'File Ext', False))
            self.tree.heading('File Path', text='Путь к файлу', command=lambda: self.treeview_sort_column(self.tree, 'File Path', False))
            self.tree.heading('File Size', text='Размер файла', command=lambda: self.treeview_sort_column(self.tree, 'File Size', False))
            self.tree.heading('Num Pages', text='Кол-во страниц', command=lambda: self.treeview_sort_column(self.tree, 'Num Pages', False))
            self.tree.heading('Metadata', text='Метаданные', command=lambda: self.treeview_sort_column(self.tree, 'Metadata', False))
            self.tree.grid(row=1, column=0, columnspan=8, sticky="nsew")

            def change_favorite(event):
                item = self.tree.selection()[0]
                is_favorite = self.tree.item(item, "values")[0]
                file_path = self.tree.item(item, "values")[4]
                # Обновление значения в столбце "Избранное"
                if is_favorite == 'Да':
                    self.tree.set(item, '#1', 'Нет')
                    self.analyzer.update_book_favorite_status(file_path) 
                else:
                    self.tree.set(item, '#1', 'Да')
                    self.analyzer.update_book_favorite_status(file_path) 

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[4]  # предполагается, что путь к файлу хранится в 5-м столбце
                os.startfile(file_path)

            def show_preview(event):
                try:
                    item = self.tree.identify('item', event.x, event.y)
                    file_path = self.tree.item(item, "values")[4]  # Путь к файлу хранится в 5 столбце

                    # Получаем данные предварительного просмотра из базы данных
                    preview_bytes = self.analyzer.get_book_preview_path(file_path)

                    if preview_bytes:
                        # Преобразуем данные предварительного просмотра в изображение
                        image = Image.open(BytesIO(preview_bytes))

                        # Выводим изображение
                        plt.figure()
                        plt.imshow(image)
                        plt.axis('off')
                        plt.show()

                    else:
                        messagebox.showinfo("Превью", "Нет Превью для этой книги.")

                except Exception as e:
                    messagebox.showerror("Ошибка", str(e))

            def show_metadata(event):
                try:
                    item = self.tree.selection()[0]
                    metadata_str = self.tree.item(item, "values")[7]  # Метаданные хранятся в 8-м столбце
                    metadata = eval(metadata_str)  # Преобразуем строку метаданных в словарь

                    dialog = tk.Toplevel(self.root)
                    dialog.title("Метаданные книги")

                    # Создаем новую таблицу с метаданными
                    metadata_tree = ttk.Treeview(dialog, columns=('Key', 'Value'), show='headings')
                    metadata_tree.heading('Key', text='Ключ')
                    metadata_tree.heading('Value', text='Значение')
                    metadata_tree.grid(row=0, column=0, sticky='nsew')

                    # Конфигурируем ряды и столбцы для корректного изменения размера таблицы
                    dialog.grid_rowconfigure(0, weight=1)
                    dialog.grid_columnconfigure(0, weight=1)

                    # Вставляем метаданные
                    for key, value in metadata.items():
                        metadata_tree.insert('', 'end', values=(key, value))

                except Exception as e:
                    messagebox.showerror("Ошибка", str(e))

            self.tree.bind('<Double-1>', open_file)
            self.tree.bind('<Double-3>', show_preview)  # Правый клик для предварительного просмотра
            self.tree.bind("<Control-f>", change_favorite) # ctrl + f
            self.tree.bind("<Control-m>", show_metadata) # ctrl + m

            only_favorites = self.favorites_var.get() == 1

            if limit is None and offset is None and self.last_method == self.display_all_books and only_favorites:
                limit = self.last_args.get('limit')
                offset = self.last_args.get('offset')
            elif limit is None and offset is None:
                dialog = tk.Toplevel(self.root)
                dialog.title("Выбор файлов")

                limit_label = tk.Label(dialog, text="Введите сколько файлов хотите выбрать")
                limit_label.pack()
                limit_entry = tk.Entry(dialog)
                limit_entry.pack()

                offset_label = tk.Label(dialog, text="Начать с ? (начинается с 1)")
                offset_label.pack()
                offset_entry = tk.Entry(dialog)
                offset_entry.pack()

                def on_submit():
                    nonlocal limit
                    nonlocal offset
                    limit = int(limit_entry.get())
                    offset = int(offset_entry.get()) - 1
                    dialog.destroy()

                submit_button = tk.Button(dialog, text="Подтвердить", command=on_submit)
                submit_button.pack()
                self.root.wait_window(dialog)

            books = self.analyzer.get_all_books(only_favorites, limit, offset)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

            self.last_method = self.display_all_books
            self.last_args = dict(limit=limit, offset=offset)

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
    
    # Функция для сортировки столбцов
    def treeview_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]

        # Маппинг единиц измерения к их весу в байтах
        units = {'байт': 1, 'КБ': 1024, 'МБ': 1024**2, 'ГБ': 1024**3, 'ТБ': 1024**4, 'ПБ': 1024**5}

        # Проверяем, являются ли данные числами или строками, и применяем соответствующую функцию сортировки
        if col == "File Size":
            l.sort(key=lambda t: float(t[0].split()[0]) * units[t[0].split()[1]], reverse=reverse)
        elif col == "Num Pages" or col == "Rank":
            l.sort(key=lambda t: int(t[0]), reverse=reverse)
        else:
            l.sort(reverse=reverse)

        # Переставляем элементы в отсортированном порядке.
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # Перевернем направление сортировки для следующего щелчка.
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def process_directory(self):
        try:
            directory = filedialog.askdirectory()

            dialog = tk.Toplevel(self.root)
            dialog.title("Выбор параметров обработки")

            # Для хранения значений
            file_types = tk.StringVar(value='odt,docx,pdf,epub')
            exclude_dirs = tk.StringVar(value='')
            max_depth = tk.IntVar(value=1)
            convert_odt_to_pdf = tk.BooleanVar(value=False)
            convert_docx_to_pdf = tk.BooleanVar(value=False)

            # Виджеты для ввода данных
            tk.Label(dialog, text="Типы файлов (через запятую):").pack()
            tk.Entry(dialog, textvariable=file_types).pack()

            tk.Label(dialog, text="Исключить директории (через запятую):").pack()
            tk.Entry(dialog, textvariable=exclude_dirs).pack()

            tk.Label(dialog, text="Максимальная глубина:").pack()
            tk.Spinbox(dialog, from_=1, to=10, textvariable=max_depth).pack()

            tk.Label(dialog, text="Конвертировать ODT в PDF:").pack()
            tk.Radiobutton(dialog, text="Да", variable=convert_odt_to_pdf, value=True).pack()
            tk.Radiobutton(dialog, text="Нет", variable=convert_odt_to_pdf, value=False).pack()

            tk.Label(dialog, text="Конвертировать DOCX в PDF:").pack()
            tk.Radiobutton(dialog, text="Да", variable=convert_docx_to_pdf, value=True).pack()
            tk.Radiobutton(dialog, text="Нет", variable=convert_docx_to_pdf, value=False).pack()

            def on_submit():
                file_types_list = [file_type.strip() for file_type in file_types.get().split(',')]
                exclude_dirs_list = [dir_.strip() for dir_ in exclude_dirs.get().split(',')]
                self.analyzer.process_directory(directory, file_types=file_types_list, exclude=exclude_dirs_list,
                                            max_depth=max_depth.get(),
                                            convert_odt_to_pdf=convert_odt_to_pdf.get(),
                                            convert_docx_to_pdf=convert_docx_to_pdf.get())
                messagebox.showinfo("Успех", "Директория обработана успешно")
                dialog.destroy()

            tk.Button(dialog, text="Подтвердить", command=on_submit).pack()

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            
    def reset_database(self):
        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("Перезапуск базы данных")

            # Для хранения значений
            reset_database = tk.BooleanVar(value=False)

            # Виджеты для ввода данных
            warning_label = tk.Label(dialog, text="Вы уверены? Это приведет к потери данных, что есть в приложении.", fg="red")
            warning_label.pack()

            reset_label = tk.Label(dialog, text="Перезапустить базу данных?")
            reset_label.pack()
            yes_button = tk.Radiobutton(dialog, text="Да", variable=reset_database, value=True)
            yes_button.pack()
            no_button = tk.Radiobutton(dialog, text="Нет", variable=reset_database, value=False)
            no_button.pack()

            def on_submit():
                if reset_database.get():
                    self.analyzer = BookAnalyzer('books.db', reset=reset_database.get())
                    messagebox.showinfo("Успех", "База данных перезапущена успешно")
                else:
                    messagebox.showinfo("Спасибо", "Вы сохранили жизнь книгам")
                dialog.destroy()

            submit_button = tk.Button(dialog, text="Подтвердить", command=on_submit)
            submit_button.pack()

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def show_metadata(self, tree):
        def show_metadata(event):
            # Получаем выбранный элемент в таблице
            item = self.tree.selection()[0]

            # Читаем путь файла из последнего столбца
            file_path = self.tree.item(item, "values")[-1]

            metadata = self.analyzer.get_book_metadata(file_path)

            dialog = tk.Toplevel(self.root)
            dialog.title("Метаданные книги")

            # Создаем новую таблицу с метаданными
            metadata_tree = ttk.Treeview(dialog, columns=('Key', 'Value'), show='headings')
            metadata_tree.heading('Key', text='Ключ')
            metadata_tree.heading('Value', text='Значение')
            metadata_tree.grid(row=0, column=0, sticky='nsew')

            # Конфигурируем ряды и столбцы для корректного изменения размера таблицы
            dialog.grid_rowconfigure(0, weight=1)
            dialog.grid_columnconfigure(0, weight=1)

            # Вставляем метаданные
            for key, value in metadata.items():
                metadata_tree.insert('', 'end', values=(key, value))
        tree.bind("<Control-m>", show_metadata) # ctrl + m


    def change_favorite(self, tree):
        def change_favorite(event):
            item = tree.selection()[0]
            is_favorite = tree.item(item, "values")[0]
            file_path = tree.item(item, "values")[-1]
            # Обновление значения в столбце "Избранное"
            if is_favorite == 'Да':
                self.tree.set(item, '#1', 'Нет')
                self.analyzer.update_book_favorite_status(file_path) 
            else:
                self.tree.set(item, '#1', 'Да')
                self.analyzer.update_book_favorite_status(file_path) 
        tree.bind("<Control-f>", change_favorite) # ctrl + f

    #когда сначала ранг, а потом избранное
    def change_favorite2(self, tree):
        def change_favorite(event):
            item = tree.selection()[0]
            is_favorite = tree.item(item, "values")[1]
            file_path = tree.item(item, "values")[-1]
            # Обновление значения в столбце "Избранное"
            if is_favorite == 'Да':
                self.tree.set(item, '#2', 'Нет')
                self.analyzer.update_book_favorite_status(file_path) 
            else:
                self.tree.set(item, '#2', 'Да')
                self.analyzer.update_book_favorite_status(file_path) 
        tree.bind("<Control-f>", change_favorite) # ctrl + f

    def open_file(self, tree):
        def op_file(event):
            item = tree.identify('item', event.x, event.y)
            file_path = tree.item(item, "values")[-1]  # Путь к файлу хранится в последнем столбце
            os.startfile(file_path)

        tree.bind('<Double-1>', op_file) # реагирует на ЛКМ


    def bind_preview(self, tree):
        def open_file(event):
            item = tree.identify('item', event.x, event.y)
            file_path = tree.item(item, "values")[-1]  # Путь к файлу хранится в последнем столбце

            # Получаем данные предварительного просмотра из базы данных
            preview_bytes = self.analyzer.get_book_preview_path(file_path)

            if preview_bytes:
                # Преобразуем данные предварительного просмотра в изображение
                image = Image.open(BytesIO(preview_bytes))

                # Выводим изображение
                plt.figure()
                plt.imshow(image)
                plt.axis('off')
                plt.show()

            else:
                messagebox.showinfo("Превью", "Нет Превью для этой книги.")

        tree.bind('<Double-3>', open_file) # реагирует на ПКМ

    # Поиск книг по названию
    def search_books_by_title(self, title=None):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Favorite', 'Title', 'Author', 'Num Pages', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Author', text='Автор', command=lambda: self.treeview_sort_column(self.tree, 'Author', False))
            self.tree.heading('Num Pages', text='Кол-во страниц', command=lambda: self.treeview_sort_column(self.tree, 'Num Pages', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite(self.tree)
            self.show_metadata(self.tree)

            if title is None and self.last_method == self.search_books_by_title:
                title = self.last_args.get('title')
            elif title is None:
                title = simpledialog.askstring("Ввод", "Введите название книги:")

            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.search_books_by_title(title, only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.search_books_by_title
            self.last_args = dict(title=title)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Поиск книг по автору
    def search_books_by_author(self, author=None):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Favorite', 'Title', 'Author', 'Num Pages', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Author', text='Автор', command=lambda: self.treeview_sort_column(self.tree, 'Author', False))
            self.tree.heading('Num Pages', text='Кол-во страниц', command=lambda: self.treeview_sort_column(self.tree, 'Num Pages', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite(self.tree)
            self.show_metadata(self.tree)

            if author is None and self.last_method == self.search_books_by_author:
                author = self.last_args.get('author')
            elif author is None:
                author = simpledialog.askstring("Ввод", "Введите имя автора:")

            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.search_books_by_author(author, only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.search_books_by_author
            self.last_args = dict(author=author)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Поиск книг по расширению файла
    def search_books_by_extension(self, extension=None):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Favorite', 'Title', 'Author', 'Extension', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Author', text='Автор', command=lambda: self.treeview_sort_column(self.tree, 'Author', False))
            self.tree.heading('Extension', text='Расширение файла', command=lambda: self.treeview_sort_column(self.tree, 'Extension', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite(self.tree)
            self.show_metadata(self.tree)

            if extension is None and self.last_method == self.search_books_by_extension:
                extension = self.last_args.get('extension')
            elif extension is None:
                extension = simpledialog.askstring("Ввод", "Введите расширение файла:")

            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.search_books_by_extension(extension, only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.search_books_by_extension
            self.last_args = dict(extension=extension)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Показать книги с наибольшим размером
    def display_largest_books(self, limit=None, offset=None):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Rank', 'Favorite', 'Title', 'Author', 'File Size', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.column('Rank', width=30)  # Здесь мы задаём ширину столбца 'Rank'
            self.tree.heading('Rank', text='Ранг', command=lambda: self.treeview_sort_column(self.tree, 'Rank', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Author', text='Автор', command=lambda: self.treeview_sort_column(self.tree, 'Author', False))
            self.tree.heading('File Size', text='Размер файла', command=lambda: self.treeview_sort_column(self.tree, 'File Size', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=6, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite2(self.tree)
            self.show_metadata(self.tree)

            if limit is None and offset is None and self.last_method == self.display_largest_books:
                limit = self.last_args.get('limit')
                offset = self.last_args.get('offset')
            elif limit is None and offset is None:
                dialog = tk.Toplevel(self.root)
                dialog.title("Выбор файлов")

                limit_label = tk.Label(dialog, text="Введите сколько файлов хотите выбрать")
                limit_label.pack()
                limit_entry = tk.Entry(dialog)
                limit_entry.pack()

                offset_label = tk.Label(dialog, text="Начать с ? (начинается с 1)")
                offset_label.pack()
                offset_entry = tk.Entry(dialog)
                offset_entry.pack()

                def on_submit():
                    nonlocal limit
                    nonlocal offset
                    limit = int(limit_entry.get())
                    offset = int(offset_entry.get()) - 1
                    dialog.destroy()

                submit_button = tk.Button(dialog, text="Подтвердить", command=on_submit)
                submit_button.pack()
                self.root.wait_window(dialog)

            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.get_largest_books(limit, offset, only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for i, book in enumerate(books, start=1 + offset):
                self.tree.insert('', 'end', values=(i, *book))

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.display_largest_books
            self.last_args = dict(limit=limit, offset=offset)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Показать книги с наибольшим количеством страниц
    def display_books_with_most_pages(self, limit=None, offset=None):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Rank', 'Favorite', 'Title', 'Author', 'Num Pages', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.column('Rank', width=50)  # Задаем ширину столбца 'Rank'
            self.tree.heading('Rank', text='Ранг', command=lambda: self.treeview_sort_column(self.tree, 'Rank', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Author', text='Автор', command=lambda: self.treeview_sort_column(self.tree, 'Author', False))
            self.tree.heading('Num Pages', text='Кол-во страниц', command=lambda: self.treeview_sort_column(self.tree, 'Num Pages', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=6, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite2(self.tree)
            self.show_metadata(self.tree)

            if limit is None and offset is None and self.last_method == self.display_books_with_most_pages:
                limit = self.last_args.get('limit')
                offset = self.last_args.get('offset')
            elif limit is None and offset is None:
                dialog = tk.Toplevel(self.root)
                dialog.title("Выбор файлов")

                limit_label = tk.Label(dialog, text="Введите сколько файлов хотите выбрать")
                limit_label.pack()
                limit_entry = tk.Entry(dialog)
                limit_entry.pack()

                offset_label = tk.Label(dialog, text="Начать с ? (начинается с 1)")
                offset_label.pack()
                offset_entry = tk.Entry(dialog)
                offset_entry.pack()

                def on_submit():
                    nonlocal limit
                    nonlocal offset
                    limit = int(limit_entry.get())
                    offset = int(offset_entry.get()) - 1
                    dialog.destroy()

                submit_button = tk.Button(dialog, text="Подтвердить", command=on_submit)
                submit_button.pack()
                self.root.wait_window(dialog)

            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.get_books_with_most_pages(limit, offset, only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for i, book in enumerate(books, start=1 + offset):
                self.tree.insert('', 'end', values=(i, *book))

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.display_books_with_most_pages
            self.last_args = dict(limit=limit, offset=offset)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Показать недавно добавленные книги
    def display_recently_added_books(self, limit=None, offset=None):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Rank', 'Favorite', 'Title', 'Author', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.column('Rank', width=50)  # Задаем ширину столбца 'Rank'
            self.tree.heading('Rank', text='Ранг', command=lambda: self.treeview_sort_column(self.tree, 'Rank', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Author', text='Автор', command=lambda: self.treeview_sort_column(self.tree, 'Author', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite2(self.tree)
            self.show_metadata(self.tree)

            if limit is None and offset is None and self.last_method == self.display_recently_added_books:
                limit = self.last_args.get('limit')
                offset = self.last_args.get('offset')
            elif limit is None and offset is None:
                dialog = tk.Toplevel(self.root)
                dialog.title("Выбор файлов")

                limit_label = tk.Label(dialog, text="Введите сколько файлов хотите выбрать")
                limit_label.pack()
                limit_entry = tk.Entry(dialog)
                limit_entry.pack()

                offset_label = tk.Label(dialog, text="Начать с ? (начинается с 1)")
                offset_label.pack()
                offset_entry = tk.Entry(dialog)
                offset_entry.pack()

                def on_submit():
                    nonlocal limit
                    nonlocal offset
                    limit = int(limit_entry.get())
                    offset = int(offset_entry.get()) - 1
                    dialog.destroy()

                submit_button = tk.Button(dialog, text="Подтвердить", command=on_submit)
                submit_button.pack()
                self.root.wait_window(dialog)

            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.get_recently_added_books(limit, offset, only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for i, book in enumerate(books, start=1 + offset):
                self.tree.insert('', 'end', values=(i, *book))

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.display_recently_added_books
            self.last_args = dict(limit=limit, offset=offset)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


    # Отображаем книги без автора
    def display_books_without_author(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Favorite', 'Title', 'Num Pages', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('Num Pages', text='Кол-во страниц', command=lambda: self.treeview_sort_column(self.tree, 'Num Pages', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite(self.tree)
            self.show_metadata(self.tree)
            
            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.get_books_without_author(only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.display_books_without_author
            self.last_args = dict()

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Отображаем книги без метаданных
    def display_books_without_metadata(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Favorite', 'Title', 'File Ext', 'File Size', 'Path'), show='headings')
            self.tree.column('Favorite', width=30)
            self.tree.heading('Favorite', text='Избранное', command=lambda: self.treeview_sort_column(self.tree, 'Favorite', False))
            self.tree.heading('Title', text='Название', command=lambda: self.treeview_sort_column(self.tree, 'Title', False))
            self.tree.heading('File Ext', text='Расширение файла', command=lambda: self.treeview_sort_column(self.tree, 'File Ext', False))
            self.tree.heading('File Size', text='Размер файла', command=lambda: self.treeview_sort_column(self.tree, 'File Size', False))
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew")

            self.open_file(self.tree)
            self.bind_preview(self.tree)
            self.change_favorite(self.tree)

            only_favorites = self.favorites_var.get() == 1
            books = self.analyzer.get_books_without_metadata(only_favorites)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

            # обновляем последний вызванный метод и его аргументы
            self.last_method = self.display_books_without_metadata
            self.last_args = dict()

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def display_file_extension_statistics(self):
        try:
            # Получить статистику
            data_list  = self.analyzer.get_file_extension_statistics()
            data = {ext: count for ext, count in data_list}
            total_files = sum(data.values())
            labels = [f"{ext}: {round((count/total_files)*100, 2)}%" for ext, count in data.items()]
            sizes = list(data.values()) #создание списка размеров для каждого сектора диаграммы

            fig, ax = plt.subplots()
            wedges, texts, autotexts = ax.pie(sizes, autopct='', pctdistance=1.1)

            bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
            kw = dict(arrowprops=dict(arrowstyle="-"),
                    bbox=bbox_props, zorder=0, va="center")

            for i, p in enumerate(wedges):
                ang = (p.theta2 - p.theta1)/2. + p.theta1 #вычисление угла середины каждого сектора
                y = np.sin(np.deg2rad(ang))
                x = np.cos(np.deg2rad(ang))
                horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
                connectionstyle = "angle,angleA=0,angleB={}".format(ang)
                kw["arrowprops"].update({"connectionstyle": connectionstyle})
                ax.annotate(labels[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.4*y),
                            horizontalalignment=horizontalalignment, **kw)

            # Показываем диаграмму
            plt.show()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def display_books_pages_chart(self):
        try:
            self.analyzer.plot_books_pages()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


root = tk.Tk()
analyzer = BookAnalyzer('books.db')  # создайте свой экземпляр анализатора здесь
app = App(root, analyzer)
root.mainloop()
