import os
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, Menu, ttk
from bookAnalyzer import BookAnalyzer


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

        self.display_button = tk.Button(self.root, text="Показать книги", command=self.display_books)
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

        # Задаём растягиваемость строк и столбцов
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)  # Изменим номер строки для растягиваемости, т.к. текстовое поле теперь во второй строке (нумерация с 0)

    def display_books(self):
        try:
            books = self.analyzer.display_all_books()

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Очищаем текстовое поле
            self.text_widget.configure(state='normal')
            self.text_widget.delete(1.0, 'end')

            # Вставляем новые данные
            for book in books:
                self.text_widget.insert('end', str(book) + '\n')

            self.text_widget.configure(state='disabled')
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def process_directory(self):
        try:
            directory = filedialog.askdirectory()
            file_types = simpledialog.askstring("Ввод", "Введите типы файлов (через запятую):")
            file_types = [file_type.strip() for file_type in file_types.split(',')]
            exclude = simpledialog.askstring("Ввод", "Введите директории для исключения (через запятую):")
            exclude = [dir_.strip() for dir_ in exclude.split(',')]
            max_depth = simpledialog.askinteger("Ввод", "Введите максимальную глубину:")
            convert_odt_to_pdf = simpledialog.askstring("Ввод", "Конвертировать ODT в PDF? (д/н):").lower() == 'д'
            convert_docx_to_pdf = simpledialog.askstring("Ввод", "Конвертировать DOCX в PDF? (д/н):").lower() == 'д'

            self.analyzer.process_directory(directory, file_types=file_types, exclude=exclude,
                                            max_depth=max_depth, convert_odt_to_pdf=convert_odt_to_pdf,
                                            convert_docx_to_pdf=convert_docx_to_pdf)
            messagebox.showinfo("Успех", "Директория обработана успешно")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            
    def reset_database(self):
        try:
            reset = simpledialog.askstring("Ввод", "Перезапустить базу данных? (д/н):").lower() == 'д'
            convert_odt_to_pdf = simpledialog.askstring("Ввод", "Конвертировать ODT в PDF? (д/н):").lower() == 'д'
            convert_docx_to_pdf = simpledialog.askstring("Ввод", "Конвертировать DOCX в PDF? (д/н):").lower() == 'д'

            self.analyzer = BookAnalyzer('books.db', reset=reset, convert_odt_to_pdf=convert_odt_to_pdf, convert_docx_to_pdf=convert_docx_to_pdf)
            messagebox.showinfo("Успех", "База данных перезапущена успешно")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def search_books_by_title(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'Num Pages', 'Path'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Num Pages', text='Кол-во страниц')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[3]  # предполагается, что путь к файлу хранится в 4-м столбце
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)


            title = simpledialog.askstring("Ввод", "Введите название книги:")
            books = self.analyzer.search_books_by_title(title)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def search_books_by_author(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'Num Pages'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Num Pages', text='Кол-во страниц')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")


            author = simpledialog.askstring("Ввод", "Введите имя автора:")
            books = self.analyzer.search_books_by_author(author)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def search_books_by_extension(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'File Extension'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('File Extension', text='Расширение файла')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")


            extension = simpledialog.askstring("Ввод", "Введите расширение файла:")
            books = self.analyzer.search_books_by_extension(extension)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


root = tk.Tk()
analyzer = BookAnalyzer('books.db')  # создайте свой экземпляр анализатора здесь
app = App(root, analyzer)
root.mainloop()
