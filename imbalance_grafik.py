import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, Label, Button, Entry, Scale, HORIZONTAL, messagebox
import os

def select_file():
    file_path = filedialog.askopenfilename(
        title="Выберите файл market_report.xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if file_path:
        process_file(file_path)
    else:
        messagebox.showwarning("Предупреждение", "Файл не выбран!")

def update_entry(value):
    """Обновляет значение в поле ввода при изменении ползунка."""
    window_entry.delete(0, 'end')
    window_entry.insert(0, int(float(value)))  # Преобразуем значение из ползунка в целое число

def update_scale(event=None):
    """Обновляет значение ползунка при изменении поля ввода."""
    try:
        value = int(window_entry.get())
        if 1 <= value <= 100:
            window_scale.set(value)
        else:
            raise ValueError
    except ValueError:
        messagebox.showerror("Ошибка", "Введите корректное значение (целое число от 1 до 100).")
        window_entry.delete(0, 'end')
        window_entry.insert(0, window_scale.get())  # Восстановить предыдущее значение

def process_file(file_path):
    try:
        # Чтение данных из файла Excel
        df = pd.read_excel(file_path)

        # Преобразование столбца Time в формат datetime с использованием errors='coerce'
        df['Time'] = pd.to_datetime(df['Time'], format='%Y.%m.%d %H:%M', errors='coerce')

        # Удаление строк с некорректными значениями
        df.dropna(subset=['Time'], inplace=True)

        # Получение значения степени сглаживания из поля ввода
        try:
            window_size = int(window_entry.get())
            if window_size <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное положительное число для степени сглаживания.")
            return

        # Сглаживание данных
        df['Smoothed Total Dizbalance'] = df['Total Dizbalance'].rolling(window=window_size, center=True).mean()

        # Построение графика
        plt.figure(figsize=(12, 6))
        plt.plot(df['Time'], df['Smoothed Total Dizbalance'], label=f"Сглаженные значения (window={window_size})", color="blue")
        plt.title("Сглаженный график значений из market_report.xlsx")
        plt.xlabel("Время")
        plt.ylabel("Значение")
        plt.grid(True)
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Сохранение графика в ту же папку, где находится файл market_report.xlsx
        save_path = os.path.join(os.path.dirname(file_path), "market_report.png")
        plt.savefig(save_path)
        plt.show()

        messagebox.showinfo("Успех", f"График успешно сохранен в файл: {save_path}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при обработке файла:\n{e}")

# Создание главного окна GUI
root = Tk()
root.title("Графический анализ market_report.xlsx")
root.geometry("500x400")

# Добавление элементов интерфейса
Label(root, text="Выберите файл и установите степень сглаживания:", font=("Arial", 12)).pack(pady=10)

Button(root, text="Выбрать файл", command=select_file, font=("Arial", 12)).pack(pady=10)

# Поле ввода для степени сглаживания
Label(root, text="Степень сглаживания (window):", font=("Arial", 12)).pack(pady=5)
window_entry = Entry(root, font=("Arial", 12))
window_entry.pack(pady=5)
window_entry.insert(0, "20")  # Значение по умолчанию
window_entry.bind("<Return>", update_scale)  # Обновление ползунка при нажатии Enter

# Ползунок для степени сглаживания
window_scale = Scale(root, from_=1, to=100, orient=HORIZONTAL, font=("Arial", 12), command=update_entry)
window_scale.set(20)  # Значение по умолчанию
window_scale.pack(pady=10)

# Запуск главного цикла GUI
root.mainloop()