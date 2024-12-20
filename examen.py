import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def load_data():
    file_path = filedialog.askopenfilename(
        title="Виберіть файл з даними", 
        filetypes=(("CSV files", "*.csv"), ("Excel files", "*.xlsx")))
    if file_path:
        try:
            if file_path.endswith('.csv'):
                data = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                data = pd.read_excel(file_path)
            else:
                messagebox.showerror("Помилка", "Непідтримуваний формат файлу.")
                return None
            messagebox.showinfo("Успіх", "Дані успішно завантажені.")
            return data
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося завантажити дані: {e}")
    return None

def generate_report(data, start_date, end_date):
    try:
        data['Date'] = pd.to_datetime(data['Date'])
        
        filtered_data = data[(data['Date'] >= start_date) & (data['Date'] <= end_date)]
        
        report = filtered_data.groupby('Category').agg({
            'Sales': 'sum',
            'Quantity': 'sum'
        }).reset_index()
        
        total = pd.DataFrame({
            'Category': ['Загалом'],
            'Sales': [filtered_data['Sales'].sum()],
            'Quantity': [filtered_data['Quantity'].sum()]
        })
        report = pd.concat([report, total], ignore_index=True)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Зберегти звіт як")
        if save_path:
            report.to_excel(save_path, index=False)
            messagebox.showinfo("Успіх", "Звіт успішно збережено.")
    except Exception as e:
        messagebox.showerror("Помилка", f"Не вдалося згенерувати звіт: {e}")

class SalesReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор звітів з продажів")

        self.data = None

        self.load_button = tk.Button(root, text="Завантажити дані", command=self.load_data)
        self.load_button.pack(pady=10)

        self.start_date_label = tk.Label(root, text="Початкова дата (YYYY-MM-DD):")
        self.start_date_label.pack()
        self.start_date_entry = tk.Entry(root)
        self.start_date_entry.pack()

        self.end_date_label = tk.Label(root, text="Кінцева дата (YYYY-MM-DD):")
        self.end_date_label.pack()
        self.end_date_entry = tk.Entry(root)
        self.end_date_entry.pack()

        self.generate_button = tk.Button(root, text="Згенерувати звіт", command=self.generate_report)
        self.generate_button.pack(pady=10)

    def load_data(self):
        self.data = load_data()

    def generate_report(self):
        if self.data is None:
            messagebox.showerror("Помилка", "Спочатку завантажте дані.")
            return

        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Помилка", "Невірний формат дати. Використовуйте YYYY-MM-DD.")
            return

        generate_report(self.data, start_date, end_date)

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesReportApp(root)
    root.mainloop()
