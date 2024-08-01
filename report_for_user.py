import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class SalesAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Sales Data Analyzer")
        self.root.geometry("1200x600")
        
        self.filepath = ""
        
        # Create frame for buttons
        self.function_frame = tk.Frame(self.root, bg="#f0f0f0")
        self.function_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # Create buttons
        self.load_button = tk.Button(self.function_frame, text="Chọn tệp dữ liệu", command=self.load_file)
        self.load_button.pack(pady=10, padx=20)
        
        self.analyze_button = tk.Button(self.function_frame, text="Bắt đầu phân tích", command=self.analyze_data)
        self.analyze_button.pack(pady=10, padx=20)
        
        self.plot_option = tk.StringVar()
        self.plot_option.set("Doanh thu theo tháng")
        self.plot_option_menu = tk.OptionMenu(self.function_frame, self.plot_option, "Doanh thu theo tháng", "Doanh thu theo thành phố", "Số đơn hàng theo giờ", "Top 10 Combo sản phẩm được bán cùng nhau nhiều nhất", "Biểu đồ số lượng đơn hàng và giá của mỗi sản phẩm")
        self.plot_option_menu.pack(pady=10, padx=20)

        self.plot_button = tk.Button(self.function_frame, text="Vẽ biểu đồ", command=self.plot_data)
        self.plot_button.pack(pady=10, padx=20)

        self.save_button = tk.Button(self.function_frame, text="Xuất PDF", command=self.save_pdf)
        self.save_button.pack(pady=10, padx=20)
        
        self.clear_button = tk.Button(self.function_frame, text="Xóa biểu đồ", command=self.clear_plot)
        self.clear_button.pack(pady=10, padx=20)
        
        # Create frame for plot
        self.plot_frame = tk.Frame(self.root, bg="#f0f0f0")
        self.plot_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Create figure and canvas for plot
        self.figure = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.df = None
        
    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.filepath:
            messagebox.showinfo("Thông báo", f"Đã chọn tệp: {self.filepath}")
        else:
            messagebox.showerror("Lỗi", "Không tìm thấy tệp CSV nào!")
        
    def analyze_data(self):
        if not self.filepath:
            messagebox.showerror("Lỗi", "Hãy chọn tệp CSV trước!")
            return
        
        try:
            self.df = pd.read_csv(self.filepath)
            self.df = self.clean_data(self.df)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi trong quá trình phân tích: {str(e)}")
    
    def clean_data(self, df):
        # Add 'Month' column
        df['Month'] = df['Order Date'].str.slice(start=0, stop=2)
        df = df.dropna(how='all')
        df = df[df['Month'] != 'Or']
        # Convert data types
        df['Quantity Ordered'] = pd.to_numeric(df['Quantity Ordered'], downcast="integer",  errors='coerce')
        df['Price Each'] = pd.to_numeric(df['Price Each'], downcast="float",  errors='coerce')
        # Create 'Sales' column
        df['Sales'] = df['Quantity Ordered'] * df['Price Each']
        moving_c = df.pop("Sales")
        df.insert(4, 'Sales', moving_c)

        # Convert 'Order Date' to datetime
        df['Order Date'] = pd.to_datetime(df['Order Date'], format='mixed')

        # Extract city from purchase address
        df['City'] = df['Purchase Address'].apply(lambda address: address.split(',')[1].strip())
        df['Hours'] = df['Order Date'].dt.hour
        
        return df

    def plot_data(self):
        if self.df is None:
            messagebox.showerror("Lỗi", "Hãy phân tích dữ liệu trước!")
            return
        
        self.clear_plot()  # Clear existing plot
        
        if self.plot_option.get() == "Doanh thu theo tháng":
            self.plot_sales_by_month()
        elif self.plot_option.get() == "Doanh thu theo thành phố":
            self.plot_sales_by_city()
        elif self.plot_option.get() == "Số đơn hàng theo giờ":
            self.plot_sales_by_hour()
        elif self.plot_option.get() == "Top 10 Combo sản phẩm được bán cùng nhau nhiều nhất":
            self.plot_products_sold_together()
        elif self.plot_option.get() == "Biểu đồ số lượng đơn hàng và giá của mỗi sản phẩm":
            self.plot_most_sold_products()
    
    def clear_plot(self):
        self.ax.clear()
        self.canvas.draw()

    def plot_sales_by_month(self):
        df = self.df
        df['Month'] = df['Order Date'].dt.month
        sales_value = df.groupby('Month').count()['Sales']
        months = range(1, 13)
        self.ax.bar(x=months, height=sales_value.reindex(months, fill_value=0), color='skyblue', edgecolor='black', width=0.4)
        self.ax.set_xticks(months)
        self.ax.set_xlabel('Tháng', fontsize=12)
        self.ax.set_ylabel('Doanh thu (USD)', fontsize=12)
        self.ax.set_title('Doanh thu theo tháng', fontsize=14, fontweight='bold')
        self.ax.grid(axis='y', linestyle='--', alpha=0.7)
        for bar in self.ax.patches:
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width() / 2, height + 50, f'{height:.2f}', ha='center', va='bottom', fontsize=8, color='blue')
        self.canvas.draw()
    
    def plot_sales_by_city(self):
        df = self.df
        sales_value_city = df.groupby('City').count()['Sales']
        cities = sales_value_city.index.tolist()
        self.ax.bar(x=cities, height=sales_value_city, color='skyblue', edgecolor='black', width=0.4)
        self.ax.set_xticks(cities)
        self.ax.set_xticklabels(cities, rotation=45, ha='right')
        self.ax.set_xlabel('Thành phố', fontsize=12)
        self.ax.set_ylabel('Doanh thu (USD)', fontsize=12)
        self.ax.set_title('Doanh thu theo thành phố', fontsize=14, fontweight='bold')
        self.ax.grid(axis='y', linestyle='--', alpha=0.7)
        for bar in self.ax.patches:
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width() / 2, height + 50, f'{height:.2f}', ha='center', va='bottom', fontsize=8, color='blue')
        self.canvas.draw()

    def plot_sales_by_hour(self):
        df = self.df
        sales_value_hours = df.groupby('Hours').count()['Sales']
        hours = sales_value_hours.index.tolist()
        self.ax.bar(hours, sales_value_hours, color='skyblue')
        self.ax.set_xticks(hours)
        self.ax.set_xlabel('Giờ', fontsize=14)
        self.ax.set_ylabel('Số đơn hàng', fontsize=14)
        self.ax.set_title('Số đơn hàng theo giờ', fontsize=16, fontweight='bold')
        self.ax.grid(axis='y', linestyle='--', alpha=0.7)
        self.canvas.draw()

    def plot_products_sold_together(self):
        df = self.df
        df_dup = df[df['Order ID'].duplicated(keep=False)]
        df_dup['All Product'] = df_dup.groupby('Order ID')['Product'].transform(lambda x: ', '.join(x))
        df_dup = df_dup[['Order ID', 'All Product']].drop_duplicates()
        product_combo = df_dup['All Product'].value_counts().head(10)
        self.ax.bar(product_combo.index, product_combo.values, color='skyblue', edgecolor='black')
        self.ax.set_xticks(range(len(product_combo)))
        self.ax.set_xticklabels(product_combo.index, rotation=45, ha='right')
        self.ax.set_xlabel('Combo sản phẩm', fontsize=14)
        self.ax.set_ylabel('Số lượng', fontsize=14)
        self.ax.set_title('Top 10 Combo sản phẩm được bán cùng nhau nhiều nhất', fontsize=16, fontweight='bold')
        self.ax.grid(axis='y', linestyle='--', alpha=0.7)
        self.canvas.draw()
    
    def plot_most_sold_products(self):
        df = self.df
        prices = df.groupby('Product')['Price Each'].mean()

        # Cộng tất cả số lượng của từng sản phẩm
        all_products = df.groupby('Product')['Quantity Ordered'].sum()

        # Sản phẩm bán ra
        products_ls = all_products.index
        y1 = all_products.values
        y2 = prices[products_ls].reindex(products_ls).values  # Đảm bảo y2 có cùng kích thước với y1

        self.clear_plot()  # Xóa biểu đồ cũ

        # Vẽ biểu đồ cột cho số lượng đặt hàng
        self.ax.bar(products_ls, y1, color='g', align='center')
        
        # Tạo trục thứ hai để vẽ biểu đồ đường cho giá của từng sản phẩm
        ax2 = self.ax.twinx()
        ax2.plot(products_ls, y2, 'b-o')

        # Thiết lập các nhãn và tiêu đề
        self.ax.set_xticklabels(products_ls, rotation=90, size=12)
        self.ax.set_xlabel('Sản phẩm', fontsize=14)
        self.ax.set_ylabel('Số lượng đặt hàng', color='g', fontsize=14)
        ax2.set_ylabel('Giá mỗi sản phẩm (USD)', color='b', fontsize=14)
        self.ax.set_title('Biểu đồ số lượng đơn hàng và giá của mỗi sản phẩm', fontsize=16)

        # Thêm chú thích (legend)
        ax2.legend(['Giá mỗi sản phẩm'], loc='upper left')
        self.ax.legend(['Số lượng đặt hàng'], loc='upper right')
        ax2.grid(False)
        self.ax.grid(True)
        
        # Sử dụng tight_layout để tránh chồng lấn các nhãn
        self.figure.tight_layout()
        
        # Cập nhật Canvas để hiển thị biểu đồ mới
        self.canvas.draw()


        
    def save_pdf(self):
        if self.df is None:
            messagebox.showerror("Lỗi", "Hãy phân tích dữ liệu trước!")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            try:
                self.figure.savefig(file_path, format='pdf')
                messagebox.showinfo("Thông báo", f"Đã lưu biểu đồ vào: {file_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Đã xảy ra lỗi khi lưu tệp: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1200x600")
    app = SalesAnalyzer(root)
    root.mainloop()
