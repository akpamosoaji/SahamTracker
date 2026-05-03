import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import os
from tkcalendar import DateEntry

class PortfolioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Stock & Crypto Portfolio Tracker")
        self.root.geometry("1000x850")
        
        self.file_name = "portfolio_history.xlsx"
        self.portfolio = {} 
        self.cash_balance = 0.0 
        self.history_df = pd.DataFrame(columns=['Tanggal', 'Tipe', 'Ticker', 'Jumlah', 'Harga', 'Total', 'Realized_PL'])
        
        self.create_widgets()
        self.load_default_data()

    def load_default_data(self):
        if os.path.exists(self.file_name):
            try:
                self.history_df = pd.read_excel(self.file_name)
                self.rebuild_portfolio()
                self.update_tables()
            except Exception as e:
                print(f"Gagal memuat {self.file_name}: {e}")

    def create_widgets(self):
        # --- Top Section: Tombol File ---
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=5, padx=10, fill="x")
        
        ttk.Button(file_frame, text="Buka File Excel (Load)", command=self.load_from_excel).pack(side="left", padx=5)
        ttk.Button(file_frame, text="Simpan sebagai Excel (Save As)", command=self.save_to_excel).pack(side="left", padx=5)

        # --- Manajemen Cash Section ---
        cash_frame = ttk.LabelFrame(self.root, text="Manajemen Cash (Setor / Tarik Dana)")
        cash_frame.pack(pady=10, padx=10, fill="x")

        self.lbl_cash = ttk.Label(cash_frame, text="Saldo Cash: 0.00", font=("Arial", 16, "bold"), foreground="blue")
        self.lbl_cash.grid(row=0, column=0, columnspan=6, padx=10, pady=10, sticky="w")

        ttk.Label(cash_frame, text="Tanggal:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        # PERUBAHAN DI SINI: Menambahkan state='normal' agar bisa diketik manual
        self.cash_date_entry = DateEntry(cash_frame, width=12, background='darkblue', 
                                         foreground='white', borderwidth=2, 
                                         date_pattern='dd-mm-yyyy', state='normal')
        self.cash_date_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(cash_frame, text="Nominal Uang:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.cash_amount_entry = ttk.Entry(cash_frame)
        self.cash_amount_entry.grid(row=1, column=3, padx=5, pady=5)

        btn_deposit = ttk.Button(cash_frame, text="SETOR DANA", command=self.action_deposit)
        btn_deposit.grid(row=1, column=4, padx=5, pady=5)

        btn_withdraw = ttk.Button(cash_frame, text="TARIK DANA", command=self.action_withdraw)
        btn_withdraw.grid(row=1, column=5, padx=5, pady=5)

        # --- Input Section (Transaksi Saham) ---
        input_frame = ttk.LabelFrame(self.root, text="Transaksi Saham")
        input_frame.pack(pady=5, padx=10, fill="x")

        ttk.Label(input_frame, text="Tanggal:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # PERUBAHAN DI SINI: Menambahkan state='normal' agar bisa diketik manual
        self.date_entry = DateEntry(input_frame, width=12, background='darkblue', 
                                    foreground='white', borderwidth=2, 
                                    date_pattern='dd-mm-yyyy', state='normal')
        self.date_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Ticker (Kode):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.ticker_entry = ttk.Entry(input_frame)
        self.ticker_entry.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(input_frame, text="Jumlah Lot/Unit:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.qty_entry = ttk.Entry(input_frame)
        self.qty_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Harga Satuan:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.price_entry = ttk.Entry(input_frame)
        self.price_entry.grid(row=1, column=3, padx=5, pady=5)

        btn_buy = ttk.Button(input_frame, text="BELI SAHAM", command=self.action_buy)
        btn_buy.grid(row=1, column=4, padx=5, pady=5)

        btn_sell = ttk.Button(input_frame, text="JUAL SAHAM", command=self.action_sell)
        btn_sell.grid(row=1, column=5, padx=5, pady=5)

        # --- Portfolio Table ---
        table_frame = ttk.LabelFrame(self.root, text="Posisi Portofolio Saat Ini")
        table_frame.pack(pady=5, padx=10, fill="both", expand=True)

        self.tree_port = ttk.Treeview(table_frame, columns=("Ticker", "Qty", "Avg Price", "Value"), show='headings', height=4)
        
        self.tree_port.heading("Ticker", text="Ticker")
        self.tree_port.column("Ticker", anchor="center") 
        self.tree_port.heading("Qty", text="Jumlah (Unit)")
        self.tree_port.column("Qty", anchor="center") 
        self.tree_port.heading("Avg Price", text="Harga Rata-rata")
        self.tree_port.column("Avg Price", anchor="e") 
        self.tree_port.heading("Value", text="Total Modal")
        self.tree_port.column("Value", anchor="e") 
        
        self.tree_port.pack(fill="both", expand=True)
        
        # --- History Table & Delete Button ---
        history_frame = ttk.LabelFrame(self.root, text="Riwayat Transaksi & Profit/Loss")
        history_frame.pack(pady=5, padx=10, fill="both", expand=True)

        btn_delete = ttk.Button(history_frame, text="Hapus Transaksi Terpilih (Bisa Multiple)", command=self.action_delete)
        btn_delete.pack(pady=5, anchor="w", padx=5)

        cols = ("Tanggal", "Tipe", "Ticker", "Jumlah", "Harga", "Total", "P/L")
        self.tree_hist = ttk.Treeview(history_frame, columns=cols, show='headings', height=7, selectmode="extended")
        
        for c in cols:
            self.tree_hist.heading(c, text=c)
            if c == "Jumlah":
                self.tree_hist.column(c, width=100, anchor="center")
            elif c in ["Harga", "Total", "P/L"]:
                self.tree_hist.column(c, width=100, anchor="e")
            else:
                self.tree_hist.column(c, width=100, anchor="center")
                
        self.tree_hist.pack(fill="both", expand=True)

    # --- FUNGSI CASH MANAGEMENT ---
    def action_deposit(self):
        try:
            # Karena state='normal', user bisa mengetik teks yang salah format.
            # Menggunakan get_date() akan otomatis mengubah teks yang diketik manual menjadi objek datetime 
            # jika formatnya benar (dd-mm-yyyy), atau menampilkan error bawaan tkcalendar.
            date_fmt = self.cash_date_entry.get_date().strftime("%d-%m-%Y")
            amount = float(self.cash_amount_entry.get().replace(',', '.'))
            
            if amount <= 0:
                messagebox.showwarning("Peringatan", "Nominal setor harus lebih dari 0!")
                return

            new_row = {'Tanggal': date_fmt, 'Tipe': 'DEPOSIT', 'Ticker': 'CASH', 'Jumlah': 1.0, 'Harga': amount, 'Total': amount, 'Realized_PL': 0.0}
            self.history_df = pd.concat([self.history_df, pd.DataFrame([new_row])], ignore_index=True)
            
            self.history_df.to_excel(self.file_name, index=False)
            self.rebuild_portfolio()
            self.update_tables()
            
            self.cash_amount_entry.delete(0, tk.END)
            messagebox.showinfo("Sukses", f"Berhasil menyetor dana sebesar {amount:,.2f}")
        except ValueError:
            messagebox.showerror("Error", "Input nominal harus berupa angka dan format tanggal harus benar (dd-mm-yyyy)!")

    def action_withdraw(self):
        try:
            date_fmt = self.cash_date_entry.get_date().strftime("%d-%m-%Y")
            amount = float(self.cash_amount_entry.get().replace(',', '.'))
            
            if amount <= 0:
                messagebox.showwarning("Peringatan", "Nominal tarik harus lebih dari 0!")
                return

            if self.cash_balance < (amount - 0.000001):
                messagebox.showwarning("Peringatan", "Saldo cash Anda tidak mencukupi untuk ditarik!")
                return

            new_row = {'Tanggal': date_fmt, 'Tipe': 'WITHDRAW', 'Ticker': 'CASH', 'Jumlah': 1.0, 'Harga': amount, 'Total': amount, 'Realized_PL': 0.0}
            self.history_df = pd.concat([self.history_df, pd.DataFrame([new_row])], ignore_index=True)
            
            self.history_df.to_excel(self.file_name, index=False)
            self.rebuild_portfolio()
            self.update_tables()
            
            self.cash_amount_entry.delete(0, tk.END)
            messagebox.showinfo("Sukses", f"Berhasil menarik dana sebesar {amount:,.2f}")
        except ValueError:
            messagebox.showerror("Error", "Input nominal harus berupa angka dan format tanggal harus benar (dd-mm-yyyy)!")

    # --- FUNGSI TRANSAKSI SAHAM ---
    def action_buy(self):
        try:
            date_fmt = self.date_entry.get_date().strftime("%d-%m-%Y")
            ticker = self.ticker_entry.get().upper()
            qty = float(self.qty_entry.get().replace(',', '.'))
            price = float(self.price_entry.get().replace(',', '.'))
            total = qty * price

            if self.cash_balance < (total - 0.000001):
                messagebox.showwarning("Dana Kurang", f"Saldo cash tidak mencukupi!\nTotal pembelian: {total:,.2f}\nSaldo Anda: {self.cash_balance:,.2f}")
                return

            new_row = {'Tanggal': date_fmt, 'Tipe': 'BELI', 'Ticker': ticker, 'Jumlah': qty, 'Harga': price, 'Total': total, 'Realized_PL': 0.0}
            self.history_df = pd.concat([self.history_df, pd.DataFrame([new_row])], ignore_index=True)
            
            self.history_df.to_excel(self.file_name, index=False)
            self.rebuild_portfolio()
            self.update_tables()
            self.clear_inputs()
        except ValueError:
            messagebox.showerror("Error", "Input Jumlah dan Harga harus berupa angka dan format tanggal harus benar (dd-mm-yyyy)!")

    def action_sell(self):
        try:
            date_fmt = self.date_entry.get_date().strftime("%d-%m-%Y")
            ticker = self.ticker_entry.get().upper()
            qty = float(self.qty_entry.get().replace(',', '.'))
            price = float(self.price_entry.get().replace(',', '.'))
            total = qty * price

            if ticker not in self.portfolio or self.portfolio[ticker]['qty'] < (qty - 0.000001):
                messagebox.showwarning("Peringatan", "Saldo saham tidak mencukupi atau kode tidak ditemukan!")
                return

            avg_price = self.portfolio[ticker]['avg_price']
            realized_pl = (price - avg_price) * qty

            new_row = {'Tanggal': date_fmt, 'Tipe': 'JUAL', 'Ticker': ticker, 'Jumlah': qty, 'Harga': price, 'Total': total, 'Realized_PL': realized_pl}
            self.history_df = pd.concat([self.history_df, pd.DataFrame([new_row])], ignore_index=True)
            
            self.history_df.to_excel(self.file_name, index=False)
            self.rebuild_portfolio()
            self.update_tables()
            self.clear_inputs()
        except ValueError:
            messagebox.showerror("Error", "Input Jumlah dan Harga harus berupa angka dan format tanggal harus benar (dd-mm-yyyy)!")

    def action_delete(self):
        selected_items = self.tree_hist.selection()
        if not selected_items:
            messagebox.showinfo("Info", "Silakan klik data di tabel riwayat terlebih dahulu.")
            return

        confirm = messagebox.askyesno(
            "Konfirmasi Hapus", 
            f"Anda memilih {len(selected_items)} data untuk dihapus.\n\nApakah Anda yakin?"
        )
        
        if confirm:
            indices_to_drop = [int(item) for item in selected_items]
            self.history_df = self.history_df.drop(index=indices_to_drop)
            self.history_df = self.history_df.reset_index(drop=True)
            self.history_df.to_excel(self.file_name, index=False)
            
            self.rebuild_portfolio()
            self.update_tables()
            messagebox.showinfo("Sukses", "Data berhasil dihapus!")

    def clear_inputs(self):
        self.ticker_entry.delete(0, tk.END)
        self.qty_entry.delete(0, tk.END)
        self.price_entry.delete(0, tk.END)

    def rebuild_portfolio(self):
        self.portfolio.clear()
        self.cash_balance = 0.0 

        for _, row in self.history_df.iterrows():
            tipe = row['Tipe']
            ticker = row['Ticker']
            total = row['Total']
            jumlah = row['Jumlah']

            if tipe == 'DEPOSIT':
                self.cash_balance += total
            elif tipe == 'WITHDRAW':
                self.cash_balance -= total
            elif tipe == 'BELI':
                self.cash_balance -= total
                if ticker not in self.portfolio:
                    self.portfolio[ticker] = {'qty': 0.0, 'avg_price': 0.0}
                
                current_qty = self.portfolio[ticker]['qty']
                current_avg = self.portfolio[ticker]['avg_price']
                new_total_cost = (current_qty * current_avg) + total
                self.portfolio[ticker]['qty'] += jumlah
                
                if self.portfolio[ticker]['qty'] > 0:
                    self.portfolio[ticker]['avg_price'] = new_total_cost / self.portfolio[ticker]['qty']
                    
            elif tipe == 'JUAL':
                self.cash_balance += total
                if ticker not in self.portfolio:
                    self.portfolio[ticker] = {'qty': 0.0, 'avg_price': 0.0}
                    
                self.portfolio[ticker]['qty'] -= jumlah
                if self.portfolio[ticker]['qty'] < 0.000001:
                    self.portfolio[ticker]['qty'] = 0.0
                    self.portfolio[ticker]['avg_price'] = 0.0

    def update_tables(self):
        self.lbl_cash.config(text=f"Saldo Cash: {self.cash_balance:,.2f}")

        for item in self.tree_port.get_children():
            self.tree_port.delete(item)
        for ticker, data in self.portfolio.items():
            if data['qty'] > 0.000001:
                self.tree_port.insert("", "end", values=(
                    ticker, 
                    f"{data['qty']:.4f}".rstrip('0').rstrip('.'), 
                    f"{data['avg_price']:,.2f}", 
                    f"{(data['qty'] * data['avg_price']):,.2f}"
                ))

        for item in self.tree_hist.get_children():
            self.tree_hist.delete(item)
            
        for index, row in self.history_df.iterrows():
            if row['Tipe'] in ['DEPOSIT', 'WITHDRAW']:
                qty_str = "-"
                pl_str = "-"
            else:
                qty_str = f"{row['Jumlah']:.4f}".rstrip('0').rstrip('.')
                pl_str = f"{row['Realized_PL']:,.2f}" if row['Tipe'] == 'JUAL' else "-"

            self.tree_hist.insert("", "end", iid=str(index), values=(
                row['Tanggal'], row['Tipe'], row['Ticker'], 
                qty_str, 
                f"{row['Harga']:,.2f}", f"{row['Total']:,.2f}", pl_str
            ))

    def save_to_excel(self):
        if self.history_df.empty:
            messagebox.showinfo("Info", "Tidak ada data transaksi untuk disimpan.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Simpan File Portofolio"
        )
        if file_path:
            try:
                self.history_df.to_excel(file_path, index=False)
                messagebox.showinfo("Sukses", f"Data berhasil disimpan ke:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menyimpan file:\n{e}")

    def load_from_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Buka File Portofolio"
        )
        if file_path:
            try:
                self.history_df = pd.read_excel(file_path)
                self.rebuild_portfolio()
                self.update_tables()
                messagebox.showinfo("Sukses", "Data portofolio berhasil dimuat.")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal membaca file:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PortfolioApp(root)
    root.mainloop()