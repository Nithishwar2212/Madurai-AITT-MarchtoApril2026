import matplotlib
matplotlib.use("TkAgg") 
import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
import os

# -- Theme & Style --
ctk.set_appearance_mode("Dark")
COLOR_BG = "#1A1A1B"   
COLOR_PANEL = "#272729" 
COLOR_ACCENT = "#7868E6"
COLOR_INCOME = "#5FAD56"
COLOR_EXPENSE = "#E76F51"
COLOR_TEXT_MAIN = "#FFFFFF"

class ProExpenseTracker(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Expenses Management")
        self.geometry("1450x850") 
        self.configure(fg_color=COLOR_BG)

        self.file_path = "expenses.xlsx"
        self.init_excel()
        
        self.editing_index = None
        self.current_bill_path = "No file attached"
        self.monthly_limit = 30000.0

        # --- MAIN CONTAINER ---
        self.main_container = ctk.CTkFrame(self, fg_color=COLOR_BG, corner_radius=0)
        self.main_container.pack(fill="both", expand=True)
        
        self.main_container.grid_columnconfigure(0, weight=2) 
        self.main_container.grid_columnconfigure(1, weight=3) 
        self.main_container.grid_columnconfigure(2, weight=2) 
        self.main_container.grid_rowconfigure(0, weight=1)

        self.setup_ui()
        # Initial refresh to load existing history
        self.refresh_ui()

    def init_excel(self):
        if not os.path.exists(self.file_path):
            df = pd.DataFrame(columns=["Description", "Amount", "Category", "Date", "Type", "Bill"])
            df.to_excel(self.file_path, index=False)

    def setup_ui(self):
        # COLUMN 1: ENTRY PANEL
        self.left_scroll = ctk.CTkScrollableFrame(self.main_container, fg_color=COLOR_PANEL, corner_radius=0, 
                                                 border_width=1, border_color="#333", label_text="Entry Panel")
        self.left_scroll.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        
        self.desc_e = self.create_input(self.left_scroll, "Description")
        self.amt_e = self.create_input(self.left_scroll, "Amount (₹)")
        
        ctk.CTkLabel(self.left_scroll, text="Category", font=("Arial", 14)).pack(anchor="w", padx=40, pady=(15, 0))
        self.cat_options = ["Food", "Rent", "Bills", "Travel", "Others"]
        self.cat_o = ctk.CTkOptionMenu(self.left_scroll, values=self.cat_options, 
                                       command=self.check_category, fg_color="#3D3D3D", height=42, corner_radius=10)
        self.cat_o.pack(fill="x", padx=40, pady=10)

        self.other_cat_e = ctk.CTkEntry(self.left_scroll, placeholder_text="Specify category name...", height=42, corner_radius=10, border_color=COLOR_ACCENT)

        ctk.CTkLabel(self.left_scroll, text="Date", font=("Arial", 14)).pack(anchor="w", padx=40, pady=(15, 0))
        self.date_p = DateEntry(self.left_scroll, date_pattern='yyyy-mm-dd')
        self.date_p.pack(pady=10, padx=40, fill="x")

        self.bill_btn = ctk.CTkButton(self.left_scroll, text="📎 Attach Bill", fg_color="#3D3D3D", command=self.attach_file)
        self.bill_btn.pack(pady=15, padx=40, fill="x")
        
        self.save_btn = ctk.CTkButton(self.left_scroll, text="Add", font=("Arial", 14, "bold"), 
                                      fg_color=COLOR_ACCENT, height=52, command=self.handle_save)
        self.save_btn.pack(pady=30, padx=40, fill="x")

        # COLUMN 2: HISTORY & CHART
        m_col = ctk.CTkFrame(self.main_container, fg_color="transparent")
        m_col.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        m_col.grid_rowconfigure(0, weight=1); m_col.grid_rowconfigure(1, weight=1)

        self.hist_scroll = ctk.CTkScrollableFrame(m_col, fg_color=COLOR_BG, corner_radius=0, label_text="Recent History", border_width=1, border_color="#333")
        self.hist_scroll.grid(row=0, column=0, sticky="nsew")

        self.chart_panel = ctk.CTkFrame(m_col, fg_color=COLOR_BG, corner_radius=0, border_width=1, border_color="#333")
        self.chart_panel.grid(row=1, column=0, sticky="nsew")
        
        self.fig, self.ax = plt.subplots(figsize=(5, 4), dpi=100, facecolor=COLOR_BG)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_panel)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=25, pady=25)

        # COLUMN 3: STATS
        r_col = ctk.CTkFrame(self.main_container, fg_color=COLOR_PANEL, corner_radius=0, border_width=1, border_color="#333")
        r_col.grid(row=0, column=2, sticky="nsew", padx=0, pady=0)

        ctk.CTkLabel(r_col, text="Monthly Status", font=("Arial", 22, "bold")).pack(pady=(55, 20))
        self.p_bar = ctk.CTkProgressBar(r_col, progress_color=COLOR_ACCENT, height=16)
        self.p_bar.pack(fill="x", padx=55, pady=12)
        self.stats_lbl = ctk.CTkLabel(r_col, text="₹0 / ₹30,000", font=("Arial", 15))
        self.stats_lbl.pack(pady=(0, 35))

        ctk.CTkLabel(r_col, text="Top Spending", font=("Arial", 22, "bold")).pack(pady=25)
        self.rank_container = ctk.CTkFrame(r_col, fg_color="transparent")
        self.rank_container.pack(fill="both", expand=True, padx=55, pady=12)

    def check_category(self, choice):
        if choice == "Others":
            self.other_cat_e.pack(fill="x", padx=40, pady=(0, 10), after=self.cat_o)
        else:
            self.other_cat_e.pack_forget()

    def handle_save(self):
        try:
            selected_cat = self.cat_o.get()
            final_cat = self.other_cat_e.get() if selected_cat == "Others" else selected_cat
            item_type = "Income" if selected_cat == "Salary" else "Expense"
            
            df = pd.read_excel(self.file_path)
            data = {
                "Description": self.desc_e.get(),
                "Amount": float(self.amt_e.get()),
                "Category": final_cat, 
                "Date": self.date_p.get_date().strftime('%Y-%m-%d'),
                "Type": item_type,
                "Bill": self.current_bill_path
            }

            if self.editing_index is not None:
                for key, value in data.items():
                    df.at[self.editing_index, key] = value
                self.editing_index = None
                self.save_btn.configure(text="Add", fg_color=COLOR_ACCENT)
            else:
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)

            df.to_excel(self.file_path, index=False)
            self.refresh_ui()
            self.reset_form()
            
        except Exception as e:
            messagebox.showerror("Error", f"Action failed: {e}")

    def refresh_ui(self):
        df = pd.read_excel(self.file_path)
        df['Date'] = pd.to_datetime(df['Date'])
        
        for w in self.hist_scroll.winfo_children(): w.destroy()
        
        for i, r in df.tail(15).iloc[::-1].iterrows():
            f = ctk.CTkFrame(self.hist_scroll, fg_color="#333", height=55, corner_radius=6)
            f.pack(fill="x", pady=6, padx=12)
            f.pack_propagate(False)
            
            ctk.CTkLabel(f, text=r['Description'], font=("Arial", 13)).pack(side="left", padx=15)
            ctk.CTkButton(f, text="🗑", width=32, fg_color=COLOR_EXPENSE, command=lambda idx=i: self.delete_item(idx)).pack(side="right", padx=5)
            ctk.CTkButton(f, text="✏️", width=32, fg_color="#444", command=lambda idx=i, row=r: self.start_edit(idx, row)).pack(side="right", padx=5)
            ctk.CTkLabel(f, text=f"₹{r['Amount']}", text_color=COLOR_TEXT_MAIN, font=("Arial", 13, "bold")).pack(side="right", padx=15)

        curr_df = df[(df['Date'].dt.month == datetime.now().month) & (df['Type'] == 'Expense')]
        spent = curr_df['Amount'].sum()
        self.p_bar.set(min(spent / self.monthly_limit, 1.0))
        self.stats_lbl.configure(text=f"Spent: ₹{spent:,.0f} / ₹{self.monthly_limit:,.0f}")
        self.update_chart(curr_df)
        self.update_rankings(curr_df)

    def update_chart(self, df):
        self.ax.clear()
        if not df.empty:
            data = df.groupby('Category')['Amount'].sum()
            self.ax.pie(data, labels=data.index, autopct='%1.0f%%', colors=[COLOR_ACCENT, "#00DDEB", COLOR_EXPENSE], textprops={'color':"w", 'size':11}, wedgeprops={'width':0.45, 'edgecolor':COLOR_BG})
        self.canvas.draw()

    def update_rankings(self, df):
        for w in self.rank_container.winfo_children(): w.destroy()
        if df.empty: return
        ranks = df.groupby('Category')['Amount'].sum().sort_values(ascending=False)
        for cat, val in ranks.items():
            r = ctk.CTkFrame(self.rank_container, fg_color="transparent")
            r.pack(fill="x", pady=6)
            ctk.CTkLabel(r, text=cat, font=("Arial", 14)).pack(side="left")
            ctk.CTkLabel(r, text=f"₹{val:,.0f}", font=("Arial", 14, "bold"), text_color=COLOR_EXPENSE).pack(side="right")

    def start_edit(self, idx, row):
        self.editing_index = idx
        self.reset_form()
        self.desc_e.insert(0, row['Description'])
        self.amt_e.insert(0, str(row['Amount']))
        
        cat = row['Category']
        if cat in self.cat_options:
            self.cat_o.set(cat)
        else:
            self.cat_o.set("Others")
            self.other_cat_e.pack(fill="x", padx=40, pady=(0, 10), after=self.cat_o)
            self.other_cat_e.insert(0, cat)
            
        self.date_p.set_date(datetime.strptime(str(row['Date']).split(' ')[0], '%Y-%m-%d'))
        self.current_bill_path = row['Bill']
        self.save_btn.configure(text="Update Entry", fg_color="#FFB100")

    def delete_item(self, idx):
        if messagebox.askyesno("Confirm", "Delete this entry?"):
            df = pd.read_excel(self.file_path)
            df = df.drop(idx).reset_index(drop=True)
            df.to_excel(self.file_path, index=False)
            self.refresh_ui()

    def attach_file(self):
        p = filedialog.askopenfilename()
        if p: self.current_bill_path = p; self.bill_btn.configure(text="Attached ✅", fg_color=COLOR_INCOME)

    def create_input(self, p, lbl):
        ctk.CTkLabel(p, text=lbl, font=("Arial", 14)).pack(anchor="w", padx=40, pady=(20, 0))
        e = ctk.CTkEntry(p, height=42, corner_radius=10, border_color="#444"); e.pack(fill="x", padx=40, pady=8)
        return e

    def reset_form(self):
        self.desc_e.delete(0, 'end'); self.amt_e.delete(0, 'end')
        self.other_cat_e.delete(0, 'end'); self.other_cat_e.pack_forget()
        self.cat_o.set("Food")
        self.bill_btn.configure(text="📎 Attach Bill", fg_color="#3D3D3D")
        self.current_bill_path = "No file attached"

if __name__ == "__main__":
    app = ProExpenseTracker()
    app.mainloop()
