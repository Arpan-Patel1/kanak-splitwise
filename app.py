import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import pyodbc
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

class EmployeeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Employee Management System")
        self.root.geometry("400x300")
        
        # Create main menu
        self.create_main_menu()
        
        # Excel file path
        self.excel_file = "Employee_Userform 2.xlsx"
        
        # Create Excel file if it doesn't exist
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame(columns=["EmpID", "Name", "Department", "Salary"])
            df.to_excel(self.excel_file, sheet_name="Sheet1", index=False)
    
    def create_main_menu(self):
        # Clear any existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create main menu buttons
        frame = tk.Frame(self.root)
        frame.pack(expand=True)
        
        btn_add = tk.Button(frame, text="Add Employee", width=20, height=2, command=self.show_add_employee_form)
        btn_add.pack(pady=10)
        
        btn_pivot = tk.Button(frame, text="Create Pivot Table", width=20, height=2, command=self.create_pivot_table)
        btn_pivot.pack(pady=10)
        
        btn_chart = tk.Button(frame, text="Create Pivot Chart", width=20, height=2, command=self.create_pivot_chart)
        btn_chart.pack(pady=10)
        
        btn_delete = tk.Button(frame, text="Delete Employee", width=20, height=2, command=self.show_delete_employee_form)
        btn_delete.pack(pady=10)
    
    def show_add_employee_form(self):
        # Clear any existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create form
        frame = tk.Frame(self.root)
        frame.pack(expand=True)
        
        # Labels
        tk.Label(frame, text="Employee ID:").grid(row=0, column=0, sticky="w", pady=5)
        tk.Label(frame, text="Name:").grid(row=1, column=0, sticky="w", pady=5)
        tk.Label(frame, text="Department:").grid(row=2, column=0, sticky="w", pady=5)
        tk.Label(frame, text="Salary:").grid(row=3, column=0, sticky="w", pady=5)
        
        # Entry fields
        self.emp_id_entry = tk.Entry(frame)
        self.emp_id_entry.grid(row=0, column=1, pady=5)
        
        self.name_entry = tk.Entry(frame)
        self.name_entry.grid(row=1, column=1, pady=5)
        
        self.dept_entry = tk.Entry(frame)
        self.dept_entry.grid(row=2, column=1, pady=5)
        
        self.salary_entry = tk.Entry(frame)
        self.salary_entry.grid(row=3, column=1, pady=5)
        
        # Buttons
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        submit_btn = tk.Button(btn_frame, text="Submit", command=self.save_employee_data)
        submit_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = tk.Button(btn_frame, text="Cancel", command=self.create_main_menu)
        cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def save_employee_data(self):
        try:
            # Validate input fields
            if not all([self.emp_id_entry.get().strip(), self.name_entry.get().strip(), 
                        self.dept_entry.get().strip(), self.salary_entry.get().strip()]):
                messagebox.showwarning("Validation Error", "Please fill in all fields before submitting.")
                return
            
            # Validate numeric fields
            try:
                emp_id = int(self.emp_id_entry.get())
                salary = float(self.salary_entry.get())
            except ValueError:
                messagebox.showwarning("Validation Error", "EmpID and Salary must be numbers.")
                return
            
            # Read existing data
            df = pd.read_excel(self.excel_file, sheet_name="Sheet1")
            
            # Check for duplicate EmpID in Excel
            if emp_id in df['EmpID'].values:
                messagebox.showwarning("Duplicate Error", "EmpID already exists in Excel. Please enter a unique EmpID.")
                return
            
            # Connect to SQL Server and check for duplicates
            try:
                conn_str = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=LIN-5CG0523B24\SQLEXPRESS;DATABASE=EmployeeDB;Trusted_Connection=yes;"
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                
                # Check for duplicate EmpID
                cursor.execute(f"SELECT COUNT(*) FROM Employee WHERE ID = {emp_id}")
                if cursor.fetchone()[0] > 0:
                    messagebox.showwarning("Duplicate Error", "EmpID already exists in database. Please enter a unique EmpID.")
                    conn.close()
                    return
                
                # Insert into database
                name = self.name_entry.get().replace("'", "''")
                dept = self.dept_entry.get().replace("'", "''")
                cursor.execute(f"INSERT INTO Employee (ID, Name, Department, Salary) VALUES ({emp_id}, '{name}', '{dept}', {salary})")
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("Database Error", f"Failed to connect to database: {str(e)}")
                return
            
            # Add to DataFrame
            new_row = pd.DataFrame({
                'EmpID': [emp_id],
                'Name': [self.name_entry.get()],
                'Department': [self.dept_entry.get()],
                'Salary': [salary]
            })
            
            df = pd.concat([df, new_row], ignore_index=True)
            
            # Save to Excel
            df.to_excel(self.excel_file, sheet_name="Sheet1", index=False)
            
            messagebox.showinfo("Success", "Employee details saved successfully!")
            self.create_main_menu()
            
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
    
    def show_delete_employee_form(self):
        # Clear any existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create form
        frame = tk.Frame(self.root)
        frame.pack(expand=True)
        
        # Labels and Entry
        tk.Label(frame, text="Enter Employee ID to Delete:").grid(row=0, column=0, sticky="w", pady=10)
        self.delete_emp_id_entry = tk.Entry(frame)
        self.delete_emp_id_entry.grid(row=0, column=1, pady=10)
        
        # Buttons
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        delete_btn = tk.Button(btn_frame, text="Delete", command=self.delete_employee)
        delete_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = tk.Button(btn_frame, text="Cancel", command=self.create_main_menu)
        cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def delete_employee(self):
        try:
            emp_id = self.delete_emp_id_entry.get().strip()
            
            if not emp_id:
                messagebox.showwarning("Input Error", "Please enter an Employee ID.")
                return
            
            # Read Excel data
            df = pd.read_excel(self.excel_file, sheet_name="Sheet1")
            
            # Check if record exists in Excel
            if int(emp_id) not in df['EmpID'].values:
                messagebox.showwarning("Not Found", f"Record with EmpID {emp_id} does not exist in Excel!")
                return
            
            # Delete from Excel
            df = df[df['EmpID'] != int(emp_id)]
            df.to_excel(self.excel_file, sheet_name="Sheet1", index=False)
            
            # Delete from SQL Server
            try:
                conn_str = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=LIN-5CG21821NR;DATABASE=EmployeeDB;Trusted_Connection=yes;"
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                
                cursor.execute(f"DELETE FROM Employee WHERE ID = {int(emp_id)}")
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("Database Error", f"Failed to delete from database: {str(e)}")
                return
            
            messagebox.showinfo("Success", f"Record with EmpID {emp_id} deleted successfully from both Excel and database!")
            self.create_main_menu()
            
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
    
    def create_pivot_table(self):
        try:
            # Read data from Excel
            df = pd.read_excel(self.excel_file, sheet_name="Sheet1")
            
            if df.empty:
                messagebox.showwarning("No Data", "No data available to create pivot table.")
                return
            
            # Create pivot table
            pivot = pd.pivot_table(df, values='EmpID', index=['Department'], aggfunc='count')
            pivot.columns = ['Employee Count']
            
            # Save to new sheet
            with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='a') as writer:
                # Check if sheet exists and remove it
                if 'PivotTableSheet' in writer.book.sheetnames:
                    idx = writer.book.sheetnames.index('PivotTableSheet')
                    writer.book.remove(writer.book.worksheets[idx])
                    
                # Write pivot table
                pivot.to_excel(writer, sheet_name='PivotTableSheet')
            
            messagebox.showinfo("Success", "Pivot table created successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create pivot table: {str(e)}")
    
    def create_pivot_chart(self):
        try:
            # First ensure pivot table exists
            self.create_pivot_table()
            
            # Load workbook
            wb = load_workbook(self.excel_file)
            
            # Check if pivot sheet exists
            if 'PivotTableSheet' not in wb.sheetnames:
                messagebox.showwarning("No Pivot Table", "Pivot table not found. Please create a pivot table first.")
                return
            
            # Check if chart sheet exists and create/clear it
            if 'PivotChartSheet' in wb.sheetnames:
                idx = wb.sheetnames.index('PivotChartSheet')
                wb.remove(wb.worksheets[idx])
            
            # Create chart sheet
            chart_sheet = wb.create_sheet('PivotChartSheet')
            
            # Get pivot data
            pivot_sheet = wb['PivotTableSheet']
            
            # Determine data range
            last_row = pivot_sheet.max_row
            
            # Create chart
            chart = BarChart()
            chart.title = "Employee Count by Department"
            chart.y_axis.title = "Count"
            chart.x_axis.title = "Department"
            
            # Add data to chart
            data = Reference(pivot_sheet, min_col=2, min_row=1, max_row=last_row, max_col=2)
            cats = Reference(pivot_sheet, min_col=1, min_row=2, max_row=last_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            
            # Add chart to sheet
            chart_sheet.add_chart(chart, "B5")
            
            # Save workbook
            wb.save(self.excel_file)
            
            messagebox.showinfo("Success", "Pivot chart created successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create pivot chart: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeApp(root)
    root.mainloop()
