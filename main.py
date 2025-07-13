import tkinter as tk
from tkinter import messagebox
import tkinter.font as tkFont
import sqlite3
import pandas as pd
import os
import sys

# Ensure data folder exists
os.makedirs("data", exist_ok=True)

# ----------------- Database Setup -----------------
conn = sqlite3.connect(os.path.join("data", "marksheet.db"))
cursor = conn.cursor()
cursor.execute("""
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        section TEXT,
        maths INTEGER,
        english INTEGER,
        science INTEGER,
        hindi INTEGER,
        sst INTEGER,
        total INTEGER,
        percentage REAL
    )
""")
conn.commit()

# ----------------- Main App Window -----------------
root = tk.Tk()
root.title("Student Marksheet System")
root.geometry("850x750")
root.configure(bg="#F3F6F9")

headerFont = tkFont.Font(family="Helvetica", size=18, weight="bold")
labelFont = tkFont.Font(family="Helvetica", size=12)

tk.Label(root, text="Student Marksheet Management", font=headerFont, bg="#F3F6F9", fg="#333").pack(pady=15)

# ----------------- Form Frame -----------------
form_frame = tk.LabelFrame(root, text="Enter Student Details", bg="#F3F6F9", font=labelFont)
form_frame.pack(padx=20, pady=10, fill="x")

tk.Label(form_frame, text="Student ID:", bg="#F3F6F9", font=labelFont).grid(row=0, column=0, padx=10, pady=5, sticky="w")
id_entry = tk.Entry(form_frame, width=30)
id_entry.grid(row=0, column=1, padx=10)

tk.Label(form_frame, text="Name:", bg="#F3F6F9", font=labelFont).grid(row=1, column=0, padx=10, pady=5, sticky="w")
name = tk.Entry(form_frame, width=30)
name.grid(row=1, column=1, padx=10)

tk.Label(form_frame, text="Section:", bg="#F3F6F9", font=labelFont).grid(row=2, column=0, padx=10, pady=5, sticky="w")
section = tk.Entry(form_frame, width=30)
section.grid(row=2, column=1, padx=10)

subjects = ["Maths", "English", "Science", "Hindi", "SST"]
entries = []
for i, subj in enumerate(subjects):
    tk.Label(form_frame, text=subj + ":", bg="#F3F6F9", font=labelFont).grid(row=i+3, column=0, padx=10, pady=5, sticky="w")
    e = tk.Entry(form_frame, width=30)
    e.grid(row=i+3, column=1, padx=10)
    entries.append(e)

s1, s2, s3, s4, s5 = entries

# ----------------- Table Frame -----------------
table_frame = tk.LabelFrame(root, text="Student Records", bg="#F3F6F9", font=labelFont)
table_frame.pack(padx=20, pady=10, fill="both", expand=True)

table = tk.Text(table_frame, height=10, width=100, font=("Courier", 10))
table.pack(side=tk.LEFT, padx=10, pady=5, fill="both", expand=True)

scroll = tk.Scrollbar(table_frame, command=table.yview)
scroll.pack(side=tk.RIGHT, fill=tk.Y)
table.config(yscrollcommand=scroll.set)

# ----------------- Functions -----------------
def clearForm():
    id_entry.delete(0, tk.END)
    name.delete(0, tk.END)
    section.delete(0, tk.END)
    for entry in entries:
        entry.delete(0, tk.END)
    table.delete('1.0', tk.END)

def getMarks():
    return [s1.get(), s2.get(), s3.get(), s4.get(), s5.get()]

def totalMarks(marks):
    return sum(map(int, marks))

def percentage(total):
    return round(total / 5, 2)

def validate(marks):
    try:
        return all(0 <= int(m) <= 100 for m in marks)
    except:
        return False

def submit():
    student_name = name.get()
    student_sec = section.get()
    marks = getMarks()

    if not student_name or not student_sec or not all(marks):
        messagebox.showerror("Error", "All fields are required.")
        return
    if not validate(marks):
        messagebox.showerror("Error", "Marks must be integers between 0-100.")
        return

    marks = list(map(int, marks))
    total = totalMarks(marks)
    per = percentage(total)

    cursor.execute("INSERT INTO students (name, section, maths, english, science, hindi, sst, total, percentage) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                   (student_name, student_sec, *marks, total, per))
    conn.commit()
    messagebox.showinfo("Success", "Student added.")
    clearForm()

def showRecords():
    table.delete('1.0', tk.END)
    rows = cursor.execute("SELECT id, name, section, total, percentage FROM students").fetchall()
    if not rows:
        table.insert(tk.END, "No records found.\n")
        return
    table.insert(tk.END, f"{'ID':<5}{'Name':<20}{'Section':<10}{'Total':<10}{'Percentage':<10}\n")
    table.insert(tk.END, "-"*60 + "\n")
    for r in rows:
        table.insert(tk.END, f"{r[0]:<5}{r[1]:<20}{r[2]:<10}{r[3]:<10}{r[4]:<10.2f}\n")

def searchByName():
    table.delete('1.0', tk.END)
    query = name.get()
    if not query:
        messagebox.showerror("Error", "Enter a name to search.")
        return
    rows = cursor.execute("SELECT * FROM students WHERE name LIKE ?", ('%' + query + '%',)).fetchall()
    if not rows:
        table.insert(tk.END, "No student found with that name.\n")
        return
    for r in rows:
        table.insert(tk.END, f"ID: {r[0]}, Name: {r[1]}, Section: {r[2]}, Marks: {r[3:8]}, Total: {r[8]}, %: {r[9]}\n")

def deleteRecord():
    sid = id_entry.get()
    if not sid.isdigit():
        messagebox.showerror("Error", "Enter a valid ID to delete.")
        return
    cursor.execute("DELETE FROM students WHERE id = ?", (int(sid),))
    conn.commit()
    messagebox.showinfo("Deleted", f"Record with ID {sid} deleted.")
    clearForm()

def updateRecord():
    sid = id_entry.get()
    if not sid.isdigit():
        messagebox.showerror("Error", "Enter a valid ID to update.")
        return
    marks = getMarks()
    if not validate(marks):
        return messagebox.showerror("Error", "Invalid marks.")
    marks = list(map(int, marks))
    total = totalMarks(marks)
    per = percentage(total)

    cursor.execute("UPDATE students SET maths=?, english=?, science=?, hindi=?, sst=?, total=?, percentage=? WHERE id=?",
                   (*marks, total, per, int(sid)))
    conn.commit()
    messagebox.showinfo("Updated", f"Record ID {sid} updated.")

def exportToExcel():
    data = cursor.execute("SELECT * FROM students").fetchall()
    if not data:
        messagebox.showinfo("No Data", "No records to export.")
        return
    columns = ["ID", "Name", "Section", "Maths", "English", "Science", "Hindi", "SST", "Total", "Percentage"]
    df = pd.DataFrame(data, columns=columns)
    os.makedirs("exports", exist_ok=True)
    file_path = os.path.join("exports", "student_marks.xlsx")
    df.to_excel(file_path, index=False)
    messagebox.showinfo("Exported", f"Exported to {file_path}")
    try:
        os.startfile(file_path)
    except Exception:
        if sys.platform.startswith("darwin"):
            os.system(f"open {file_path}")
        elif sys.platform.startswith("linux"):
            os.system(f"xdg-open {file_path}")

# ----------------- Buttons -----------------
btn_frame = tk.Frame(root, bg="#F3F6F9")
btn_frame.pack(pady=20)

tk.Button(btn_frame, text="Add Student", command=submit, bg="#2ecc71", fg="white", width=15, font=labelFont).grid(row=0, column=0, padx=10, pady=5)
tk.Button(btn_frame, text="View Records", command=showRecords, bg="#1abc9c", fg="white", width=15, font=labelFont).grid(row=0, column=1, padx=10, pady=5)
tk.Button(btn_frame, text="Search by Name", command=searchByName, bg="#3498db", fg="white", width=15, font=labelFont).grid(row=1, column=0, padx=10, pady=5)
tk.Button(btn_frame, text="Delete by ID", command=deleteRecord, bg="#e74c3c", fg="white", width=15, font=labelFont).grid(row=1, column=1, padx=10, pady=5)
tk.Button(btn_frame, text="Update by ID", command=updateRecord, bg="#f39c12", fg="white", width=15, font=labelFont).grid(row=2, column=0, padx=10, pady=5)
tk.Button(btn_frame, text="Clear Form", command=clearForm, bg="#bdc3c7", width=15, font=labelFont).grid(row=2, column=1, padx=10, pady=5)
tk.Button(btn_frame, text="Export to Excel", command=exportToExcel, bg="#9b59b6", fg="white", width=32, font=labelFont).grid(row=3, column=0, columnspan=2, pady=10)

# ----------------- Start App -----------------
root.mainloop()
