#pip install openpyxl
#pip install xlrd
#pip install tkcalendar
from tkinter import *
import tkinter as tk
from tkinter import messagebox 
import openpyxl ,xlrd
from openpyxl import Workbook
import pathlib
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
import sys
import os

#window setup
window = tk.Tk()
window.title("BatuKhan Meal Tracker")
window.geometry("273x260")
window.resizable(False,False)
window.configure(bg="#326273")

#excel configuration
file=pathlib.Path('BatuKhanData.xlsx')
if file.exists():
    pass
else:
    file=Workbook()
    sheet=file.active
    sheet['A1']="Ngày"
    sheet['B1']="Thức ăn"
    sheet['C1']="Calorie"
    sheet['D1']="Đạm"
    sheet['E1']="Chết béo"
    sheet['F1']="Tinh bột"
    sheet['G1']="Nhận xét"

    file.save('BatuKhanData.xlsx')

#labels
date = tk.Label(window, text = "Ngày",  bg="#326273")
food = tk.Label(window, text = "Tên đồ ăn",  bg="#326273")
calorie = tk.Label(window, text = "Calorie",  bg="#326273")
protein = tk.Label(window, text = "Đạm",  bg="#326273")
fat = tk.Label(window, text = "Béo",  bg="#326273")
carb = tk.Label(window, text = "Tinh bột",  bg="#326273")
nxet = tk.Label(window, text = "Đánh giá",  bg="#326273")

date.grid(row=0, column=0, padx=5, pady=5, sticky=W)
food.grid(row=1, column=0, padx=5, pady=5, sticky=W)
calorie.grid(row=2, column=0, padx=5, pady=5, sticky=W)
protein.grid(row=3, column=0, padx=5, pady=5, sticky=W)
fat.grid(row=4, column=0, padx=5, pady=5, sticky=W)
carb.grid(row=5, column=0, padx=5, pady=5, sticky=W)
nxet.grid(row=6, column=0, padx=5, pady=5, sticky=W)

#entry
dateValue = StringVar()
foodValue = StringVar()
calorieValue = StringVar()
proteinValue = StringVar()
fatValue = StringVar()
carbValue = StringVar()
nxetValue = StringVar()

date_entry = DateEntry(window, textvariable=dateValue, width=23)
food_entry = tk.Entry(window, textvariable=foodValue, width=25)
calorie_entry = tk.Entry(window, textvariable=calorieValue, width=25)
protein_entry = tk.Entry(window, textvariable=proteinValue, width=25)
fat_entry = tk.Entry(window, textvariable=fatValue, width=25)
carb_entry = tk.Entry(window, textvariable=carbValue, width=25)
nxet_entry = tk.Entry(window, textvariable=nxetValue, width=25)


date_entry.grid(row=0, column=1, padx=5, pady=5)
food_entry.grid(row=1, column=1, padx=5, pady=5)
calorie_entry.grid(row=2, column=1, padx=5, pady=5)
protein_entry.grid(row=3, column=1, padx=5, pady=5)
fat_entry.grid(row=4, column=1, padx=5, pady=5)
carb_entry.grid(row=5, column=1, padx=5, pady=5)
nxet_entry.grid(row=6, column=1, padx=5, pady=5)

#functions
def submit():
    #pull user entries
    date = date_entry.get()
    if date_entry.get():
        pass
    else:
        messagebox.showerror("showerror", "Thiếu ngày")
        sys.exit("Missing Data")
    food = food_entry.get()
    if food_entry.get():
        pass
    else:
        messagebox.showerror("showerror", "Thiếu tên đồ ăn")
        sys.exit("Missing Data")
    calorie = calorie_entry.get()
    if calorie_entry.get():
        pass
    else:
        messagebox.showerror("showerror", "Thiếu calorie")
        sys.exit("Missing Data")
    protein = protein_entry.get()
    if protein_entry.get():
        pass
    else:
        messagebox.showerror("showerror", "Thiếu đạm")
        sys.exit("Missing Data")
    fat = fat_entry.get()
    if fat_entry.get():
        pass
    else:
        messagebox.showerror("showerror", "Thiếu chất béo")
        sys.exit("Missing Data")
    carb = carb_entry.get()
    if carb_entry.get():
        pass
    else:
        messagebox.showerror("showerror", "Thiếu tinh bột")
        sys.exit("Missing Data")
    nxet = nxet_entry.get()
    if nxet_entry.get():
        pass
    else:
        messagebox.showerror("showerror", "Thiếu nhận xét")
        sys.exit("Missing Data")
        
    #enter data in excel file
    file=openpyxl.load_workbook('BatuKhanData.xlsx')
    sheet=file.active
    sheet.append([date,food,calorie,protein,fat,carb,nxet])
    file.save(r'BatuKhanData.xlsx')

    dateValue.set('')
    foodValue.set('')
    calorieValue.set('')
    proteinValue.set('')
    fatValue.set('')
    carbValue.set('')
    nxetValue.set('')

def clear():
    date_entry.delete(0, tk.END)
    food_entry.delete(0, tk.END)
    calorie_entry.delete(0, tk.END)
    protein_entry.delete(0, tk.END)
    fat_entry.delete(0, tk.END)
    carb_entry.delete(0, tk.END)
    nxet_entry.delete(0, tk.END)

    
#Buttons
submit_button = tk.Button(window, text="Submit", command=submit, bg="#326273")
clear_button = tk.Button(window, text="Clear", command=clear, bg="#326273")

submit_button.grid(row=7, column=0, padx=5, pady=5)
clear_button.grid(row=7, column=1, padx=5, pady=5)

window.mainloop()
