import Tkinter as tk
from SUMXL import SUMXL

def manage_entry_fields():
    SUMXL(e1.get(),e2.get(),e3.get())
    e1.delete(0, tk.END)
    e2.delete(0, tk.END)

master = tk.Tk()
tk.Label(master, text="First File Name").grid(row=0)
tk.Label(master, text="Last File Name").grid(row=1)
tk.Label(master, text="Output File Name").grid(row=2)

e1 = tk.Entry(master)
e2 = tk.Entry(master)
e3=tk.Entry(master)


e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2, column=1)

tk.Button(master, 
          text='Quit', 
          command=master.quit).grid(row=4, 
                                    column=0, 
                                    sticky=tk.W, 
                                    pady=4)
tk.Button(master, text='Sum', command=manage_entry_fields).grid(row=4, 
                                                               column=1, 
                                                               sticky=tk.W, 
                                                               pady=4)

master.mainloop()

tk.mainloop()
