import pandas as pd
import os
import time
import tkinter as tk
from csv import reader

path = os.getcwd()
def show_entry_fields():
    name = e1.get()
    row_num = int(e2.get())
    op = e3.get()
    output = pd.DataFrame(columns=['Folder','File','Time','Row'])
    os.chdir(path)
    for folder in os.listdir():
        if folder.find('.') == -1:
            print(folder)
            os.chdir(path+'\\'+folder)
            for file in os.listdir():
                if name in file:
                    try:
                        final=[]
                        with open(file) as f:
                            row = next((x for i, x in enumerate(reader(f)) if i ==row_num-1), None)
                        crt_dt = time.ctime(os.path.getctime(file))
                        fol = folder
                        final.append(fol)
                        final.append(file)
                        final.append(crt_dt)
                        final.append(row)
                        output.loc[len(output)] = final
                        os.chdir(path)
                    except:
                        os.chdir(path)
    os.chdir(path)
    output.to_excel(op +'.xlsx')
    

master = tk.Tk()
master.title('row extractor')
tk.Label(master, text="File Name").grid(row=0)
tk.Label(master, text="Row Number").grid(row=1)
tk.Label(master, text="Output Name").grid(row=2)

e1 = tk.Entry(master)
e2 = tk.Entry(master)
e3 = tk.Entry(master)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2, column=1)

tk.Button(master, text='Quit', command=master.quit).grid(row=3, column=0, sticky=tk.W, pady=4)
tk.Button(master, text='Run', command=show_entry_fields).grid(row=3, column=1, sticky=tk.W, pady=4)


tk.mainloop()
