from lib.tk_ui import GUI
import tkinter as tk

def run():
    root = tk.Tk()
    ui = GUI(root)
    root.mainloop()

if __name__ == '__main__':
    run()
