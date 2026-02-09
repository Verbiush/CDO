import tkinter as tk
from organizador import OrganizadorArchivosApp

def main():
    root = tk.Tk()
    app = OrganizadorArchivosApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
