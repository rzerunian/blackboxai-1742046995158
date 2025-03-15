import tkinter as tk
from ui_main import MainApplication

def main():
    """
    Entry point of the application.
    Initializes the main window and starts the event loop.
    """
    root = tk.Tk()
    app = MainApplication(root)
    root.mainloop()

if __name__ == "__main__":
    main()
