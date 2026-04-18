import tkinter as tk
from tkinter import ttk
from Auto import AutoTab
from Dws import DwsTab
from Filter import FilterTab


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("AD:P")
        self.geometry("900x720")
        self.resizable(False, False)

        self.build_header()
        self.style_tabs()
        self.build_tabs()

    def build_header(self):
        header = tk.Frame(self, bg="#34495e", height=48)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Auto Packing & DWS Processor",
            fg="white",
            bg="#34495e",
            font=("Segoe UI", 14, "bold")
        ).pack(side="left", padx=20)

    def style_tabs(self):
        style = ttk.Style()
        style.theme_use("default")

        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding=[18, 10]
        )

        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", "#ffffff"),
                ("active", "#ecf0f1")
            ]
        )

    def build_tabs(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        notebook.add(AutoTab(notebook), text="📦 AUTO")
        notebook.add(DwsTab(notebook), text="📊 DWS")
        notebook.add(FilterTab(notebook), text="🔎 Filter ค่าปรับ")

        notebook.select(0)

if __name__ == "__main__":
    MainApp().mainloop()
