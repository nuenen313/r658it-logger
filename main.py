"""
R6581T Data Logger — entry point.

Run with:
    poetry run python main.py
"""

from gui.main_window import MainWindow


def main() -> None:
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
