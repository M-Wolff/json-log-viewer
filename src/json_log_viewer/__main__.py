from __future__ import annotations

import argparse
import tkinter as tk

from .gui import JsonLogViewerApp


def main() -> None:
    parser = argparse.ArgumentParser(description="Open the JSON log viewer GUI.")
    parser.add_argument("path", nargs="?", help="Optional path to a JSON file")
    args = parser.parse_args()

    root = tk.Tk()
    JsonLogViewerApp(root, initial_path=args.path)
    root.mainloop()


if __name__ == "__main__":
    main()
