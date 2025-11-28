#!/usr/bin/env python3
# WPS Office Theme Selector (GUI з підписами під зображеннями)
# Використовує лише tkinter (без Pillow)
# Зображення: 256×189, формат PNG

import sys
from pathlib import Path
from configparser import ConfigParser
import tkinter as tk
from tkinter import messagebox

# --- Шляхи ---
SCRIPT_DIR = Path(__file__).parent.resolve()
SKIN_IMG_DIR = SCRIPT_DIR / "skin-img"
USER_SKINS_DIR = Path.home() / ".skins"
HISTORY_INI = Path.home() / ".local/share/Kingsoft/office6/skinsv2/default/histroy.ini"
HISTORY_INI.parent.mkdir(parents=True, exist_ok=True)

def get_skin_names():
    """Повертає список тем (назви .png без розширення)."""
    if not SKIN_IMG_DIR.is_dir():
        messagebox.showerror("Помилка", f"Не знайдено теку зі зразками:\n{SKIN_IMG_DIR}")
        sys.exit(1)
    return sorted(
        f.stem for f in SKIN_IMG_DIR.iterdir()
        if f.is_file() and f.suffix.lower() == ".png"
    )

def ensure_user_skins(skin_names):
    """Створює ~/.skins/<skin>, якщо не існує."""
    USER_SKINS_DIR.mkdir(exist_ok=True)
    for name in skin_names:
        skin_dir = USER_SKINS_DIR / name
        if not skin_dir.is_dir():
            skin_dir.mkdir()

def write_history_ini(selected_skin):
    """Записує history.ini з абсолютними шляхами."""
    config = ConfigParser()
    config.optionxform = str

    if HISTORY_INI.exists():
        config.read(HISTORY_INI, encoding='utf-8')

    if 'skinPathPool' not in config:
        config['skinPathPool'] = {}
    if 'wpsoffice' not in config:
        config['wpsoffice'] = {}

    pool = {}
    for skin in get_skin_names():
        pool[skin] = str((USER_SKINS_DIR / skin).resolve())
    config['skinPathPool'] = pool
    config['wpsoffice']['lastSkin'] = selected_skin

    with open(HISTORY_INI, 'w', encoding='utf-8') as f:
        config.write(f)

def on_skin_click(skin_name, root):
    if messagebox.askyesno("Підтвердження", f"Встановити тему:\n{skin_name}?"):
        ensure_user_skins([skin_name])
        write_history_ini(skin_name)
        messagebox.showinfo("Готово", f"Тема '{skin_name}' встановлена!\nПерезапустіть WPS Office.")
        root.quit()

def main():
    skin_names = get_skin_names()
    if not skin_names:
        messagebox.showerror("Помилка", "У 'skin-img/' немає файлів .png")
        return

    root = tk.Tk()
    root.title("Вибір теми WPS Office")
    root.geometry("1200x700+50+50")
    root.minsize(1000, 600)

    # Прокручуване полотно
    canvas = tk.Canvas(root, highlightthickness=0)
    scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    cols = 4
    for i, skin in enumerate(skin_names):
        row = i // cols
        col = i % cols
        if col == 0:
            row_frame = tk.Frame(scrollable_frame)
            row_frame.pack(pady=15)

        # Фрейм для зображення + підпису
        item_frame = tk.Frame(row_frame)
        item_frame.pack(side=tk.LEFT, padx=15)

        img_path = SKIN_IMG_DIR / f"{skin}.png"
        try:
            photo = tk.PhotoImage(file=img_path)
        except tk.TclError as e:
            print(f"Помилка завантаження {img_path}: {e}")
            # fallback: лише назва
            label = tk.Label(item_frame, text=skin, font=("Arial", 10), width=32, height=12, relief="raised")
            label.pack()
            label.bind("<Button-1>", lambda e, s=skin: on_skin_click(s, root))
            continue

        # Кнопка зі зображенням
        btn = tk.Button(item_frame, image=photo, command=lambda s=skin: on_skin_click(s, root), bd=2)
        btn.image = photo
        btn.pack()

        # Підпис **під** зображенням
        caption = tk.Label(item_frame, text=skin, font=("Arial", 9), pady=4)
        caption.pack()

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    root.mainloop()

if __name__ == '__main__':
    main()
