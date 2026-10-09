"""
LaunchPad - one click opens your whole workspace.
Profiles ("setups") hold apps, folders and websites. Pick a setup, hit Launch.
"""
import json
import os
import sys
import subprocess
import threading
import time
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

APP_DIR = Path(__file__).resolve().parent
CONFIG = APP_DIR / "profiles.json"
PYTHONW = Path(sys.executable).with_name("pythonw.exe")
if not PYTHONW.exists():
    PYTHONW = Path(sys.executable)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT = "#6C5CE7"
ACCENT_HOVER = "#5A4BD1"
BG = "#0F1117"
CARD = "#181B25"
CARD_HOVER = "#1F2330"
MUTED = "#8B90A5"
TEXT = "#ECEFF4"
DANGER = "#E04F5F"
OK = "#2ECC71"

ICONS = ["💼", "🎓", "🎮", "🎨", "🎬", "🧪", "🚀", "💻", "🎵", "📚", "🏠", "⚡"]


# ----------------------------------------------------------------- storage --
def load_profiles():
    if CONFIG.exists():
        try:
            return json.loads(CONFIG.read_text(encoding="utf-8"))
        except Exception:
            pass
    return [
        {
            "name": "Work",
            "icon": "💼",
            "delay": 1.5,
            "items": [
                {"type": "url", "path": "https://github.com", "name": "GitHub"},
                {"type": "url", "path": "https://mail.google.com", "name": "Gmail"},
            ],
        }
    ]


def save_profiles(profiles):
    CONFIG.write_text(json.dumps(profiles, indent=2, ensure_ascii=False), encoding="utf-8")


def item_label(item):
    if item.get("name"):
        return item["name"]
    p = item["path"]
    if item["type"] == "url":
        return p.replace("https://", "").replace("http://", "").rstrip("/")
    return Path(p).stem or p


def item_icon(item):
    return {"app": "🖥️", "folder": "📁", "url": "🌐", "file": "📄"}.get(item["type"], "•")


# ----------------------------------------------------------------- launch ---
def open_item(item):
    t, p = item["type"], item["path"]
    if t == "url":
        webbrowser.open(p)
    elif t == "app" and p.lower().endswith(".exe"):
        subprocess.Popen([p], cwd=str(Path(p).parent), shell=False)
    else:  # folder, file, .lnk, .bat
        os.startfile(p)


def launch_profile(profile, on_progress=None):
    items = profile.get("items", [])
    delay = float(profile.get("delay", 1.0))
    for i, item in enumerate(items):
        try:
            open_item(item)
            ok, err = True, ""
        except Exception as e:
            ok, err = False, str(e)
        if on_progress:
            on_progress(i, item, ok, err)
        if i < len(items) - 1:
            time.sleep(delay)


# --------------------------------------------------------------- shortcuts --
def _write_shortcut(lnk, profile):
    from win32com.client import Dispatch
    shell = Dispatch("WScript.Shell")
    sc = shell.CreateShortCut(str(lnk))
    sc.Targetpath = str(PYTHONW)
    sc.Arguments = f'"{APP_DIR / "launchpad.pyw"}" --launch "{profile["name"]}"'
    sc.WorkingDirectory = str(APP_DIR)
    sc.IconLocation = str(PYTHONW)
    sc.save()


def create_desktop_shortcut(profile):
    import winshell
    lnk = Path(winshell.desktop()) / f"{profile['name']} - LaunchPad.lnk"
    _write_shortcut(lnk, profile)
    return lnk


def startup_path(profile):
    import winshell
    return Path(winshell.startup()) / f"LaunchPad - {profile['name']}.lnk"


def set_startup(profile, enabled):
    lnk = startup_path(profile)
    if enabled:
        _write_shortcut(lnk, profile)
    elif lnk.exists():
        lnk.unlink()


# ------------------------------------------------------------------- UI -----
class HoverCard(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, fg_color=CARD, corner_radius=18, **kw)
        self.bind("<Enter>", lambda e: self.configure(fg_color=CARD_HOVER))
        self.bind("<Leave>", lambda e: self.configure(fg_color=CARD))


class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=BG)
        self.title("LaunchPad")
        self.geometry("900x600")
        self.minsize(760, 520)
        self.profiles = load_profiles()
        self.current = None
        self.show_home()

    # ---------- helpers
    def clear(self):
        for w in self.winfo_children():
            w.destroy()

    def header(self, title, subtitle, back=None):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=36, pady=(28, 10))
        if back:
            ctk.CTkButton(bar, text="←", width=40, height=40, corner_radius=12,
                          fg_color=CARD, hover_color=CARD_HOVER, text_color=TEXT,
                          font=ctk.CTkFont(size=18), command=back).pack(side="left", padx=(0, 14))
        box = ctk.CTkFrame(bar, fg_color="transparent")
        box.pack(side="left")
        ctk.CTkLabel(box, text=title, text_color=TEXT, anchor="w",
                     font=ctk.CTkFont("Segoe UI", 28, "bold")).pack(anchor="w")
        ctk.CTkLabel(box, text=subtitle, text_color=MUTED, anchor="w",
                     font=ctk.CTkFont("Segoe UI", 13)).pack(anchor="w")
        return bar

    # ---------- HOME: choose a setup
    def show_home(self):
        self.clear()
        bar = self.header("Choose your setup", "One click opens everything in a setup.")
        ctk.CTkButton(bar, text="+  New setup", height=40, corner_radius=12,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      font=ctk.CTkFont("Segoe UI", 13, "bold"),
                      command=self.new_profile).pack(side="right")

        grid = ctk.CTkScrollableFrame(self, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=28, pady=(6, 24))
        cols = 3
        for c in range(cols):
            grid.grid_columnconfigure(c, weight=1, uniform="col")

        for i, prof in enumerate(self.profiles):
            self.profile_card(grid, prof, i // cols, i % cols)

        if not self.profiles:
            ctk.CTkLabel(grid, text="No setups yet. Click  + New setup  to create one.",
                         text_color=MUTED, font=ctk.CTkFont("Segoe UI", 14)).grid(
                row=0, column=0, columnspan=cols, pady=80)

    def profile_card(self, parent, prof, r, c):
        card = HoverCard(parent)
        card.grid(row=r, column=c, padx=8, pady=8, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(card, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 4))
        ctk.CTkLabel(top, text=prof.get("icon", "🚀"), font=ctk.CTkFont(size=34)).pack(side="left")
        ctk.CTkButton(top, text="⋯", width=32, height=32, corner_radius=10, fg_color="transparent",
                      hover_color=BG, text_color=MUTED, font=ctk.CTkFont(size=18),
                      command=lambda p=prof: self.show_editor(p)).pack(side="right")

        ctk.CTkLabel(card, text=prof["name"], text_color=TEXT, anchor="w",
                     font=ctk.CTkFont("Segoe UI", 20, "bold")).grid(
            row=1, column=0, sticky="w", padx=20)
        n = len(prof.get("items", []))
        plural = "s" if n != 1 else ""
        ctk.CTkLabel(card, text=f"{n} item{plural}", text_color=MUTED, anchor="w",
                     font=ctk.CTkFont("Segoe UI", 12)).grid(row=2, column=0, sticky="w", padx=20)

        preview = "   ".join(item_icon(it) for it in prof.get("items", [])[:8])
        ctk.CTkLabel(card, text=preview or " ", text_color=MUTED, anchor="w",
                     font=ctk.CTkFont(size=16)).grid(row=3, column=0, sticky="w", padx=20, pady=(10, 0))

        ctk.CTkButton(card, text="Launch  ▶", height=42, corner_radius=12,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      font=ctk.CTkFont("Segoe UI", 14, "bold"),
                      command=lambda p=prof: self.run_profile(p)).grid(
            row=4, column=0, sticky="ew", padx=20, pady=(18, 20))

    # ---------- RUN: progress screen
    def run_profile(self, prof):
        self.clear()
        self.header(f"Launching {prof['name']}", "Sit back, opening everything for you.")
        box = ctk.CTkFrame(self, fg_color=CARD, corner_radius=18)
        box.pack(fill="both", expand=True, padx=36, pady=(6, 28))

        bar = ctk.CTkProgressBar(box, progress_color=ACCENT, fg_color=BG, height=10, corner_radius=6)
        bar.pack(fill="x", padx=24, pady=(26, 10))
        bar.set(0)

        rows = {}
        lst = ctk.CTkScrollableFrame(box, fg_color="transparent")
        lst.pack(fill="both", expand=True, padx=12, pady=6)
        for i, it in enumerate(prof["items"]):
            r = ctk.CTkFrame(lst, fg_color="transparent")
            r.pack(fill="x", padx=10, pady=4)
            ctk.CTkLabel(r, text=item_icon(it), font=ctk.CTkFont(size=16), width=30).pack(side="left")
            ctk.CTkLabel(r, text=item_label(it), text_color=TEXT, anchor="w",
                         font=ctk.CTkFont("Segoe UI", 14)).pack(side="left", padx=6)
            st = ctk.CTkLabel(r, text="waiting", text_color=MUTED, font=ctk.CTkFont("Segoe UI", 12))
            st.pack(side="right", padx=8)
            rows[i] = st

        done_btn = ctk.CTkButton(box, text="Done  →  Back to setups", height=42, corner_radius=12,
                                 fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                 font=ctk.CTkFont("Segoe UI", 14, "bold"),
                                 command=self.show_home, state="disabled")
        done_btn.pack(fill="x", padx=24, pady=(10, 22))

        total = max(len(prof["items"]), 1)

        def progress(i, item, ok, err):
            def ui():
                rows[i].configure(text="opened ✓" if ok else f"failed: {err[:40]}",
                                  text_color=OK if ok else DANGER)
                bar.set((i + 1) / total)
                if i + 1 == total:
                    done_btn.configure(state="normal")
            self.after(0, ui)

        def worker():
            launch_profile(prof, progress)
            if not prof["items"]:
                self.after(0, lambda: done_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    # ---------- EDITOR
    def new_profile(self):
        dlg = ctk.CTkInputDialog(text="Name for the new setup:", title="New setup")
        name = (dlg.get_input() or "").strip()
        if not name:
            return
        prof = {"name": name, "icon": ICONS[len(self.profiles) % len(ICONS)], "delay": 1.5, "items": []}
        self.profiles.append(prof)
        save_profiles(self.profiles)
        self.show_editor(prof)

    def show_editor(self, prof):
        self.clear()
        self.current = prof
        bar = self.header(f"{prof.get('icon', '')}  {prof['name']}", "Add apps, folders and websites to this setup.",
                          back=self.show_home)
        ctk.CTkButton(bar, text="Launch  ▶", height=40, corner_radius=12, fg_color=ACCENT,
                      hover_color=ACCENT_HOVER, font=ctk.CTkFont("Segoe UI", 13, "bold"),
                      command=lambda: self.run_profile(prof)).pack(side="right")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=36, pady=(6, 28))
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        # left: items
        left = ctk.CTkFrame(body, fg_color=CARD, corner_radius=18)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        addbar = ctk.CTkFrame(left, fg_color="transparent")
        addbar.pack(fill="x", padx=16, pady=(16, 6))
        for txt, cmd in (("+ App", self.add_app), ("+ Folder", self.add_folder),
                         ("+ Website", self.add_url), ("+ File", self.add_file)):
            ctk.CTkButton(addbar, text=txt, height=34, corner_radius=10, fg_color=BG,
                          hover_color=CARD_HOVER, text_color=TEXT,
                          font=ctk.CTkFont("Segoe UI", 12, "bold"), command=cmd).pack(
                side="left", padx=(0, 6))

        self.items_frame = ctk.CTkScrollableFrame(left, fg_color="transparent")
        self.items_frame.pack(fill="both", expand=True, padx=8, pady=(4, 12))
        self.render_items()

        # right: settings
        right = ctk.CTkFrame(body, fg_color=CARD, corner_radius=18)
        right.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(right, text="Settings", text_color=TEXT,
                     font=ctk.CTkFont("Segoe UI", 16, "bold")).pack(anchor="w", padx=20, pady=(18, 10))

        ctk.CTkLabel(right, text="Icon", text_color=MUTED).pack(anchor="w", padx=20)
        icon_row = ctk.CTkFrame(right, fg_color="transparent")
        icon_row.pack(anchor="w", padx=16, pady=(2, 10))
        for i, ic in enumerate(ICONS):
            ctk.CTkButton(icon_row, text=ic, width=36, height=36, corner_radius=10,
                          fg_color=ACCENT if ic == prof.get("icon") else BG, hover_color=CARD_HOVER,
                          font=ctk.CTkFont(size=16),
                          command=lambda ic=ic: self.set_icon(ic)).grid(row=i // 6, column=i % 6, padx=3, pady=3)

        ctk.CTkLabel(right, text="Delay between items (seconds)", text_color=MUTED).pack(anchor="w", padx=20, pady=(6, 0))
        self.delay_lbl = ctk.CTkLabel(right, text=f"{float(prof.get('delay', 1.5)):.1f}s", text_color=TEXT)
        self.delay_lbl.pack(anchor="w", padx=20)
        sl = ctk.CTkSlider(right, from_=0, to=5, number_of_steps=10, progress_color=ACCENT,
                           button_color=ACCENT, button_hover_color=ACCENT_HOVER, command=self.set_delay)
        sl.set(float(prof.get("delay", 1.5)))
        sl.pack(fill="x", padx=20, pady=(0, 14))

        self.startup_sw = ctk.CTkSwitch(right, text="Launch when Windows starts", text_color=TEXT,
                                        progress_color=ACCENT, command=self.toggle_startup)
        try:
            if startup_path(prof).exists():
                self.startup_sw.select()
        except Exception:
            pass
        self.startup_sw.pack(anchor="w", padx=20, pady=6)

        ctk.CTkButton(right, text="Create desktop shortcut", height=38, corner_radius=10, fg_color=BG,
                      hover_color=CARD_HOVER, text_color=TEXT, command=self.make_shortcut).pack(
            fill="x", padx=20, pady=(14, 6))
        ctk.CTkButton(right, text="Rename setup", height=38, corner_radius=10, fg_color=BG,
                      hover_color=CARD_HOVER, text_color=TEXT, command=self.rename).pack(fill="x", padx=20, pady=6)
        ctk.CTkButton(right, text="Delete setup", height=38, corner_radius=10, fg_color="transparent",
                      border_width=1, border_color=DANGER, hover_color="#3a1b20", text_color=DANGER,
                      command=self.delete).pack(fill="x", padx=20, pady=(6, 20))

    def render_items(self):
        for w in self.items_frame.winfo_children():
            w.destroy()
        items = self.current["items"]
        if not items:
            ctk.CTkLabel(self.items_frame, text="Nothing here yet. Use the buttons above to add things.",
                         text_color=MUTED, font=ctk.CTkFont("Segoe UI", 13)).pack(pady=60)
            return
        for i, it in enumerate(items):
            row = ctk.CTkFrame(self.items_frame, fg_color=BG, corner_radius=12)
            row.pack(fill="x", padx=8, pady=4)
            ctk.CTkLabel(row, text=item_icon(it), font=ctk.CTkFont(size=18), width=34).pack(side="left", padx=(10, 0), pady=10)
            txt = ctk.CTkFrame(row, fg_color="transparent")
            txt.pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkLabel(txt, text=item_label(it), text_color=TEXT, anchor="w",
                         font=ctk.CTkFont("Segoe UI", 13, "bold")).pack(anchor="w")
            ctk.CTkLabel(txt, text=it["path"], text_color=MUTED, anchor="w",
                         font=ctk.CTkFont("Segoe UI", 11)).pack(anchor="w")
            ctk.CTkButton(row, text="✕", width=30, height=30, corner_radius=8, fg_color="transparent",
                          hover_color="#3a1b20", text_color=DANGER,
                          command=lambda i=i: self.remove_item(i)).pack(side="right", padx=(0, 8))
            ctk.CTkButton(row, text="↓", width=30, height=30, corner_radius=8, fg_color="transparent",
                          hover_color=CARD_HOVER, text_color=MUTED,
                          command=lambda i=i: self.move_item(i, 1)).pack(side="right")
            ctk.CTkButton(row, text="↑", width=30, height=30, corner_radius=8, fg_color="transparent",
                          hover_color=CARD_HOVER, text_color=MUTED,
                          command=lambda i=i: self.move_item(i, -1)).pack(side="right")

    # ---------- editor actions
    def _add(self, item):
        self.current["items"].append(item)
        save_profiles(self.profiles)
        self.render_items()

    def add_app(self):
        p = filedialog.askopenfilename(title="Pick an app", initialdir=os.environ.get("ProgramFiles", "C:\\"),
                                       filetypes=[("Programs", "*.exe *.lnk *.bat *.cmd"), ("All", "*.*")])
        if p:
            self._add({"type": "app", "path": p, "name": Path(p).stem})

    def add_file(self):
        p = filedialog.askopenfilename(title="Pick a file")
        if p:
            self._add({"type": "file", "path": p, "name": Path(p).name})

    def add_folder(self):
        p = filedialog.askdirectory(title="Pick a folder")
        if p:
            self._add({"type": "folder", "path": p, "name": Path(p).name or p})

    def add_url(self):
        dlg = ctk.CTkInputDialog(text="Website address:", title="Add website")
        u = (dlg.get_input() or "").strip()
        if not u:
            return
        if not u.startswith(("http://", "https://")):
            u = "https://" + u
        self._add({"type": "url", "path": u, "name": ""})

    def remove_item(self, i):
        del self.current["items"][i]
        save_profiles(self.profiles)
        self.render_items()

    def move_item(self, i, d):
        items = self.current["items"]
        j = i + d
        if 0 <= j < len(items):
            items[i], items[j] = items[j], items[i]
            save_profiles(self.profiles)
            self.render_items()

    def set_icon(self, ic):
        self.current["icon"] = ic
        save_profiles(self.profiles)
        self.show_editor(self.current)

    def set_delay(self, v):
        self.current["delay"] = round(float(v), 1)
        self.delay_lbl.configure(text=f"{self.current['delay']:.1f}s")
        save_profiles(self.profiles)

    def toggle_startup(self):
        try:
            set_startup(self.current, bool(self.startup_sw.get()))
        except Exception as e:
            messagebox.showerror("LaunchPad", f"Could not change startup setting:\n{e}")

    def make_shortcut(self):
        try:
            lnk = create_desktop_shortcut(self.current)
            messagebox.showinfo("LaunchPad", f"Shortcut created on your desktop:\n{lnk.name}")
        except Exception as e:
            messagebox.showerror("LaunchPad", f"Could not create shortcut:\n{e}")

    def rename(self):
        dlg = ctk.CTkInputDialog(text="New name:", title="Rename setup")
        n = (dlg.get_input() or "").strip()
        if n:
            self.current["name"] = n
            save_profiles(self.profiles)
            self.show_editor(self.current)

    def delete(self):
        if messagebox.askyesno("Delete setup", f"Delete '{self.current['name']}'? This cannot be undone."):
            try:
                set_startup(self.current, False)
            except Exception:
                pass
            self.profiles.remove(self.current)
            save_profiles(self.profiles)
            self.show_home()


# ------------------------------------------------------------------ main ----
def main():
    if "--launch" in sys.argv:
        name = sys.argv[sys.argv.index("--launch") + 1]
        for p in load_profiles():
            if p["name"].lower() == name.lower():
                launch_profile(p)
                return
        return
    App().mainloop()


if __name__ == "__main__":
    main()
