import os
import json
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD

try:
    import pythoncom
    import win32com.client
except ImportError:
    win32com = None

APP_GROUP_FILE = 'apps_groups.json'

def clean_path(p):
    return os.path.abspath(p.strip().strip('"').strip("'")).lower()

def resolve_lnk(lnk_path):
    if not win32com or not lnk_path.lower().endswith('.lnk'):
        return None
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        path = shortcut.Targetpath
        if path and path.lower().endswith('.exe') and os.path.isfile(path):
            return path
    except Exception as e:
        print(f"Failed to resolve shortcut: {lnk_path}, {e}")
    return None

class GroupLauncherApp:
    def __init__(self, master):
        self.master = master
        master.title("Batch Group Launcher - Steven Edition")

        self.groups = self.load_groups()
        self.current_group = None

        # Group selection & management
        top_frame = tk.Frame(master)
        top_frame.pack(padx=10, pady=6)

        tk.Label(top_frame, text="Batch Group:").pack(side=tk.LEFT)
        self.group_var = tk.StringVar()
        self.group_menu = tk.OptionMenu(top_frame, self.group_var, *self.groups.keys(), command=self.switch_group)
        self.group_menu.pack(side=tk.LEFT, padx=5)

        self.add_group_btn = tk.Button(top_frame, text="Add Group", command=self.add_group)
        self.add_group_btn.pack(side=tk.LEFT, padx=3)
        self.del_group_btn = tk.Button(top_frame, text="Delete Group", command=self.delete_group)
        self.del_group_btn.pack(side=tk.LEFT, padx=3)

        # Software list area (supports drag & drop)
        self.listbox = tk.Listbox(master, width=60, height=10)
        self.listbox.pack(padx=10, pady=8)
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.drop_software)

        # Software management buttons
        btn_frame = tk.Frame(master)
        btn_frame.pack(pady=5)

        self.add_btn = tk.Button(btn_frame, text="Add Program", command=self.add_software)
        self.add_btn.pack(side=tk.LEFT, padx=5)
        self.del_btn = tk.Button(btn_frame, text="Remove Selected", command=self.remove_selected)
        self.del_btn.pack(side=tk.LEFT, padx=5)

        self.exit_btn = tk.Button(btn_frame, text="Exit", command=self.master.quit, fg="green")
        self.exit_btn.pack(side=tk.LEFT, padx=10)

        # Launch button
        self.launch_btn = tk.Button(master, text="Launch All in Current Group", command=self.launch_softwares, bg='#4caf50', fg='white')
        self.launch_btn.pack(pady=8, ipadx=10, ipady=3)

        # Right-click paste path
        self.listbox.bind("<Button-3>", self.paste_from_clipboard)

        # Select the first group by default
        if self.groups:
            default_group = list(self.groups.keys())[0]
            self.group_var.set(default_group)
            self.switch_group(default_group)
        else:
            self.add_group()

    def refresh_group_menu(self):
        menu = self.group_menu["menu"]
        menu.delete(0, "end")
        for group in self.groups.keys():
            menu.add_command(label=group, command=lambda g=group: self.group_var.set(g))
        if self.groups:
            self.group_var.set(list(self.groups.keys())[0])
            self.switch_group(self.group_var.get())

    def load_groups(self):
        if os.path.exists(APP_GROUP_FILE):
            try:
                with open(APP_GROUP_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for group, paths in data.items():
                        data[group] = [clean_path(p) for p in paths if os.path.isfile(clean_path(p))]
                    return data
            except Exception as e:
                print(f"Failed to load groups: {e}")
        return {"Default Group": []}

    def save_groups(self):
        with open(APP_GROUP_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.groups, f, ensure_ascii=False, indent=2)

    def add_group(self):
        group_name = simpledialog.askstring("New Group", "Enter group name:")
        if group_name and group_name not in self.groups:
            self.groups[group_name] = []
            self.refresh_group_menu()
            self.group_var.set(group_name)
            self.switch_group(group_name)
            self.save_groups()
        elif group_name:
            messagebox.showinfo("Info", "Group name already exists!")

    def delete_group(self):
        if not self.groups:
            return
        group_name = self.group_var.get()
        if messagebox.askyesno("Confirm", f"Are you sure to delete group '{group_name}'?"):
            self.groups.pop(group_name)
            self.save_groups()
            if self.groups:
                new_group = list(self.groups.keys())[0]
                self.group_var.set(new_group)
                self.switch_group(new_group)
            else:
                self.listbox.delete(0, tk.END)

    def switch_group(self, group_name):
        self.current_group = group_name
        self.listbox.delete(0, tk.END)
        for path in self.groups.get(group_name, []):
            self.listbox.insert(tk.END, path)

    def add_software(self):
        path = filedialog.askopenfilename(title="Select Program", filetypes=[("Executable or Shortcut", "*.exe *.lnk")])
        if path:
            real_path = path
            if path.lower().endswith('.lnk'):
                resolved = resolve_lnk(path)
                if resolved:
                    real_path = resolved
                else:
                    messagebox.showwarning("Invalid Shortcut", "Failed to resolve this shortcut to an EXE, ignored.")
                    return
            path_clean = clean_path(real_path)
            current_list = [clean_path(x) for x in self.groups[self.current_group]]
            if path_clean and path_clean not in current_list:
                self.groups[self.current_group].append(path_clean)
                self.listbox.insert(tk.END, path_clean)
                self.save_groups()

    def remove_selected(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showinfo("Info", "Please select the program(s) to remove first.")
            return
        for idx in reversed(selected):
            del self.groups[self.current_group][idx]
            self.listbox.delete(idx)
        self.save_groups()

    def launch_softwares(self):
        paths = self.groups.get(self.current_group, [])
        if not paths:
            messagebox.showwarning("Warning", "No program in the current group!")
            return
        errors = []
        for path in paths:
            try:
                subprocess.Popen(path)
            except Exception as e:
                errors.append(f"{path}\n    Launch failed: {e}")
        if errors:
            messagebox.showerror("Failed to launch these programs", "\n\n".join(errors))
        else:
            messagebox.showinfo("OK", f"All programs in group '{self.current_group}' have been launched!")

    def paste_from_clipboard(self, event):
        try:
            path = self.master.clipboard_get()
            if path.lower().endswith('.lnk'):
                resolved = resolve_lnk(path)
                if not resolved:
                    messagebox.showwarning("Invalid Shortcut", "Failed to resolve this shortcut to an EXE, ignored.")
                    return
                path = resolved
            path_clean = clean_path(path)
            current_list = [clean_path(x) for x in self.groups[self.current_group]]
            if os.path.isfile(path_clean) and path_clean.endswith(".exe") and path_clean not in current_list:
                self.groups[self.current_group].append(path_clean)
                self.listbox.insert(tk.END, path_clean)
                self.save_groups()
        except:
            pass

    def drop_software(self, event):
        raw = event.data
        try:
            raw_list = self.master.tk.splitlist(raw)
        except Exception:
            raw_list = raw.split()
        paths = []
        for item in raw_list:
            if item.startswith("{") and item.endswith("}"):
                inner = item[1:-1]
                paths.extend(inner.split())
            else:
                paths.append(item)
        current_list = [clean_path(x) for x in self.groups[self.current_group]]
        added = 0
        for path in paths:
            target = path
            if path.lower().endswith('.lnk'):
                resolved = resolve_lnk(path)
                if not resolved:
                    continue
                target = resolved
            path_clean = clean_path(target)
            if os.path.isfile(path_clean) and path_clean.endswith('.exe') and path_clean not in current_list:
                self.groups[self.current_group].append(path_clean)
                self.listbox.insert(tk.END, path_clean)
                current_list.append(path_clean)
                added += 1
        if added > 0:
            self.save_groups()
        else:
            messagebox.showinfo("Info", "All dropped EXE files are already in the current group, no duplicates added.")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = GroupLauncherApp(root)
    root.resizable(False, False)
    root.mainloop()
