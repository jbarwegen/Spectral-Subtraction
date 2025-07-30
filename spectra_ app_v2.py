# file: spectra_subtraction_app.py
from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import os


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, _=None):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.geometry(f"+{x}+{y}")
        label = ttk.Label(
            tw,
            text=self.text,
            background="yellow",
            relief="solid",
            borderwidth=1,
            padding=2,
        )
        label.pack()

    def hide_tip(self, _=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


class SpectraApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Spectra Subtraction App")
        self.geometry("450x580")

        self.dual_listbox = None
        self.c_var = None
        self.a_var = None
        self.sheet_names = []
        self.data_dict = {}
        self.save_path: str | None = None

        self._create_widgets()
        self.status_label = ttk.Label(self, text="Ready", relief="sunken", anchor="w")
        self.status_label.grid(
            row=11, column=0, columnspan=2, sticky="we", padx=5, pady=5
        )
        self._enable_drag_drop()

    # ---- Drag and Drop ----
    def _enable_drag_drop(self):
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self._handle_drop)

    def _handle_drop(self, event):
        self.set_status("Loading file...")
        path = event.data.strip("{}")
        if path.lower().endswith((".xls", ".xlsx")):
            try:
                skip = int(self.skip_spin.get())
                self.sheet_names, self.data_dict = self.load_data(skip, path)
                self._populate_widgets()
                messagebox.showinfo(
                    "File Loaded",
                    f"Loaded {len(self.sheet_names)} sheets from:\n{path}",
                )
                self.set_status("Loaded successfully")
            except Exception as e:
                messagebox.showerror("Load Error", f"Could not load file:\n{e}")
        else:
            messagebox.showwarning(
                "Invalid File", "Please drop a valid Excel file (.xls, .xlsx)"
            )

    # ---- Widgets ----
    def _create_widgets(self):
        ttk.Label(self, text="Rows to skip:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.skip_spin = ttk.Spinbox(self, from_=0, to=100, width=5)
        self.skip_spin.set(6)
        self.skip_spin.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.add_tooltip(
            self.skip_spin,
            "Number of rows to skip at the top of each sheet (rows before headers)",
        )

        load_btn = ttk.Button(self, text="Load File…", command=self.load_file)
        load_btn.grid(row=1, column=0, columnspan=2, pady=5)
        self.add_tooltip(
            load_btn, "Load an Excel file with spectra on separate sheets"
        )

        self._make_dropdown("Spectrum A", 2, "a")
        self._make_dropdown("Control C:", 3, "c")

        ttk.Label(self, text="Spectra to subtract (B1, B2, ...):").grid(
            row=4, column=0, sticky="ne", padx=5
        )
        self.b_listbox = tk.Listbox(
            self, selectmode=tk.MULTIPLE, height=5, exportselection=False
        )
        self.b_listbox.grid(row=4, column=1, sticky="w", padx=5)
        self.add_tooltip(
            self.b_listbox,
            "Select one or more spectra to subtract from spectrum A",
        )

        ttk.Label(self, text="Annotate top N peaks:").grid(
            row=6, column=0, sticky="w", padx=5, pady=5
        )
        self.peak_spin = ttk.Spinbox(self, from_=1, to=50, width=5)
        self.peak_spin.set(10)
        self.peak_spin.grid(row=6, column=1, sticky="w", padx=5, pady=5)
        self.add_tooltip(self.peak_spin, "Number of peaks to annotate in each plot")

        self.normalize_var = tk.BooleanVar(self)
        norm_cb = ttk.Checkbutton(
            self, text="Normalize", variable=self.normalize_var
        )
        norm_cb.grid(row=7, column=0, columnspan=2, pady=5)
        self.add_tooltip(norm_cb, "Scale Relative values to 0-100 range")

        self.save_var = tk.BooleanVar(self)
        save_cb = ttk.Checkbutton(
            self, text="Save Figures", variable=self.save_var
        )
        save_cb.grid(row=8, column=0, columnspan=2, pady=5)
        self.add_tooltip(save_cb, "Save plots as SVG files instead of displaying them")

        self._make_dropdown("Spectra for Dual Plot:", 9, "b")
        self.dual_plot_var = tk.BooleanVar(self)
        dual_cb = ttk.Checkbutton(
            self, text="Plot Dual Spectrum", variable=self.dual_plot_var
        )
        dual_cb.grid(row=10, column=0, columnspan=2, pady=5)
        self.add_tooltip(
            dual_cb,
            "Show a dual plot of positive and negative peaks after subtraction",
        )

        run_btn = ttk.Button(self, text="Run", command=self.run)
        run_btn.grid(row=12, column=0, columnspan=2, pady=10)
        self.add_tooltip(run_btn, "Run subtraction and display/save plots")

        save_location_btn = tk.Button(
            self, text="Choose save location", command=self.choose_save_location
        )
        save_location_btn.grid(row=13, column=0, columnspan=2, pady=5)

    def _make_dropdown(self, label, row, var_prefix):
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="e", padx=5)
        var = tk.StringVar(self)
        menu = ttk.OptionMenu(self, var, "")
        menu.grid(row=row, column=1, sticky="w", padx=5)
        setattr(self, f"{var_prefix}_var", var)
        setattr(self, f"{var_prefix}_menu", menu)
        self.add_tooltip(menu, f"Select the {label.lower()}")

    def add_tooltip(self, widget, text):
        ToolTip(widget, text)

    # ---- Status ----
    def set_status(self, text):
        self.status_label.config(text=text)
        self.update_idletasks()

    # ---- File Handling ----
    def choose_save_location(self):
        path = filedialog.askdirectory(
            title="Choose file save location",
        )
        if path:
            self.save_path = path
            messagebox.showinfo(
                "Save Location", f"Graphs will be saved to:\n{self.save_path}"
            )

    def save_data(self, data: str):
        if not self.save_path:
            messagebox.showwarning("No Path", "Please choose a save location first!")
            return
        with open(self.save_path, "w", encoding="utf-8") as file:
            file.write(data)
        messagebox.showinfo("Saved", f"Data saved to {self.save_path}")

    def load_file(self):
        path = filedialog.askopenfilename(
            title="Select the spectra file",
            filetypes=[("Excel files", "*.xlsx;*.xls")],
        )
        if not path:
            return

        skip = int(self.skip_spin.get())
        try:
            self.sheet_names, self.data_dict = self.load_data(skip, path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            return

        # Update dropdowns
        for var_name in ["a_var", "b_var", "c_var"]:
            var = getattr(self, var_name)
            menu = getattr(self, var_name.replace("_var", "_menu"))["menu"]
            menu.delete(0, "end")
            for name in self.sheet_names:
                menu.add_command(
                    label=name, command=lambda n=name, v=var: v.set(n)
                )
        self.a_var.set(self.sheet_names[0])
        self.b_var.set(self.sheet_names[0])
        self.c_var.set(self.sheet_names[0])

        self.b_listbox.delete(0, tk.END)
        for name in self.sheet_names:
            self.b_listbox.insert(tk.END, name)

    @staticmethod
    def load_data(skip_rows, path):
        xls = pd.ExcelFile(path)
        names = xls.sheet_names
        raw = pd.read_excel(xls, sheet_name=names, skiprows=skip_rows)
        filtered = {}
        for name, df in raw.items():
            required = {"m/z", "Relative", "Resolution", "Noise"}
            if not required.issubset(df.columns):
                raise ValueError(
                    f"Sheet '{name}' is missing columns: {required - set(df.columns)}"
                )
            filtered[name] = df.loc[df["Intensity"] > 10 * df["Noise"]]
        return names, filtered

    # ---- Main Processing ----
    def run(self):
        if not self.data_dict:
            messagebox.showwarning("No data", "Please load a file first.")
            return
        self.set_status("Running analysis...")
        name_a = self.a_var.get()
        name_b = self.b_var.get()
        name_c = self.c_var.get()
        selected_indices = self.b_listbox.curselection()
        N = int(self.peak_spin.get())
        if not selected_indices:
            messagebox.showwarning(
                "No spectra selected", "Please select spectra to subtract from A."
            )
            return

        self.plot_spectrum(self.data_dict[name_a], N, [name_a])
        for idx in selected_indices:
            name1 = self.sheet_names[idx]
            self.plot_spectrum(self.data_dict[name1], N, [name1])
        self.plot_spectrum(self.data_dict[name_c], N, [name_c])

        title = name_a + " Subtracted " + " Subtracted ".join(
            self.sheet_names[i] for i in selected_indices
        )
        df_down = self.data_dict[name_a].copy()
        for idx in selected_indices:
            name_d = self.sheet_names[idx]
            df_down = self.compare_dfs(df_down, self.data_dict[name_d])

        df_up = self.compare_dfs(self.data_dict[name_b], self.data_dict[name_a].copy())

        if self.normalize_var.get():
            df_down = self.normalize_df(df_down)
            df_up = self.normalize_df(df_up)

        if self.dual_plot_var.get():
            self.plot_dual_spectrum(df_down, df_up, N, title + " (Dual)")
        else:
            self.plot_spectrum(df_down, N, [title])
        self.set_status("Done")

    # ---- Helpers ----
    def compare_dfs(self, df1, df2):
        dfB = df2.copy()
        dfB["half_width_B"] = dfB["m/z"] / dfB["Resolution"] / 2

        def peak_match(row):
            m1 = row["m/z"]
            R1 = row["Resolution"]
            hw1 = m1 / R1 / 2
            max_hwB = dfB["half_width_B"].max()
            lo, hi = m1 - (hw1 + max_hwB), m1 + (hw1 + max_hwB)
            cand = dfB[(dfB["m/z"] >= lo) & (dfB["m/z"] <= hi)]
            if cand.empty:
                return False
            sep = (cand["m/z"] - m1).abs()
            overlap = sep <= (hw1 + cand["half_width_B"])
            delta_ppm = sep / ((cand["m/z"] + m1) / 2) * 1e6
            return (overlap & (delta_ppm <= 3)).any()

        dfA = df1.dropna(subset=["m/z"]).copy()
        mask = dfA.apply(peak_match, axis=1)
        return dfA.loc[~mask].reset_index(drop=True)

    def normalize_df(self, df):
        new = df.copy()
        new["Relative"] = new["Relative"] / new["Relative"].max() * 100
        return new

    def plot_spectrum(self, df, N, names):
        top = df.nlargest(N, "Relative")
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.vlines(df["m/z"], 0, df["Relative"], color="black")
        ax.set_title(names[0])
        ax.set_xlabel("m/z")
        ax.set_ylabel("Relative")
        ax.set_xlim(right=350)
        ax.set_ylim(bottom=0, top=115)
        for _, r in top.iterrows():
            ax.annotate(
                f"{r['m/z']:.4f}",
                xy=(r["m/z"], r["Relative"]),
                xytext=(0, 5),
                textcoords="offset points",
                ha="center",
                va="bottom",
                rotation=45,
                fontsize=8,
            )
        plt.tight_layout()
        if self.save_var.get():
            filename = names[0].replace(" ", "_") + ".svg"
            if self.save_path:
                filepath = os.path.join(self.save_path,filename)
            else: 
                filepath = filename
            fig.savefig(filepath)
            plt.close(fig)
        else:
            plt.show()
            plt.close(fig)

    def plot_dual_spectrum(self, df_up, df_down, N, title):
        top_up = df_up.nlargest(N, "Relative")
        top_down = df_down.nlargest(N, "Relative")
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.vlines(df_up["m/z"], 0, df_up["Relative"], color="black")
        ax.vlines(df_down["m/z"], 0, -df_down["Relative"], color="red")
        ax.set_title(title)
        ax.set_xlabel("m/z")
        ax.set_ylabel("Relative")
        ax.set_xlim(right=350)
        ax.set_ylim(-130, 130)
        ax.axhline(0, color="gray", linewidth=1)
        for _, r in top_up.iterrows():
            ax.annotate(
                f"{r['m/z']:.4f}",
                xy=(r["m/z"], r["Relative"]),
                xytext=(0, 5),
                textcoords="offset points",
                ha="center",
                va="bottom",
                rotation=45,
                fontsize=8,
            )
        for _, r in top_down.iterrows():
            ax.annotate(
                f"{r['m/z']:.4f}",
                xy=(r["m/z"], -r["Relative"]),
                xytext=(0, -5),
                textcoords="offset points",
                ha="center",
                va="top",
                rotation=45,
                fontsize=8,
                color="red",
            )
        plt.tight_layout()
        if self.save_var.get():
            fig.savefig(title.replace(" ", "_") + "_dual.svg")
            plt.close(fig)
        else:
            plt.show()
            plt.close(fig)


if __name__ == "__main__":
    app = SpectraApp()
    app.mainloop()
