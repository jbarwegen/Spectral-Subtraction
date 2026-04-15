"""
SPECTRA SUBTRACTION APP
-----------------------
Description:
    A GUI application to load Mass Spectrometry data (Excel), visualize spectra,
    and perform spectral subtraction.


INPUT DATA REQUIREMENTS:
    - Input must be an Excel file (.xlsx).
    - Sheets must contain these columns: 'm/z', 'Intensity', 'Relative', 'Resolution', 'Noise'.
"""

############################
# This code was developed by Jonathan Barwegen at The Kings University so if you have any questions please reach
# to Cassidy Vanderschee at The King's University to contact me: cassidy.vanderschee@kingsu.ca

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import sys
from PyQt6 import uic, QtWidgets as qw, QtCore as qc
from typing import Dict, List, Tuple

class SpectraSubtractionApp(qw.QMainWindow):
    @staticmethod
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def __init__(self):
        super().__init__()
        uic.loadUi(self.resource_path("Spectra2.ui"), self)
        
        self.save_path: str = ""
        self.excel_path: str = ""
        self.sheet_names: List[str] = []
        self.data_by_sheet: Dict[str, pd.DataFrame] = {}
        
        # Wire up required UI
        self.rowSkipSpinBox.setValue(6)
        self.peaksAnnotate.setValue(10)
        self.selectFolderButton.clicked.connect(self.choose_save_location)
        self.loadfileButton.clicked.connect(self.load_excel_file)
        self.plotGraphs.clicked.connect(self._on_plot_graphs_clicked)
        self.plotSubtractionButton.clicked.connect(self._on_plot_subtraction_clicked)
        self.graphsWidget.itemActivated.connect(self._on_plot_selected_item)
        self.plotDualButton.clicked.connect(self._on_dual_clicked)
        
        # Optional: Wire up export button if you want a dedicated general export
        # self.exportDataButton.clicked.connect(self._on_export_data_clicked)

    ####################
    # ACTIONS
    
    def choose_save_location(self) -> None:
        """Opens a file dialog to select the directory where spectra images/data will be saved."""
        path = qw.QFileDialog.getExistingDirectory(
            self,
            "Choose file save location",
            "",
            qw.QFileDialog.Option.ShowDirsOnly,
        )
        if path:
            self.save_path = path
            self.saveLocationLineEdit.setText(self.save_path)

    ####################
    # HELPER FUNCTIONS

    def load_excel_file(self) -> None:
        """Opens a file dialog to select and load an Excel file."""
        file_path, _ = qw.QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if not file_path:
            return
        try:
            qw.QApplication.setOverrideCursor(qc.Qt.CursorShape.WaitCursor)
            skip_rows = self.rowSkipSpinBox.value()
            names, data = self.load_data(skip_rows, file_path)

            self.excel_path = file_path
            self.sheet_names = names
            self.data_by_sheet = data

            # Update UI elements
            for box in [self.mainSpectraBox, self.subtractBox, self.spectraABox, self.spectraBBox]:
                box.clear()
                box.addItems(names)
                
            self.graphsWidget.clear()
            self.graphsWidget.addItems(names)

            qw.QApplication.restoreOverrideCursor()
            qw.QMessageBox.information(
                self, "Loaded", f"Loaded {len(names)} sheets from\n{file_path}"
            )
        except Exception as e:
            qw.QApplication.restoreOverrideCursor()
            qw.QMessageBox.warning(self, "Error", f"Failed to load Excel file:\n{e}")

    def _on_plot_selected_item(self, item: qw.QListWidgetItem) -> None:
        self._plot_single_sheet(item.text())

    def _on_plot_graphs_clicked(self) -> None:
        items = self.graphsWidget.selectedItems() or [self.graphsWidget.currentItem()]
        if not items or not items[0]:
            qw.QMessageBox.information(
                self, "Select a sheet", "Choose a sheet in the list to plot."
            )
            return
        for item in items:
            self._plot_single_sheet(item.text())

    def _on_plot_subtraction_clicked(self) -> None:
        main_name = self.mainSpectraBox.currentText()
        sub_name = self.subtractBox.currentText()
        
        if not main_name or not sub_name:
            qw.QMessageBox.warning(self, "Select sheets", "Select both A and B sheets.")
            return
        if main_name not in self.data_by_sheet or sub_name not in self.data_by_sheet:
            qw.QMessageBox.warning(self, "Data missing", "Selected sheets not loaded.")
            return

        qw.QApplication.setOverrideCursor(qc.Qt.CursorShape.WaitCursor)
        n = self._get_peaks_to_annotate()
        
        # 1. Get exclusion list and base data
        exclude_list = self._get_excluded_mz()
        df_main = self.data_by_sheet[main_name]
        df_sub = self.data_by_sheet[sub_name]
        
        # 2. Strip out the manually excluded peaks
        df_main = self._remove_specific_peaks(df_main, exclude_list)
        df_sub = self._remove_specific_peaks(df_sub, exclude_list)

       # 3. Proceed with existing subtraction logic
        unique_df = self.compare_dfs(df_main, df_sub)
        unique_df = self._maybe_normalize(unique_df)
        title = f"{main_name} subtracted {sub_name}"
        
        self._export_to_excel(unique_df, f"{title.replace(' ', '_')}.xlsx")
        qw.QApplication.restoreOverrideCursor()
        
        self.plot_spectrum(unique_df, title, n_peaks=n)

    def _on_dual_clicked(self) -> None:
        main_name = self.spectraABox.currentText()
        sub_name = self.spectraBBox.currentText()
        
        if not main_name or not sub_name:
            qw.QMessageBox.warning(self, "Select sheets", "Select both main and subtract sheets.")
            return
        if main_name not in self.data_by_sheet or sub_name not in self.data_by_sheet:
            qw.QMessageBox.warning(self, "Data missing", "Selected sheets not loaded.")
            return

        qw.QApplication.setOverrideCursor(qc.Qt.CursorShape.WaitCursor)
        n = self._get_peaks_to_annotate()
        
        # 1. Get the original datasets
        df_main_orig = self.data_by_sheet[main_name]
        df_sub_orig = self.data_by_sheet[sub_name]
        title = f"{main_name} subtracted {sub_name}"
        
        # 2. Compare against the originals to find unique peaks for both
        df_main_unique = self.compare_dfs(df_main_orig, df_sub_orig)
        df_sub_unique = self.compare_dfs(df_sub_orig, df_main_orig)
        
        # 3. Normalize the newly subtracted datasets
        df_main_final = self._maybe_normalize(df_main_unique)
        df_sub_final = self._maybe_normalize(df_sub_unique)
        
        # 4. Export the data if checked
        self._export_dual_to_excel(
            df_main_final, main_name, 
            df_sub_final, sub_name, 
            f"{title.replace(' ', '_')}_dual.xlsx"
        )
        qw.QApplication.restoreOverrideCursor()
        
        # 5. Plot the graph
        self.plot_dual_spectrum(df_main_final, df_sub_final, title=title, n_peaks=n)

    @staticmethod
    def load_data(skip_rows: int, path: str) -> Tuple[List[str], Dict[str, pd.DataFrame]]:
        xls = pd.ExcelFile(path)
        names = xls.sheet_names
        raw = pd.read_excel(xls, sheet_name=names, skiprows=skip_rows)
        filtered: Dict[str, pd.DataFrame] = {}
        
        required = {"m/z", "Intensity", "Relative", "Resolution", "Noise"}
        
        for name, df in raw.items():
            missing = required - set(df.columns)
            if missing:
                raise ValueError(f"Sheet '{name}' is missing columns: {sorted(missing)}")
            
            # Vectorized type conversion and filtering
            for col in required:
                df[col] = pd.to_numeric(df[col], errors="coerce")
                
            keep = df[df["Intensity"] > 10 * df["Noise"]].copy()
            keep = keep.dropna(subset=["m/z", "Relative", "Resolution"]).reset_index(drop=True)
            filtered[name] = keep
            
        return names, filtered

    def _maybe_normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty or not self.toggleNormalization.isChecked():
            return df.copy()
            
        normalized = df.copy()
        max_rel = normalized["Relative"].max()
        if pd.notna(max_rel) and max_rel > 0:
            normalized["Relative"] = (normalized["Relative"] / max_rel) * 100.0
        return normalized

    def _get_peaks_to_annotate(self) -> int:
        return int(self.peaksAnnotate.value())

    def _should_save_graphs(self) -> bool:
        return bool(self.saveGraphBox.isChecked())

    def _plot_single_sheet(self, name: str) -> None:
        if name not in self.data_by_sheet:
            qw.QMessageBox.warning(self, "Not found", f"Sheet '{name}' not loaded.")
            return
        df = self._maybe_normalize(self.data_by_sheet[name])
        self.plot_spectrum(df=df, title=name, n_peaks=self._get_peaks_to_annotate())

    def plot_spectrum(self, df: pd.DataFrame, title: str, n_peaks: int = 10) -> None:
        top = df.nlargest(n_peaks, "Relative") if not df.empty else df
        fig, ax = plt.subplots(figsize=(10, 5))
        
        if not df.empty:
            ax.vlines(df["m/z"], 0, df["Relative"], colors="black")
            ax.set_xlim(right=max(350, float(df["m/z"].max())))
            
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
                
        ax.set_title(title)
        ax.set_xlabel("m/z")
        ax.set_ylabel("Relative")
        ax.set_ylim(bottom=0, top=115)
        plt.tight_layout()

        if self._should_save_graphs():
            filename = f"{title.replace(' ', '_')}.svg"
            filepath = os.path.join(self.save_path or "", filename)
            fig.savefig(filepath)
            plt.close(fig)
            qw.QMessageBox.information(
                self, "Saved", f"Figure saved to:\n{os.path.abspath(filepath)}"
            )
        else:
            plt.show()

    def plot_dual_spectrum(self, df_up: pd.DataFrame, df_down: pd.DataFrame, title: str, n_peaks: int = 10) -> None:
        fig, ax = plt.subplots(figsize=(10, 5))
        xmax = 350

        if not df_up.empty:
            top_up = df_up.nlargest(n_peaks, "Relative")
            ax.vlines(df_up["m/z"], 0, df_up["Relative"], colors="#13f034")
            xmax = max(xmax, float(df_up["m/z"].max()))
            for _, r in top_up.iterrows():
                ax.annotate(
                    f"{r['m/z']:.4f}", xy=(r["m/z"], r["Relative"]),
                    xytext=(0, 5), textcoords="offset points",
                    ha="center", va="bottom", rotation=45, fontsize=8
                )

        if not df_down.empty:
            top_down = df_down.nlargest(n_peaks, "Relative")
            ax.vlines(df_down["m/z"], 0, -df_down["Relative"], colors="#f51c0c")
            xmax = max(xmax, float(df_down["m/z"].max()))
            for _, r in top_down.iterrows():
                ax.annotate(
                    f"{r['m/z']:.4f}", xy=(r["m/z"], -r["Relative"]),
                    xytext=(0, -5), textcoords="offset points",
                    ha="center", va="top", rotation=45, fontsize=8
                )

        ax.set_title(title)
        ax.set_xlabel("m/z")
        ax.set_ylabel("Relative")
        ax.set_xlim(right=xmax)
        ax.set_ylim(-130, 130)
        ax.axhline(0, linewidth=1, color='black')

        plt.tight_layout()
        if self._should_save_graphs():
            filename = f"{title.replace(' ', '_')}_dual.svg"
            filepath = os.path.join(self.save_path or "", filename)
            fig.savefig(filepath)
            plt.close(fig)
            qw.QMessageBox.information(
                self, "Saved", f"Figure saved to:\n{os.path.abspath(filepath)}"
            )
        else:
            plt.show()

    def _get_excluded_mz(self) -> List[float]:
        """Parses the text input into a list of floats."""
        if not hasattr(self, 'excludePeaksInput'):
            return []
        
        text = self.excludePeaksInput.text().strip()
        if not text:
            return []
            
        try:
            # Replace commas with spaces, split, and convert to float
            import re
            parts = re.split(r'[,\s]+', text)
            return [float(p) for p in parts if p]
        except ValueError:
            qw.QMessageBox.warning(
                self, "Input Error", 
                "Invalid m/z values in exclusion list. Please use numbers separated by commas."
            )
            return []

    def _remove_specific_peaks(self, df: pd.DataFrame, exclude_list: List[float], ppm_tol: float = 3.0) -> pd.DataFrame:
        """Removes specific m/z values from a dataframe within a given PPM tolerance."""
        if df.empty or not exclude_list:
            return df.copy()

        m_A = df["m/z"].to_numpy()
        mask = np.zeros(len(df), dtype=bool)

        # Flag any peak that falls within the PPM tolerance of our excluded list
        for ex_mz in exclude_list:
            delta_ppm = (np.abs(m_A - ex_mz) / ((m_A + ex_mz) / 2.0)) * 1e6
            mask |= (delta_ppm <= ppm_tol)

        # Return the dataframe, keeping only the peaks that were NOT flagged
        return df.loc[~mask].reset_index(drop=True)    

    @staticmethod
    def compare_dfs(df1: pd.DataFrame, df2: pd.DataFrame, ppm_tol: float = 3.0) -> pd.DataFrame:
        """
        Optimized, fully vectorized comparison using NumPy to calculate mass differences 
        and peak overlaps drastically faster than a row-by-row pd.apply.
        """
        if df1.empty: return df1.copy()
        if df2.empty: return df1.dropna(subset=["m/z"]).reset_index(drop=True)

        dfA = df1.dropna(subset=["m/z"]).copy()
        dfB = df2.dropna(subset=["m/z"]).copy()

        # Extract underlying numpy arrays for speed
        m_A = dfA["m/z"].to_numpy()
        R_A = dfA["Resolution"].fillna(float('inf')).to_numpy()
        hw_A = np.where(R_A > 0, m_A / R_A / 2.0, 0.0)

        m_B = dfB["m/z"].to_numpy()
        R_B = dfB["Resolution"].fillna(float('inf')).to_numpy()
        hw_B = np.where(R_B > 0, m_B / R_B / 2.0, 0.0)

        # Sort B for fast window searching
        sort_idx = np.argsort(m_B)
        m_B_sorted = m_B[sort_idx]
        hw_B_sorted = hw_B[sort_idx]

        mask = np.zeros(len(dfA), dtype=bool)
        max_hwB = hw_B.max() if len(hw_B) > 0 else 0.0

        for i in range(len(m_A)):
            m1 = m_A[i]
            hw1 = hw_A[i]
            
            # Determine search boundaries
            search_window = hw1 + max_hwB + (m1 * ppm_tol / 1e6) + 0.1 
            left = np.searchsorted(m_B_sorted, m1 - search_window, side='left')
            right = np.searchsorted(m_B_sorted, m1 + search_window, side='right')

            if left < right:
                cand_m = m_B_sorted[left:right]
                cand_hw = hw_B_sorted[left:right]
                
                sep = np.abs(cand_m - m1)
                overlap = sep <= (hw1 + cand_hw)
                delta_ppm = (sep / ((cand_m + m1) / 2.0)) * 1e6
                
                if np.any(overlap | (delta_ppm <= ppm_tol)):
                    mask[i] = True

        return dfA.loc[~mask].reset_index(drop=True)

    def _export_to_excel(self, df: pd.DataFrame, filename: str) -> None:
        if not hasattr(self, 'exportDataBox') or not self.exportDataBox.isChecked() or df.empty:
            return

        filepath = os.path.join(self.save_path or "", filename)
        try:
            df.to_excel(filepath, index=False)
            qw.QMessageBox.information(
                self, "Data Exported", f"Data saved to:\n{os.path.abspath(filepath)}"
            )
        except Exception as e:
            qw.QMessageBox.warning(self, "Export Error", f"Failed to save Excel file:\n{e}")
            
    def _export_dual_to_excel(self, df1: pd.DataFrame, name1: str, df2: pd.DataFrame, name2: str, filename: str) -> None:
        if not hasattr(self, 'exportDataBox') or not self.exportDataBox.isChecked():
            return
            
        if df1.empty and df2.empty:
            return

        filepath = os.path.join(self.save_path or "", filename)
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                if not df1.empty:
                    df1.to_excel(writer, sheet_name=f"Unique to {name1}"[:31], index=False)
                if not df2.empty:
                    df2.to_excel(writer, sheet_name=f"Unique to {name2}"[:31], index=False)
            
            qw.QMessageBox.information(
                self, "Data Exported", f"Dual data saved to:\n{os.path.abspath(filepath)}"
            )
        except Exception as e:
            qw.QMessageBox.warning(self, "Export Error", f"Failed to save dual Excel file:\n{e}")      

def main() -> int:
    app = qw.QApplication(sys.argv)
    window = SpectraSubtractionApp()
    window.show()
    return app.exec()  

if __name__ == "__main__":
    sys.exit(main())