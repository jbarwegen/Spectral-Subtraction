import pandas as pd
import matplotlib.pyplot as plt
import os
import sys
from PyQt5 import uic, QtWidgets as qw,QtCore as qc
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
    uic.loadUi(self.resource_path("Spectra.ui"),self)
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
    
    
    


####################
# ACTIONS
  def choose_save_location(self) -> None:
        path = qw.QFileDialog.getExistingDirectory(
            self,
            "Choose file save location",
            "",
            qw.QFileDialog.ShowDirsOnly,
        )
        if path:
            self.save_path = path
            self.saveLocationLineEdit.setText(self.save_path)


  
####################
#HELPER FUNCTIONS  
#       
  def load_excel_file(self) -> None:
    file_path, _ = qw.QFileDialog.getOpenFileName(
        self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
    )
    if not file_path:
        return
    try:
        skip_rows = self.rowSkipSpinBox.value()
        names, data = self.load_data(skip_rows, file_path)

        self.excel_path = file_path
        self.sheet_names = names
        self.data_by_sheet = data

        self.mainSpectraBox.clear(); self.mainSpectraBox.addItems(names)
        self.subtractBox.clear(); self.subtractBox.addItems(names)
        self.graphsWidget.clear(); self.graphsWidget.addItems(names)
        self.spectraABox.clear(); self.spectraABox.addItems(names)
        self.spectraBBox.clear(); self.spectraBBox.addItems(names)

        qw.QMessageBox.information(self, "Loaded", f"Loaded {len(names)} sheets from\n{file_path}")
    except Exception as e:  # 
        qw.QMessageBox.warning(self, "Error", f"Failed to load Excel file:\n{e}")
  def _on_plot_selected_item(self, item: qw.QListWidgetItem) -> None:
        self._plot_single_sheet(item.text())
  def _on_plot_graphs_clicked(self) -> None:
        items = self.graphsWidget.selectedItems() or [self.graphsWidget.currentItem()]
        if not items or not items[0]:
            qw.QMessageBox.information(self, "Select a sheet", "Choose a sheet in the list to plot.")
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

        n = self._get_peaks_to_annotate()
        df_main = self._maybe_normalize(self.data_by_sheet[main_name])
        df_sub = self._maybe_normalize(self.data_by_sheet[sub_name])
        title = f"{main_name} subtracted {sub_name}"
        #self.plot_dual_spectrum(df_main, df_sub, title=title, n_peaks=n)

        unique_df = self.compare_dfs(self.data_by_sheet[main_name], self.data_by_sheet[sub_name])
        unique_df=self._maybe_normalize(unique_df)
        self.plot_spectrum(unique_df,title,n_peaks=n)
            
 
        
  def _on_dual_clicked(self) -> None:
      main_name = self.spectraABox.currentText()
      sub_name = self.spectraBBox.currentText()
      if not main_name or not sub_name:
            qw.QMessageBox.warning(self, "Select sheets", "Select both main and subtract sheets.")
            return
      if main_name not in self.data_by_sheet or sub_name not in self.data_by_sheet:
            qw.QMessageBox.warning(self, "Data missing", "Selected sheets not loaded.")
            return

      n = self._get_peaks_to_annotate()
      df_main = self._maybe_normalize(self.data_by_sheet[main_name])
      df_sub = self._maybe_normalize(self.data_by_sheet[sub_name])
      title = f"{main_name} subtracted {sub_name}"
      df_main = self.compare_dfs(df_main,df_sub)
      df_main =self._maybe_normalize(df_main)
      df_sub = self.compare_dfs(df_sub,df_main)
      df_sub = self._maybe_normalize(df_sub)
      self.plot_dual_spectrum(df_main, df_sub, title=title, n_peaks=n)

  @staticmethod
  def load_data(skip_rows: int, path: str) -> Tuple[List[str], Dict[str, pd.DataFrame]]:
        xls = pd.ExcelFile(path)
        names = xls.sheet_names
        raw = pd.read_excel(xls, sheet_name=names, skiprows=skip_rows)
        filtered: Dict[str, pd.DataFrame] = {}
        for name, df in raw.items():
            required = {"m/z", "Intensity", "Relative", "Resolution", "Noise"}
            missing = required - set(df.columns)
            if missing:
                raise ValueError(f"Sheet '{name}' is missing columns: {sorted(missing)}")
            keep = df.loc[df["Intensity"] > 10 * df["Noise"]].copy()
            for col in ("m/z", "Intensity", "Relative", "Resolution", "Noise"):
                keep[col] = pd.to_numeric(keep[col], errors="coerce")
            keep = keep.dropna(subset=["m/z", "Relative", "Resolution"]).reset_index(drop=True)
            filtered[name] = keep
        return names, filtered

  def _maybe_normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        normalized = df.copy()
        if self.toggleNormalization.isChecked():
            max_rel = normalized["Relative"].max()
            if pd.notna(max_rel) and max_rel > 0:
                normalized["Relative"] = normalized["Relative"] / max_rel * 100.0
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
        ax.vlines(df["m/z"], 0, df["Relative"],colors="black")
        ax.set_title(title)
        ax.set_xlabel("m/z")
        ax.set_ylabel("Relative")
        ax.set_xlim(right=max(350, float(df["m/z"].max()) if not df.empty else 350))
        ax.set_ylim(bottom=0, top=115)

        for _, r in top.iterrows():
            ax.annotate(f"{r['m/z']:.4f}", xy=(r["m/z"], r["Relative"]), xytext=(0, 5),
                        textcoords="offset points", ha="center", va="bottom", rotation=45, fontsize=8)
        plt.tight_layout()

        if self._should_save_graphs():
            filename = f"{title.replace(' ', '_')}.svg"
            filepath = os.path.join(self.save_path or "", filename)
            fig.savefig(filepath)
            plt.close(fig)
            qw.QMessageBox.information(self, "Saved", f"Figure saved to:\n{os.path.abspath(filepath)}")
        else:
          
            plt.show()
            

  def plot_dual_spectrum(self, df_up: pd.DataFrame, df_down: pd.DataFrame, title: str, n_peaks: int = 10) -> None:
        top_up = df_up.nlargest(n_peaks, "Relative") if not df_up.empty else df_up
        top_down = df_down.nlargest(n_peaks, "Relative") if not df_down.empty else df_down

        fig, ax = plt.subplots(figsize=(10, 5))
        ax.vlines(df_up["m/z"], 0, df_up["Relative"],colors="#13f034")
        ax.vlines(df_down["m/z"], 0, -df_down["Relative"], colors="#f51c0c")
        ax.set_title(title)
        ax.set_xlabel("m/z")
        ax.set_ylabel("Relative")
        xmax = max(350, float(df_up["m/z"].max()) if not df_up.empty else 0,
                   float(df_down["m/z"].max()) if not df_down.empty else 0)
        ax.set_xlim(right=xmax)
        ax.set_ylim(-130, 130)
        ax.axhline(0, linewidth=1)

        for _, r in top_up.iterrows():
            ax.annotate(f"{r['m/z']:.4f}", xy=(r["m/z"], r["Relative"]), xytext=(0, 5),
                        textcoords="offset points", ha="center", va="bottom", rotation=45, fontsize=8)
        for _, r in top_down.iterrows():
            ax.annotate(f"{r['m/z']:.4f}", xy=(r["m/z"], -r["Relative"]), xytext=(0, -5),
                        textcoords="offset points", ha="center", va="top", rotation=45, fontsize=8)

        plt.tight_layout()
        if self._should_save_graphs():
            filename = f"{title.replace(' ', '_')}_dual.svg"
            filepath = os.path.join(self.save_path or "", filename)
            fig.savefig(filepath)
            plt.close(fig)
            qw.QMessageBox.information(self, "Saved", f"Figure saved to:\n{os.path.abspath(filepath)}")
        else:
            plt.show()
            

  @staticmethod
  def compare_dfs(df1: pd.DataFrame, df2: pd.DataFrame, ppm_tol: float = 3.0) -> pd.DataFrame:
        if df1.empty:
            return df1.copy()
        if df2.empty:
            return df1.dropna(subset=["m/z"]).reset_index(drop=True)

        dfB = df2.copy()
        dfB["half_width_B"] = dfB["m/z"] / dfB["Resolution"] / 2

        def peak_match(row: pd.Series) -> bool:
            m1 = float(row["m/z"])
            R1 = float(row["Resolution"]) if pd.notna(row["Resolution"]) else float("inf")
            hw1 = m1 / R1 / 2 if R1 and R1 > 0 else 0.0
            max_hwB = float(dfB["half_width_B"].max()) if not dfB.empty else 0.0
            lo, hi = m1 - (hw1 + max_hwB), m1 + (hw1 + max_hwB)
            cand = dfB[(dfB["m/z"] >= lo) & (dfB["m/z"] <= hi)]
            if cand.empty:
                return False
            sep = (cand["m/z"] - m1).abs()
            overlap = sep <= (hw1 + cand["half_width_B"])
            delta_ppm = sep / ((cand["m/z"] + m1) / 2.0) * 1e6
            return bool((overlap & (delta_ppm <= ppm_tol)).any())

        dfA = df1.dropna(subset=["m/z"]).copy()
        mask = dfA.apply(peak_match, axis=1)
        return dfA.loc[~mask].reset_index(drop=True)
def main() -> int:
  if hasattr(qc.Qt, 'AA_EnableHighDpiScaling'):
      qw.QApplication.setAttribute(qc.Qt.AA_EnableHighDpiScaling, True)
  if hasattr(qc.Qt, 'AA_UseHighDpiPixmaps'):
      qw.QApplication.setAttribute(qc.Qt.AA_UseHighDpiPixmaps, True)  
  app = qw.QApplication(sys.argv)
  
   
  window = SpectraSubtractionApp()
  window.show()
  return app.exec_()


if __name__ == "__main__":
    sys.exit(main())