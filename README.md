# Mass Spectrometry Spectral Subtraction App

### Developed by Jonathan Barwegen at The King's University 

This GUI application is designed to load, visualize, and perform spectral subtraction on Mass Spectrometry data. It was built specifically for research conducted under **Dr. Cassidy Vanderschee** to handle Orbitrap and similar high-resolution data formats.

## 🧪 Research Context

This tool was developed as part of research involving a custom spectral subtraction application for Orbitrap mass spectra at **The King's University**. For inquiries regarding the research or to contact the developer, please reach out to **Dr. Cassidy Vanderschee** at `cassidy.vanderschee@kingsu.ca`.

## 🚀 Features

- **Excel Data Loading**: Import multi-sheet Excel files with customizable row-skipping.
- **Spectral Subtraction**: Identifies unique peaks between two datasets using a vectorized comparison algorithm with a $3.0\text{ ppm}$ default tolerance.
- **Dual View**: Visualize "A subtracted B" and "B subtracted A" simultaneously in a mirrored plot.
- **Automatic Export**: Optional checkbox to automatically save subtracted peak lists as new Excel files.
- **Interactive Visualization**: View and annotate the top $n$ peaks for easy identification.

## 📊 Input Data Requirements

To ensure the application parses your data correctly, your Excel sheets must meet the following criteria:

- **File Format**: `.xlsx` or `.xls`.
- **Required Columns**: Each sheet must contain the following headers:
  - `m/z` 
  - `Intensity` 
  - `Relative` 
  - `Resolution` 
  - `Noise` 
- **Filtering**: The app automatically filters out low-signal peaks where $Intensity \le 10 \times Noise$.

## 🛠️ Installation & Setup

### Running the Executable

If you are using the standalone `.exe` version, simply ensure your data follows the requirements above and run the file. No Python installation is required for the executable version.

---

_Developed by Jonathan Barwegen as part of undergraduate research at The King's University._
