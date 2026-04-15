# Mass Spectrometry Spectral Subtraction App

### [cite_start]Developed by Jonathan Barwegen at The King's University [cite: 1]

[cite_start]This GUI application is designed to load, visualize, and perform spectral subtraction on Mass Spectrometry data[cite: 1]. It was built specifically for research conducted under **Dr. [cite_start]Cassidy Vanderschee** to handle Orbitrap and similar high-resolution data formats[cite: 1].

## 🧪 Research Context

[cite_start]This tool was developed as part of research involving a custom spectral subtraction application for Orbitrap mass spectra at **The King's University**[cite: 1]. For inquiries regarding the research or to contact the developer, please reach out to **Dr. [cite_start]Cassidy Vanderschee** at `cassidy.vanderschee@kingsu.ca`[cite: 1].

## 🚀 Features

- **Excel Data Loading**: Import multi-sheet Excel files with customizable row-skipping[cite: 1].
- [cite_start]**Spectral Subtraction**: Identifies unique peaks between two datasets using a vectorized comparison algorithm with a $3.0\text{ ppm}$ default tolerance[cite: 1].
- [cite_start]**Dual View**: Visualize "A subtracted B" and "B subtracted A" simultaneously in a mirrored plot[cite: 1].
- **Automatic Export**: Optional checkbox to automatically save subtracted peak lists as new Excel files[cite: 1].
- [cite_start]**Interactive Visualization**: View and annotate the top $n$ peaks for easy identification[cite: 1].

## 📊 Input Data Requirements

[cite_start]To ensure the application parses your data correctly, your Excel sheets must meet the following criteria[cite: 1]:

- **File Format**: `.xlsx` or `.xls`[cite: 1].
- **Required Columns**: Each sheet must contain the following headers:
  - `m/z` [cite: 1]
  - [cite_start]`Intensity` [cite: 1]
  - [cite_start]`Relative` [cite: 1]
  - `Resolution` [cite: 1]
  - [cite_start]`Noise` [cite: 1]
- [cite_start]**Filtering**: The app automatically filters out low-signal peaks where $Intensity \le 10 \times Noise$[cite: 1].

## 🛠️ Installation & Setup

### Running from Source

1.  **Clone the Repository**:
    ```bash
    git clone [https://github.com/jbarwegen/tidytuesday-blog](https://github.com/jbarwegen/tidytuesday-blog)
    ```
    [cite_start]_(Note: Update the URL above if you move this project to a new repository[cite: 1].)_
2.  **Install Dependencies**:
    ```bash
    pip install pandas matplotlib pyqt5 openpyxl numpy
    ```
3.  [cite_start]**Run the App**: Ensure `Spectra.ui` is in the same directory as the script[cite: 1].
    ```bash
    python spectra_app_Updated.py
    ```

### Running the Executable

If you are using the standalone `.exe` version, simply ensure your data follows the requirements above and run the file. [cite_start]No Python installation is required for the executable version[cite: 1].

---

[cite_start]_Developed by Jonathan Barwegen as part of undergraduate research at The King's University[cite: 1]._
