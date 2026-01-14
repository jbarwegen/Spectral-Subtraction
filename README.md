Spectral-Subtraction (Orbitrap MS)
Home of the custom spectral subtraction application used for research at The King's University under Dr. Cassidy Vanderschee.

This repository hosts the source code and executable for a specialized tool designed to process Orbitrap Mass Spectrometry data. It performs spectral subtraction to compare different samples after they have been ran through the Orbitrap MS

You have two options to run this application
Option 1: Download the .exe and run

or

Option 2: Follow the following instructions

Step 1: Install Python If you don't have it, download and install Python from python.org.

Important: During installation, check the box that says "Add Python to PATH".

Step 2: Get the Files Create a new folder on your desktop (e.g., "SpectraApp"). You must put both of these files inside it:

spectra_app_NEWGUI.py (The script)

Spectra.ui (The layout file - Required)

Step 3: Install Libraries We need to install the tools the app uses.

Open your Start Menu, type cmd, and press Enter to open the Command Prompt.

Copy and paste the following line into the black box and hit Enter: pip install pandas matplotlib pyqt5 openpyxl

Step 4: Run the App

In the Command Prompt, type cd followed by a space.

Drag your "SpectraApp" folder from the desktop into the Command Prompt window (this creates the path) and hit Enter.

Type the following and hit Enter: python spectra_app_NEWGUI.py

Step 5: Using the App
Load Data: Click "Load File" and select your Excel file.

Excel Format: Your Excel sheet must have columns named: m/z, Intensity, Relative, Resolution, and Noise.

Plot: Select a sheet name from the list and click "Plot Graphs".
