PPTView 🎨
 
A powerful and elegant PPT preview & export tool built with PyQt5, designed for high-fidelity slide visualization and flexible export.
 
 
 
✨ Features
 
- 🖼️ High-Fidelity Preview: Convert PPT/PPTX slides to crisp PNG images and display them in a clean grid layout.
- 📄 Flexible Export:
- Save all/single/multiple slides as PNG images
- Export to a merged PDF or separate single-page PDFs
- 🖱️ Intuitive Interaction:
- Double-click to zoom in for detailed preview
- Right-click context menu for quick operations
- Drag-and-drop file import
- Click-to-select multi-selection mode (no Ctrl needed)
- ⚙️ Customizable Settings: Adjust export resolution to meet your quality needs.
- 🎨 Modern UI: Clean and responsive interface with smooth animations.
 
 
 
🚀 Quick Start
 
1. Install Dependencies
 
bash
  
pip install pyqt5 pillow pywin32
 
 
2. Run the Application
 
bash
  
python ppt_viewer.py
 
 
3. Import a PPT File
 
- Method 1: Click  Import PPT File  in the toolbar and select your  .ppt  or  .pptx  file.
- Method 2: Simply drag and drop a PPT file into the application window.
 
4. Preview & Export
 
- Preview: Double-click any thumbnail to open a zoomable preview window.
- Select: Use  Ctrl/Shift  for multi-selection, or enable  Click Multi-Selection  in the toolbar.
- Export: Right-click any thumbnail to access the export menu, or use the toolbar buttons.
 
 
 
📦 Build & Distribute
 
To package the application into a standalone executable:
 
bash
  
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico --name="PPTView" ppt_viewer.py
 
 
- The executable will be generated in the  dist  folder.
- Replace  icon.ico  with your own icon file for a branded look.
 
 
 
📝 Usage Guide
 
Action How to Use 
Import PPT Click  Import PPT File  or drag-and-drop a file. 
Zoom Preview Double-click on any slide thumbnail. 
Multi-Selection Hold  Ctrl/Shift  or enable  Click Multi-Selection  in the toolbar. 
Export as Images Right-click →  Save All  or  Save Selected . 
Export as PDF Right-click →  Export All as PDF  or  Export Selected as PDF . 
Adjust Resolution Click  Set Export Resolution  in the toolbar. 
 
 
 
🤝 Contributing
 
Feel free to fork this repository, create a feature branch, and submit a pull request!
 
 
 
📄 License
 
This project is licensed under the MIT License.
 
 
 
Would you like me to also add a screenshot section to this README, showing the tool in action with the "hypertension" PPT example from your image? This would make it even more visually appealing for GitHub visitors.