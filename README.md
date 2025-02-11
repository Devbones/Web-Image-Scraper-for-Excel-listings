"This project was developed with AI assistance using ChatGPT and other AI tools."
![AI-Assisted](https://img.shields.io/badge/AI-Assisted-blue?style=for-the-badge&logo=ai)

# Web-Image-Scraper-for-Excel-listings
Web Image Scraper automates product image retrieval by extracting codes from Excel, searching Google, and downloading the best-quality image from specified websites. It features a user-friendly GUI, speeds up e-commerce and inventory tasks, and eliminates manual searching for product images. It can search by user queries such as EAN codes/others🚀


## 📌 Overview  
**Web Image Scraper** automates product image retrieval by extracting product codes from an Excel file, searching Google for matching listings from specified websites, and downloading the best-quality product image. This tool is ideal for e-commerce, inventory management, and catalog creation.  

## ✨ Features  
- ✅ Extracts product codes from Excel  
- ✅ Automates Google searches for product listings  
- ✅ Finds and downloads the largest product image  
- ✅ User-friendly GUI for easy setup  
- ✅ Multi-threaded for fast processing  
- ✅ Saves time by eliminating manual searching  

## 🖥️ Installation  

1. **Clone the Repository:**  
   ```bash
   git clone https://github.com/Devbones/Web-Image-Scraper-for-Excel-listings
   cd listing-maker-3
   ```
2. **Install Dependencies:**  
   Ensure you have Python installed, then install the required libraries:  
   ```bash
   pip install -r requirements.txt
   ```
3. **Download ChromeDriver:**  
   - Ensure you have [ChromeDriver](https://sites.google.com/chromium.org/driver/) installed and placed in the project directory.  

## 🚀 Usage  

1. **Run the script:**  
   ```bash
   python listing_maker3.py
   ```
2. **Select an Excel file** containing product codes.  
3. **Choose a column** where product codes are stored.  
4. **Set the row range** (start and end rows).  
5. **Select an output folder** where images will be saved.  
6. **Enter target websites** (up to 4) where product listings are likely to be found.  
7. **Click "Start"** and let the script process the data!  

## 🛠️ Building a Standalone EXE  
To create an executable using PyInstaller, run:  
```bash
pyinstaller --onefile --noconsole listing_maker3.py
```
This will generate a standalone `.exe` file in the `dist` folder.

## 📜 License  
© 2025 Artur Kuśmirek. All rights reserved.  

