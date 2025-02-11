# Web-Image-Scraper-for-Excel-listings
Web Image Scraper automates product image retrieval by extracting codes from Excel, searching Google, and downloading the best-quality image from specified websites. It features a user-friendly GUI, speeds up e-commerce and inventory tasks, and eliminates manual searching for product images. It can search by user queries such as EAN codes/others🚀


📌 Overview
Web Image Scraper automates product image retrieval by extracting product codes from an Excel file, searching Google for matching listings from specified websites, and downloading the best-quality product image. This tool is ideal for e-commerce, inventory management, and catalog creation.

✨ Features
✅ Extracts product codes from Excel
✅ Automates Google searches for product listings
✅ Finds and downloads the largest product image
✅ User-friendly GUI for easy setup
✅ Multi-threaded for fast processing
✅ Saves time by eliminating manual searching
🖥️ Installation
Clone the Repository:
bash
Kopiuj
Edytuj
git clone https://github.com/yourusername/listing-maker-3.git
cd listing-maker-3
Install Dependencies:
Ensure you have Python installed, then install the required libraries:
bash
Kopiuj
Edytuj
pip install -r requirements.txt
Download ChromeDriver:
Ensure you have ChromeDriver installed and placed in the project directory.
🚀 Usage
Run the script:
bash
Kopiuj
Edytuj
python listing_maker3.py
Select an Excel file containing product codes.
Choose a column where product codes are stored.
Set the row range (start and end rows).
Select an output folder where images will be saved.
Enter target websites (up to 4) where product listings are likely to be found.
Click "Start" and let the script process the data!
🛠️ Building a Standalone EXE
To create an executable using PyInstaller, run:

bash
Kopiuj
Edytuj
pyinstaller --onefile --noconsole listing_maker3.py
This will generate a standalone .exe file in the dist folder.

📜 License
© 2025 Artur Kuśmirek. All rights reserved.

This README ensures a professional and informative GitHub repository. Let me know if you need adjustments! 🚀
