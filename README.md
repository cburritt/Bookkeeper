🧾 Fraternity Bookkeeper App
Overview

The Fraternity Bookkeeper App is a C#-based tool built to simplify expense tracking and record-keeping within a fraternity. The app allows members to easily submit pictures of receipts, which are automatically processed using OCR (Optical Character Recognition) to extract key information such as the date, department, and total cost.

Once processed, the data is formatted and written directly into an Excel file, making it easy for the fraternity treasurer to manage and review financial transactions without needing to manually enter receipt data.

This project was created to make fraternity budgeting more organized, efficient, and transparent.

✨ Features

Receipt Uploads — Members can take or upload photos of receipts directly into the app.

OCR Text Recognition — The app automatically reads and extracts relevant information from receipt images using the Tesseract OCR engine.

Excel Integration — Extracted data is written into an Excel sheet using ClosedXML, ready for the treasurer’s review.

Receipt Search Functions — Users can search receipts by month, year, or department.

Data Collection from Excel — The app can pull existing receipt records from an Excel sheet into its local database for review or editing.

Excel Sorting — Automatically sorts receipts in the Excel file by date for easier tracking.

⚙️ Tech Stack

Language: C# (.NET)

Libraries & Tools:

Tesseract OCR (for image-to-text conversion)

ClosedXML (for Excel integration)

Regular Expressions (for identifying dates, totals, and keywords)

Storage: Excel workbook (.xlsx)

Platform: Console-based (with GUI planned for future versions)

🧠 How It Works

A fraternity member uploads or takes a picture of their receipt.

The app processes the image using Tesseract OCR to extract raw text.

It automatically identifies the date, department, and total cost using regular expressions.

This information is written into an Excel spreadsheet, where each row represents one receipt.

The treasurer can then easily review, sort, and manage all entries in Excel.

This process saves time, reduces human error, and provides a clear, digital record of all fraternity expenses.

🧩 Installation

Install the .NET SDK and Tesseract OCR on your system.

Clone this repository to your local machine.

Ensure a tessdata folder (containing English OCR training data) is included in the project directory.

Run the application using dotnet run.

Receipt data will automatically be written to an Excel file in the project’s output directory.

💡 Future Updates

Redesigned OCR System with Machine Learning:
The next major update will replace the current OCR-based text reader with a custom machine learning model specifically trained on receipt images. This model will improve accuracy in reading faded, wrinkled, or partially obscured receipts.

Graphical User Interface (GUI):
A full GUI version of the Bookkeeper App is currently in development. The interface will feature drag-and-drop receipt uploads, live OCR previews, and easy access to stored records—all without using the console.



Author: Caden Burritt

Purpose: A fraternity financial management tool designed to make receipt collection and expense tracking faster, smarter, and easier.
