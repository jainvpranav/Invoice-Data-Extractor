# Invoice Data Extractor 🧾

A custom **Invoice Data Extractor** designed to parse PDFs, extract specific data fields, and store the information in an Excel sheet. This tool streamlines the process of handling invoice data with minimal setup required. 

## 🛠 Features

- **PDF Parsing**: Utilizes `PDFMiner` and `MuPDF` for efficient and accurate PDF parsing.
- **Data Extraction**: Custom logic to extract relevant fields (e.g., invoice numbers, dates, amounts, etc.).
- **Excel Generation**: Automates the generation of structured Excel sheets using `Pandas`.
- **User-Friendly**: Comes as a ready-to-use `.exe` file—no Python setup required!
- **Customizable**: Adaptable to suit various invoice formats if required.

## 🚀 Technologies Used

- **PDFMiner**: For text extraction from PDFs.
- **MuPDF (PyMuPDF)**: For parsing and analyzing PDF content.
- **Pandas**: For manipulating and exporting data to Excel files.

## 🖥️ How to Use

1. **Download the Application**:  
   [Download Invoice Data Extractor.exe](https://github.com/jainvpranav/Invoice-Data-Extractor)  

2. **Prepare Your PDFs**:  
   - Place your PDF files in the same directory as the `.exe` file.
   - Ensure the PDFs are stored in a folder named **`PDFs`**.  

3. **Run the Application**:  
   - Double-click the `.exe` file.
   - The extracted data will be saved in an Excel file (`output.xlsx`) in the same directory.  

## 📁 Folder Structure

```plaintext
.
├── _internal              # Necessary Libraries for the application
├── InvoiceExtractor.exe   # Executable file for the application
├── PDFs/                  # Folder containing input PDF files
├── output.xlsx            # Generated Excel file
```

## 🔧 Customization

If you need to adapt the tool for specific invoice formats, you can check out the source code in the repository. Feel free to fork the project and modify it as needed.

## 🤝 Contributions

Contributions are welcome! Open issues or submit pull requests to enhance the tool.

## 📜 License

This project is licensed under the [MIT License](LICENSE).

---

Let me know if you want to add more details or refine further!
