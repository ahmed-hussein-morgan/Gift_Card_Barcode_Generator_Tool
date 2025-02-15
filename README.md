# 🎁 **Gift Card Barcode Generator Tool** 🎁

## 📌 **Overview**
Welcome to the **Gift Card Barcode Generator Tool**! This Python-based tool is designed to help you generate serialized gift cards with dynamic barcodes, all in an easy-to-use, automated manner. With just an Excel sheet containing serial numbers and a gift card background image, the tool will generate a PowerPoint file with barcodes and serial numbers ready to print.

It’s perfect for anyone who needs to create and print a large batch of gift cards quickly and accurately. Whether you're working on a small project or handling large-scale production, this tool will save you time and effort.

## 🚀 **Key Features**
- **Excel Input:** Import serial numbers directly from an Excel file (XLSX).
- **Dynamic Barcode Generation:** Automatically creates high-quality barcodes based on each serial number.
- **Customizable Gift Card Background:** Choose a PNG or JPEG image as your gift card background.
- **PowerPoint Output:** Generate a PowerPoint (.pptx) file where each slide contains a serial number and barcode for easy printing.
- **Fast and Efficient:** Perfect for printing any number of gift cards, whether it's 10 or 1,000.

## 💡 **How It Works**
1. **Prepare Your Files:**
   - **Excel File:** Have your serial numbers in an Excel file (XLSX) with a column named "Serial Number".
   - **Background Image:** Place a PNG or JPEG image file of your gift card background in the same directory as the script.

2. **Run the Script:**
   - The tool will read the serial numbers from the Excel file and generate barcodes for each one.

3. **Generate PowerPoint:** 
   - Each serial number and its corresponding barcode will be added to a new slide in the PowerPoint file.
   - You can then print the slides to get your gift cards ready!

## 🛠️ **Installation Instructions**

### **Prerequisites:**
- Python 3.x
- Required Python libraries: `pandas`, `python-pptx`, `openpyxl`, `barcode`
  - Install these libraries using pip:
    ```bash
    pip install pandas python-pptx openpyxl barcode
    ```

### **Steps to Install and Use:**
#### **Windows:**
1. **Clone the Repository:**
   - Open **Command Prompt** or **PowerShell** and run:
     ```bash
     git clone https://github.com/ahmed-hussein-morgan/Gift_Card_Barcode_Generator_Tool.git
     ```
2. **Navigate to the Project Folder:**
   - Use `cd` to go into the project folder:
     ```bash
     cd gift-card-barcode-generator
     ```
3. **Install the Required Dependencies:**
   - In **Command Prompt** or **PowerShell**, run:
     ```bash
     pip install -r requirements.txt
     ```
4. **Prepare Your Files:**
   - Place your Excel file with serial numbers and your gift card background image (PNG or JPEG) in the same directory as the script.
   
5. **Run the Script:**
   - Run the script using Python:
     ```bash
     python generator.py
     ```

#### **Linux/Mac:**
1. **Clone the Repository:**
   - Open your terminal and run:
     ```bash
     git clone https://github.com/ahmed-hussein-morgan/Gift_Card_Barcode_Generator_Tool.git
     ```
2. **Navigate to the Project Folder:**
   - Use `cd` to go into the project folder:
     ```bash
     cd gift-card-barcode-generator
     ```
3. **Install the Required Dependencies:**
   - In your terminal, run:
     ```bash
     pip install -r requirements.txt
     ```
4. **Prepare Your Files:**
   - Place your Excel file with serial numbers and your gift card background image (PNG or JPEG) in the same directory as the script.
   
5. **Run the Script:**
   - Run the script using Python:
     ```bash
     python generator.py
     ```

## 📑 **Usage Guide**

1. **Excel File Format:**
   - Your Excel file should contain a column named "Serial Number" in first column , with each row representing a unique serial number for a gift card.
   
2. **Gift Card Background Image:**
   - Ensure the gift card background image (PNG or JPEG) is in the same directory as the script. The script will use this image to place the serial number and barcode.

3. **Running the Script:**
   - When you run the script, it will generate a PowerPoint file named `gift_cards.pptx` containing one slide for each serial number.
   
4. **Printing Your Gift Cards:**
   - Open the generated PowerPoint file, and you can print the slides directly or use it for further design editing.

## 🎨 **Customization Options**
- **Background Image:** Choose any PNG or JPEG file for the gift card background. Ensure it is the desired size for your cards.
- **Barcode Formats:** The script supports different barcode formats, such as Code128, for flexibility.
- **Font and Placement:** Currently, you can modify the font and placement of the serial number and barcode directly in the script if you need to adjust the appearance.

## 🔑 **Why This Tool?**
- **Quick and Easy:** Automates the barcode generation and PowerPoint creation process to save you time.
- **Batch Processing:** Efficiently handles any number of serial numbers.
- **No Need for a UI:** Operates entirely through script execution, perfect for advanced users who prefer minimal setups.
- **Scalable:** Whether you need 10 or 1,000 gift cards, this tool can handle it all.

## 💬 **Feedback & Support**
If you encounter any issues or have suggestions, feel free to open an issue here on GitHub. You can also reach out via email at [ahmed.h.morgan2050@gmail.com]. I'm happy to provide assistance!

## 🌟 **Contribute**
Contributions are welcome! If you’d like to improve the tool or add new features, feel free to fork the repo and submit a pull request. Please follow the coding standards for consistency.

## 📄 **License**
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for more details.
