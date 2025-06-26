# ⚙️ Economax Carport Quote Tool

## 🔧 Overview  
A Python-based tool designed to automate the generation of bills of materials (BOM) for Economax-style solar carport structures. This quoting tool simplifies the estimation process by accepting basic project inputs and producing a fully structured CSV file, compatible with SAGE Intacct or Excel workflows.

## 🚀 Features  
- **Inputs:**
  - Panel width and height  
  - Total number of panels  
  - Tilt angle  
  - Row and bay configuration  
  - Module orientation, spacing preference, and bracing option  

- **Calculates:**
  - Beam and purlin quantities  
  - Connector brackets, anchor bolts, and baseplates  
  - Row spacing and layout efficiency  
  - Total number of modules placed  

- **Outputs:**
  - Timestamped CSV file with complete bill of materials  
  - CSV is formatted for direct import into SAGE Intacct or manual editing  

## 🧠 Technologies Used  
- Python 3  
- `pandas` – for structured data handling and CSV creation  
- `math` – for all structural and dimensional logic  
- `datetime` & `os` – for file management and timestamping  

## 📈 Impact  
- Reduced manual quoting time from 30–60 minutes to under 15  
- Eliminated spreadsheet errors and manual calculations  
- Enabled faster quote generation and higher accuracy  
- Provided consistent quoting outputs usable by technical and sales staff  

## 📷 Screenshots  
Add images here to demonstrate how the tool works and what the output looks like.

Suggested examples:
- Screenshot 1 – Command-line input process  
- Screenshot 2 – Output CSV file opened in Excel  
- Screenshot 3 – Component breakdown as seen in the final quote

**To include images in this README on GitHub:**  
1. Place screenshots in a folder such as `/images/` inside your repository  
2. Embed them like this:  
   `![Screenshot Description](images/your_image_name.png)`

## 🔒 Note  
This public version has been sanitized from the original tool developed internally at **Lumax Energy**. All pricing data, client-specific configurations, and proprietary files have been removed or replaced with placeholder content for demonstration purposes.
