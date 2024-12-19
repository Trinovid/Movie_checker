# **Movie_Checker**

A Python script that scans video files in a specified directory, extracts detailed metadata, and generates an Excel document for easy analysis. This tool helps you filter and manage video files by criteria such as bitrate, resolution, and audio quality.

## **Features**  
- **Automatic Metadata Extraction**: Extracts detailed video and audio information from supported formats.  
- **Excel Output**: Generates an Excel file (`movie_list.xlsx`) with metadata for easy filtering and analysis.  
- **Support for Multiple Formats**: Handles common video file formats like `.mp4`, `.mkv`, `.avi`, and `.mov`.

## **Getting Started**

### **Prerequisites**  
- Python 3.8 or later  
- Required Python libraries: `pandas`, `openpyxl`, `xlsxwriter`, `pymediainfo`  

Install dependencies via pip:  
```bash
pip install pandas openpyxl xlsxwriter pymediainfo
