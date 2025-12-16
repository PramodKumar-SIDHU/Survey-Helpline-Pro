# 🛠 Survey Helpline Pro

![QGIS 3+](https://img.shields.io/badge/QGIS-3.x-blue?style=flat-square) ![License GPL](https://img.shields.io/badge/License-GPL--2.0-orange?style=flat-square)

---

## 📌 Overview

**Survey Helpline Pro** is a unified QGIS plugin designed to **streamline and automate spatial data processing** for survey teams.  
It provides workflows for **Survey & Land Records (SS&LR)** and **Panchayat Raj (SVAMITVA)** departments, ensuring **consistent, error-free, and standardized outputs** for shapefile processing, parcel management, and map generation.

---

## 🎯 Key Features

- 🗂️ **Automated Shapefile Processing** – Clean, validate, merge, and fix geometries.  
- 🏡 **Parcel & PPM Management** – Generate standardized parcel maps.  
- 📊 **Excel Attribute Integration** – Upload attribute tables for PPMs with 9(2) notices.  
- 🖼️ **Village Map Generation** – Automatically generate maps for villages.  
- ⚡ **Time Saving** – Reduces repetitive GIS tasks and manual errors.  

---

## 📂 Installation

1. Download the ZIP from [GitHub Releases](https://github.com/PramodKumar-SIDHU/Survey-Helpline-Pro/releases).  
2. Open **QGIS → Plugins → Manage and Install Plugins → Install from ZIP**.  
3. Enable **Survey Helpline Pro** from the Plugin Manager.  

---

## 📝 Usage

1. Open **Survey Helpline Pro** from the QGIS plugin menu.  
2. Fill in **parameters**: District, Mandal, Revenue Village, Shapefiles, and Excel attributes.  
3. Check options like **Initial PPMs** or **PPMs with 9(2) Notice**.  
4. Click **Generate Maps**.  
5. Maps and processed shapefiles are saved in the selected output folder.  

---

## 📊 Excel Requirements

- **Mandatory columns**:  
  - `PPN`  
  - `Property type (Individual/Joint/Apartment/Government)`  
  - `Panchayat Name`  
  - `Owner Name`  
  - `Relation (W/O,H/O,S/O,D/O)`  
  - `Assessment No.`  
  - `Remarks`  

- No merged cells  
- First row must contain column headers  
- Headers must match exactly (case-sensitive, no extra spaces)  

Download [Sample Excel](link-to-your-sample-excel) for reference.

---

## 🛠 Development

- QGIS 3.x Python Plugin  
- License: **GNU GPL 2.0 or later**  
- Main files: `survey_helpline_pro.py`, `dialogs/`, `resources.py`, `metadata.txt`  
- Icon: `icon.png`  

---

## 📍 Repository & Support

- **GitHub**: [https://github.com/PramodKumar-SIDHU/Survey-Helpline-Pro](https://github.com/PramodKumar-SIDHU/Survey-Helpline-Pro)  
- **Issues / Tracker**: [GitHub Issues](https://github.com/PramodKumar-SIDHU/Survey-Helpline-Pro/issues)  

---

## 🌟 Acknowledgements

- Developed for **Government of Andhra Pradesh**  
- Supports **SVAMITVA / SS&LR initiatives**  
- Follows **QGIS plugin development guidelines**  


