# 🌍 Excel to KML Geolocation Tool

Welcome to the **Excel to KML Geolocation Tool**! This tool is designed to help you effortlessly convert geospatial data into KML format, making it simple to visualize on maps such as Google Earth.

Built using Visual Basic, this Excel sheet allows you to input locations with custom placemark icons, and outputs a KML file that can be used for map rendering.

## 🚀 Features

- **Simple Data Input**:  
  Enter your geolocation data, including names, coordinates, and descriptions, directly into the Excel sheet.
  
- **Placemark Icons**:  
  Choose from a wide variety of placemark icons listed in the catalog tab, or provide your own icon URL.

- **KML File Generation**:  
  Automatically generates a KML file based on your input, ready to use in any compatible mapping software.

## 📄 How to Use

1. **Enter Your Data**:  
   In the `Data` sheet, input the following information:
   
   | Name          | Longitude  | Latitude   | Description       | Placemark Icon URL                |
   |---------------|------------|------------|-------------------|-----------------------------------|
   | Example Point | 30.2544886 | 50.1234567 | Example location  | http://example.com/icon.png       |
   
   - **Longitude & Latitude** must be in decimal format (e.g., `30.2544886`).
   - If you leave the "Placemark Icon URL" blank, a default icon will be used.

2. **Browse Icon Catalog**:  
   The `Placemark Icon Catalog` tab contains a variety of pre-defined icons to choose from. Copy and paste the icon URL into your data input for a customized map experience.

3. **Generate the KML File**:  
   Once your data is entered, the tool will automatically generate a KML file saved at `C:/myapp.kml`.

4. **Visualize the KML File**:  
   Open the `myapp.kml` file in Google Earth or any KML-compatible mapping software to see your placemarks visualized on the map.

## 📂 Folder Structure

- **Data Sheet**:  
  Input your geolocation data here, including names, coordinates, and optional icon URLs.
  
- **Placemark Icon Catalog**:  
  A reference sheet with various placemark icons and URLs to make your map visually appealing.

## ⚙️ Technology Stack

- **Excel VBA (Visual Basic for Applications)**:  
  All the magic happens behind the scenes with VBA, converting your data into a fully functional KML file.

## 📢 Feedback

We'd love to hear your feedback and suggestions to improve this tool! Feel free to create issues or submit pull requests. 

---

Start mapping your world today! 🌍📍

---
