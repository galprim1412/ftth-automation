# FTTH Automation for EMR Project

A desktop application built with Python + Tkinter for automating FTTH (Fiber to the Home) network planning tasks in EMR projects.

## Features

| Tab | Description |
|-----|-------------|
| **Cable Name Generator** | Generate cable naming based on route/segment parameters |
| **Cluster Description** | Generate cluster description strings |
| **Feeder Description** | Generate feeder description strings |
| **HP Grouping by FAT for KMZ** | Group homepass points into FAT polygons from KML files |
| **CSV → KML Converter** | Convert CSV coordinate data into KML placemarks |
| **Homepass Counter** | Count homepass from KML files |
| **KML Extractor for HPDB** | Extract KML placemark data to Excel |
| **BoQ Generator for FDDP** | Process BoQ Excel files and generate material/service upload templates |

## Requirements

```
Python 3.9+
pandas
openpyxl
```

Install dependencies:
```bash
pip install pandas openpyxl
```

## Usage

```bash
python ftthautomation.py
```

## Project Structure

```
ftthautomation.py   # Main application (single-file, self-contained)
```

## License

Internal use — EMR Project.
