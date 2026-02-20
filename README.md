<h1 align="center">⚡ FTTH Automation for EMR Project</h1>

<p align="center">
  Aplikasi desktop all-in-one untuk otomasi perencanaan jaringan FTTH pada proyek EMR.<br/>
  Dibangun dengan <strong>Python + Tkinter</strong> — ringan, portabel, tanpa instalasi rumit.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.9+-blue?logo=python&logoColor=white"/>
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows"/>
  <img src="https://img.shields.io/badge/Status-Active-brightgreen"/>
  <img src="https://img.shields.io/github/license/galprim1412/ftth-automation"/>
</p>

---

## 🗂️ Fitur

| Tab | Fungsi |
|-----|--------|
| 📎 **Cable Name Generator** | Generate penamaan kabel berdasarkan parameter rute & segmen |
| 🏘️ **Cluster Description** | Generate string deskripsi cluster secara otomatis |
| 🔌 **Feeder Description** | Generate string deskripsi feeder secara otomatis |
| 📍 **HP Grouping by FAT for KMZ** | Kelompokkan titik homepass ke dalam polygon FAT dari file KML |
| 🗺️ **CSV → KML Converter** | Konversi data koordinat CSV menjadi placemark KML |
| 🔢 **Homepass Counter** | Hitung jumlah homepass dari file KML |
| 📤 **KML Extractor for HPDB** | Ekstrak data placemark KML ke Excel untuk HPDB |
| 📊 **BoQ Generator for FDDP** | Proses file BoQ Excel & hasilkan template upload material/service FDDP |

---

## 🔧 Persyaratan

- Python **3.9** atau lebih baru
- Library:
  ```
  pandas
  openpyxl
  ```

Install dependensi:
```bash
pip install pandas openpyxl
```

---

## 🚀 Cara Menjalankan

```bash
python ftthautomation.py
```

> Tidak perlu konfigurasi tambahan. Semua logika sudah tertanam dalam satu file.

---

## 📁 Struktur Project

```
ftth-automation/
├── ftthautomation.py   # Aplikasi utama (single-file, self-contained)
├── .gitignore
└── README.md
```

---

## 👤 Author

**Galih Prima** — [@galprim1412](https://github.com/galprim1412)

---

<p align="center">Dibuat untuk kebutuhan operasional EMR Project 🇮🇩</p>
