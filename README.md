<h1 align="center">⚡ FTTH Automation for EMR Project</h1>

<p align="center">
  Aplikasi desktop all-in-one untuk otomasi perencanaan jaringan FTTH pada proyek EMR.<br/>
  Dibangun dengan <strong>Python + Tkinter</strong> — ringan, portabel, tanpa instalasi rumit.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.9+-blue?logo=python&logoColor=white"/>
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-informational"/>
  <img src="https://img.shields.io/badge/Status-Active-brightgreen"/>
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

## 🔧 Instalasi

### Prasyarat
- Python **3.9** atau lebih baru
- Semua dependensi ada di `requirements.txt`

```bash
pip install -r requirements.txt
```

### Per Platform

<details>
<summary>🪟 Windows</summary>

```bash
# Tkinter sudah termasuk dalam instalasi Python standar
pip install -r requirements.txt
python ftthautomation.py
```
</details>

<details>
<summary>🍎 macOS</summary>

```bash
# Jika tkinter belum ada, install via Homebrew
brew install python-tk

pip install -r requirements.txt
python3 ftthautomation.py
```
</details>

<details>
<summary>🐧 Linux (Ubuntu/Debian)</summary>

```bash
# Install tkinter jika belum ada
sudo apt install python3-tk

pip install -r requirements.txt
python3 ftthautomation.py
```
</details>

---

## 🚀 Cara Menjalankan

```bash
python ftthautomation.py
```

> Tidak perlu konfigurasi tambahan. Semua logika tertanam dalam satu file. Font UI otomatis menyesuaikan platform (Segoe UI di Windows, Helvetica Neue di macOS, DejaVu Sans di Linux).

---

## 📁 Struktur Project

```
ftth-automation/
├── ftthautomation.py    # Aplikasi utama (single-file, self-contained)
├── requirements.txt     # Dependensi Python
├── .gitignore
└── README.md
```

---

## 👤 Author

**Galih Prima** — [@galprim1412](https://github.com/galprim1412)

---

<p align="center">Dibuat untuk kebutuhan operasional EMR Project 🇮🇩</p>
