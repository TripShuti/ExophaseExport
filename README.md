# ExophaseExport TUI
<img width="1920" height="1080" alt="2026-05-16_05-23-48" src="https://github.com/user-attachments/assets/d0b22304-6c4b-4d07-9083-a261579ad615" />

A terminal UI application to grab, view, and export all your game data and statistics from exophase.com.

## 🔍 How to get your Player ID
To sync your data, you need your `playerProfileId` from Exophase:
1. Go to your Exophase profile in your browser.
2. Press `F12` to open Developer Tools.
3. Search in the source code for `window.playerProfileId =`
4. Copy the number and paste it into the app.

<img width="286" alt="image" src="https://github.com/user-attachments/assets/be3cc10a-bfaa-42a8-b4a8-b2d0184357cc" />

## 🚀 Installation

Recommended use a virtual environment.

### 1. Create and activate venv
**Linux / macOS:**
```bash
python3 -m venv venv
source venv/bin/activate
```

**Windows:**
```cmd
python -m venv venv
venv\Scripts\activate
```

### 2. Install dependencies
```bash
pip install textual requests pandas openpyxl
```

## 🎮 Usage
Run the application:
```bash
python ExophaseExport.py
```
