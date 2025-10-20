# نظام البحث في قواعد بيانات الناخبين العراقيين
# Iraqi Voter Database Search System

## 📋 System Overview
نظام بحث متطور لقواعد بيانات الناخبين العراقيين يدعم 19 محافظة عراقية مع إمكانية البحث السريع والدقيق.

A professional search system for Iraqi voter databases supporting 19 Iraqi provinces with fast and accurate search capabilities.

---

## 🚀 Quick Start

### Method 1: PowerShell Script (Recommended)
1. Right-click on `setup_and_run.ps1`
2. Select "Run with PowerShell"
3. Follow the on-screen instructions

### Method 2: Batch File (Simple)
1. Double-click on `START.bat`
2. Wait for the server to start
3. Open browser and go to: http://127.0.0.1:5000

### Method 3: Manual Start
```powershell
# Install requirements
pip install Flask pyodbc

# Run the application
python app.py
```

---

## 📦 System Requirements

### Required Software:
- **Python 3.8 or higher** 
  - Download: https://www.python.org/downloads/
  - ✅ Make sure to check "Add Python to PATH" during installation

- **Microsoft Access Database Engine 2016 Redistributable**
  - Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920
  - Required to connect to .accdb database files
  - Choose the version matching your Office installation (32-bit or 64-bit)

### Required Python Packages:
- Flask 3.1.2 or higher
- pyodbc 5.3.0 or higher

*(Automatically installed by the setup scripts)*

---

## 📁 Required Files

### Application Files:
```
بانات الناخbين2025/
├── app.py                          # Main Flask application
├── templates/
│   ├── index.html                  # Search page
│   └── results.html                # Results page
├── setup_and_run.ps1              # PowerShell setup script
├── START.bat                       # Simple batch startup
└── README.md                       # This file
```

### Database Files (Required):
```
├── البصرة- الناصرية - الموصل.accdb
├── بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb
├── بغداد كرخ - رصافة.accdb
├── دهوك- اربيل-السليمانية-كركوك.accdb
└── كربلاء-ميسان- المثنى- النجف- واسط.accdb
```

**⚠️ Important:** All 5 database files must be in the same directory as `app.py`

---

## 🌐 Accessing the Application

### Local Access:
- http://127.0.0.1:5000
- http://localhost:5000

### Network Access (from other computers):
1. Find your IP address:
   ```powershell
   ipconfig
   ```
   Look for "IPv4 Address" (e.g., 192.168.1.100)

2. Access from other computers on the same network:
   ```
   http://YOUR_IP_ADDRESS:5000
   ```
   Example: http://192.168.1.100:5000

3. **Firewall Note:** You may need to allow port 5000 through Windows Firewall

---

## 🔧 Supported Provinces (19 محافظة)

### Database 1: البصرة- الناصرية - الموصل
- البصرة (Basra)
- ذي قار / الناصرية (Dhi Qar / Nasiriyah)
- نينوى / الموصل (Nineveh / Mosul)

### Database 2: بابل- ديالى- الديوانية-صلاح الدين-الانبار
- الأنبار (Anbar)
- بابل (Babil)
- ديالى (Diyala)
- القادسية / الديوانية (Qadisiyah / Diwaniyah)
- صلاح الدين (Saladin)

### Database 3: بغداد كرخ - رصافة
- بغداد الكرخ (Baghdad Al-Karkh)
- بغداد الرصافة (Baghdad Al-Rusafa)

### Database 4: دهوك- اربيل-السليمانية-كركوك
- دهوك (Duhok)
- أربيل (Erbil)
- كركوك (Kirkuk)
- السليمانية (Sulaymaniyah)

### Database 5: كربلاء-ميسان- المثنى- النجف- واسط
- كربلاء (Karbala)
- ميسان (Maysan)
- المثنى (Muthanna)
- النجف (Najaf)
- واسط (Wasit)

---

## 🔍 Search Features

### Search Fields:
- 🔢 Personal ID Number (الرقم الشخصي)
- 👨‍👩‍👧‍👦 Family Number (رقم العائلة)
- 🔢 Sub Number (الرقم الفرعي)
- 👤 First Name (الاسم الأول)
- 👤 Father's Name (اسم الأب)
- 👤 Grandfather's Name (اسم الجد)
- 👤 Fourth Name (الاسم الرابع)
- 👩 Mother's Name (اسم الأم)
- ⚧️ Gender (الجنس)
- 📅 Date of Birth (تاريخ الميلاد)
- 🌍 Birth Country (بلد الولادة)
- 📱 Phone Number (رقم الهاتف)
- 🏛️ Voting Center (مركز الاقتراع)
- 🗳️ Polling Station (محطة الاقتراع)
- 🏠 Address (العنوان)
- 🇮🇶 Nationality (الجنسية)

### Search Modes:
1. **🎯 Exact Search (بحث دقيق)**
   - Perfect match
   - 10x faster
   - Best for ID numbers and complete data

2. **🔍 Partial Search (بحث جزئي)**
   - Contains match
   - More comprehensive
   - Best for names and partial information

### Performance:
- ⚡ Exact search: 2-6 seconds
- 🔍 Partial search: 5-10 seconds
- 💾 Cached searches: < 0.1 seconds
- 📊 Result limit: 1000 records per search

---

## ⚙️ Advanced Configuration

### Change Port Number:
Edit `app.py` line 309:
```python
app.run(debug=True, host='0.0.0.0', port=5000)  # Change 5000 to your desired port
```

### Enable Network Access:
The default configuration (`host='0.0.0.0'`) already allows network access.
To restrict to local only, change to:
```python
app.run(debug=True, host='127.0.0.1', port=5000)
```

### Adjust Result Limit:
Edit `app.py` line 95:
```python
max_results = 1000  # Change to your desired limit
```

### Adjust Cache Size:
Edit `app.py` line 11:
```python
cache_max_size = 50  # Change to keep more/fewer searches in cache
```

---

## 🐛 Troubleshooting

### Problem: "Python is not recognized"
**Solution:** 
1. Reinstall Python from https://www.python.org/downloads/
2. Make sure to check "Add Python to PATH" during installation
3. Restart your computer

### Problem: "Database connection error"
**Solution:**
1. Install Microsoft Access Database Engine 2016
2. Make sure all .accdb files are in the correct directory
3. Check file permissions

### Problem: "Module 'pyodbc' not found"
**Solution:**
```powershell
pip install pyodbc
```

### Problem: "Port 5000 already in use"
**Solution:**
1. Close other applications using port 5000
2. Or change the port in app.py (see Advanced Configuration)

### Problem: Can't access from other computers
**Solution:**
1. Check Windows Firewall settings
2. Make sure both computers are on the same network
3. Verify you're using the correct IP address

---

## 📊 Features

✅ Search 19 Iraqi provinces  
✅ 16 search criteria fields  
✅ Exact and partial search modes  
✅ Professional Arabic RTL interface  
✅ Export results to CSV  
✅ Print-friendly results page  
✅ Connection pooling for performance  
✅ Result caching (last 50 searches)  
✅ Schema caching for faster queries  
✅ Responsive design for mobile devices  
✅ Modern gradient UI with animations  
✅ Real-time search validation  

---

## 🔒 Security Notes

- ⚠️ This application is designed for local/internal network use
- 🔐 Do not expose to the public internet without proper authentication
- 💾 Database files contain sensitive personal information
- 🛡️ Keep the system updated and secure

---

## 📞 Support

For issues or questions:
1. Check the Troubleshooting section above
2. Review the console output for error messages
3. Ensure all requirements are properly installed

---

## 📝 Version Information

- **Version:** 2.0
- **Last Updated:** October 2025
- **Python Version:** 3.12+
- **Flask Version:** 3.1.2
- **Database Format:** Microsoft Access (.accdb)

---

## 📄 License

© 2025 Iraqi Voter Database Search System  
For authorized use only.

---

**🇮🇶 صنع بفخر للعراق | Made with pride for Iraq 🇮🇶**
