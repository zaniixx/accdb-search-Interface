# Database Search System

## 📋 System Overview

A professional search system for supporting multiple databases with fast and accurate search capabilities, modern responsive UI, and advanced features.

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
pip install Flask pyodbc pandas openpyxl

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
- pandas 2.0.0 or higher
- openpyxl 3.1.0 or higher

*(Automatically installed by the setup scripts)*

---

## 📁 Project Structure

```
├── app.py                          # Main Flask application
├── .gitignore                      # Git ignore rules
├── README.md                       # This documentation
├── setup_and_run.ps1              # PowerShell setup script
├── START.bat                       # Simple batch startup
├── template.xlsx                   # Excel search template
├── databases/                      # Database files directory
│   ├── file.accdb
│   ├── .
│   ├── .
│   ├── .
├── templates/
│   ├── index.html                  # Modern search interface
│   └── results.html                # Results display page
├── uploads/                        # Uploaded Excel files directory
└── __pycache__/                    # Python cache (ignored)
```

**⚠️ Important:** All database files must be placed in the `databases/` directory. The application automatically discovers and loads all `.accdb` files from this folder on startup, making it secure for code sharing without exposing database names.

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

## 🔍 Search Features

### Search Types:
1. **🔍 Single Search (البحث الفردي)**
   - Search using individual criteria
   - Support for all 16 search fields

2. **👥 Multi Search (البحث المتعدد)**
   - Search multiple values at once
   - Comma-separated or line-separated input
   - Perfect for bulk ID lookups

3. **📊 Excel Upload (رفع ملف Excel)**
   - Upload Excel files with multiple search criteria
   - Automatic column mapping
   - Download professional search template

### Search Fields (16 معيار):
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
- Nationality (الجنسية)

### Search Modes:
1. **🎯 Exact Search (بحث دقيق)**
   - Perfect match
   - 10x faster performance
   - Best for ID numbers and complete data

2. **🔍 Partial Search (بحث جزئي)**
   - Contains match
   - More comprehensive results
   - Best for names and partial information

### Performance:
- ⚡ Exact search: 2-6 seconds
- 🔍 Partial search: 5-10 seconds
- 💾 Cached searches: < 0.1 seconds
- 📊 Result limit: 1000 records per search
- 🔄 Parallel processing for multiple databases
- 🗄️ Connection pooling for optimal performance

---

## 🎨 Modern User Interface

### Design Features:
- **📱 Responsive Design** - Works perfectly on desktop, tablet, and mobile
- **🌙 Clean Modern UI** - Professional card-based layout with subtle shadows
- **🎭 Smooth Animations** - Elegant transitions and hover effects
- **🎨 Professional Typography** - Inter font for optimal readability
- **🔤 RTL Support** - Perfect Arabic text rendering and layout
- **🎯 Intuitive Navigation** - Tabbed interface for different search types
- **📤 Drag & Drop Upload** - Easy Excel file upload with visual feedback
- **⚡ Real-time Validation** - Instant form validation and error messages

### Accessibility:
- **♿ Screen Reader Friendly** - Proper ARIA labels and semantic HTML
- **⌨️ Keyboard Navigation** - Full keyboard accessibility
- **🔍 High Contrast** - Clear visual hierarchy and readable text
- **🌐 Multi-language Ready** - Arabic and English support

---

## 📊 Advanced Features

✅ Search across multiple dbs simultaneously  
✅ 16 comprehensive search criteria fields  
✅ Three search modes: Single, Multi, and Excel upload  
✅ Professional Arabic RTL interface with modern design  
✅ Excel template download with professional formatting  
✅ Export results to CSV format  
✅ Print-friendly results page with responsive layout  
✅ Connection pooling and database optimization  
✅ Intelligent result caching (last 50 searches)  
✅ Schema caching for faster subsequent queries  
✅ Real-time search progress indicators  
✅ Advanced form validation and error handling  
✅ Mobile-responsive design for all devices  
✅ Modern gradient UI with smooth animations  
✅ File upload with drag-and-drop support  
✅ Automatic column mapping for Excel uploads  
✅ Network access support for multi-user environments  
✅ **Dynamic database discovery** - Automatically finds and loads all database files  

---

## ⚙️ Advanced Configuration

### Change Port Number:
Edit `app.py` line 309:
```python
app.run(debug=True, host='0.0.0.0', port=5000)  # Change 5000 to your desired port
```

### Enable Local Access Only:
To restrict to local access only, change in `app.py`:
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

### Database Path Configuration:
The application automatically looks for databases in the `databases/` folder. To change this, edit the `db_files` list in `app.py`.

---

## 🐛 Troubleshooting

### Problem: "Python is not recognized"
**Solution:**
1. Reinstall Python from https://www.python.org/downloads/
2. Make sure to check "Add Python to PATH" during installation
3. Restart your computer

### Problem: "Database connection error"
**Solution:**
1. Install Microsoft Access Database Engine 2016 Redistributable
2. Ensure all .accdb files are in the `databases/` directory
3. Check file permissions and antivirus exclusions

### Problem: "Module not found" errors
**Solution:**
```powershell
pip install Flask pyodbc pandas openpyxl
```

### Problem: "Port 5000 already in use"
**Solution:**
1. Close other applications using port 5000
2. Or change the port in app.py (see Advanced Configuration)

### Problem: Can't access from other computers
**Solution:**
1. Check Windows Firewall settings for port 5000
2. Ensure both computers are on the same network
3. Verify you're using the correct IP address

### Problem: Excel upload fails
**Solution:**
1. Download the template from the application
2. Ensure column headers match exactly
3. Check that the file is not corrupted
4. Verify Microsoft Excel is properly installed

---

## 🔒 Security & Best Practices

### Database Security Features:
- **🔍 Dynamic Database Discovery** - Automatically discovers and maps all database files at startup
- **🚫 No Hardcoded Filenames** - Database names are never stored in source code for secure publishing
- **📁 Secure File Organization** - All databases stored in dedicated `databases/` folder
- **🛡️ Git Ignore Protection** - Database files automatically excluded from version control
- **🔄 Runtime Mapping** - Table-to-database relationships built dynamically by inspecting actual files

### Security Best Practices:
- ⚠️ **Local/Internal Use Only** - Designed for local network environments
- 🔐 **No Public Internet Exposure** - Do not expose without proper authentication
- 💾 **Sensitive Data Handling** - Database files contain personal information
- 🛡️ **Keep Updated** - Regularly update dependencies and security patches
- 🔒 **Access Control** - Implement proper access controls for production use
- � **Audit Logging** - All database connections and searches are logged for security monitoring

### Code Publishing Security:
The application is designed to be safely published on platforms like GitHub without exposing sensitive database information. Database filenames are discovered at runtime, making the code repository secure while maintaining full functionality when database files are present.

---

## 📞 Support & Contributing

### Getting Help:
1. Check the Troubleshooting section above
2. Review console output for detailed error messages
3. Ensure all system requirements are properly installed
4. Check file paths and permissions

---

## 📄 License & Credits

© 2025 Database Search System  
**Made with ❤️ by [Zanix](https://github.com/zaniixx)**

