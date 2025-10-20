from flask import Flask, render_template, request, send_file
import pyodbc
import os
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configuration for file uploads
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create upload folder if it doesn't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Connection cache for reusing database connections
connection_cache = {}

# Results cache to avoid redundant searches
results_cache = {}
cache_max_size = 50  # Keep last 50 searches

# Table schema cache to avoid repeated column queries
table_schema_cache = {}

# List of database files
db_files = [
    'البصرة- الناصرية - الموصل.accdb',
    'بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb',
    'بغداد كرخ - رصافة.accdb',
    'دهوك- اربيل-السليمانية-كركوك.accdb',
    'كربلاء-ميسان- المثنى- النجف- واسط.accdb'
]

# Mapping of tables to databases and their display names
table_to_db = {
    'dbo_Basra': {'db': 'البصرة- الناصرية - الموصل.accdb', 'name': 'البصرة'},
    'dbo_Diqar': {'db': 'البصرة- الناصرية - الموصل.accdb', 'name': 'ذي قار'},
    'dbo_Ninawa': {'db': 'البصرة- الناصرية - الموصل.accdb', 'name': 'نينوى'},
    'dbo_Anbar': {'db': 'بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb', 'name': 'الأنبار'},
    'dbo_Babil': {'db': 'بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb', 'name': 'بابل'},
    'dbo_diyala': {'db': 'بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb', 'name': 'ديالى'},
    'dbo_Qadisyah': {'db': 'بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb', 'name': 'القادسية'},
    'dbo_SlahAldin': {'db': 'بابل- ديالى- الديوانية-صلاح الدين-الانبار.accdb', 'name': 'صلاح الدين'},
    'dbo_Karikh': {'db': 'بغداد كرخ - رصافة.accdb', 'name': 'بغداد الكرخ'},
    'dbo_Rusafa': {'db': 'بغداد كرخ - رصافة.accdb', 'name': 'بغداد الرصافة'},
    'dbo_Dahuk': {'db': 'دهوك- اربيل-السليمانية-كركوك.accdb', 'name': 'دهوك'},
    'dbo_Erbil': {'db': 'دهوك- اربيل-السليمانية-كركوك.accdb', 'name': 'أربيل'},
    'dbo_Kirkuk': {'db': 'دهوك- اربيل-السليمانية-كركوك.accdb', 'name': 'كركوك'},
    'dbo_sulimaniah': {'db': 'دهوك- اربيل-السليمانية-كركوك.accdb', 'name': 'السليمانية'},
    'dbo_Kirbla': {'db': 'كربلاء-ميسان- المثنى- النجف- واسط.accdb', 'name': 'كربلاء'},
    'dbo_Misan': {'db': 'كربلاء-ميسان- المثنى- النجف- واسط.accdb', 'name': 'ميسان'},
    'dbo_Muthana': {'db': 'كربلاء-ميسان- المثنى- النجف- واسط.accdb', 'name': 'المثنى'},
    'dbo_Najaf': {'db': 'كربلاء-ميسان- المثنى- النجف- واسط.accdb', 'name': 'النجف'},
    'dbo_Wasit': {'db': 'كربلاء-ميسان- المثنى- النجف- واسط.accdb', 'name': 'واسط'}
}

def get_connection(db_path):
    """Create or reuse a connection to an Access database."""
    # Check if we have a cached connection
    if db_path in connection_cache:
        try:
            # Test if connection is still alive
            conn = connection_cache[db_path]
            conn.cursor().execute("SELECT 1")
            print(f"  Reusing cached connection for {db_path}")
            return conn
        except:
            # Connection is dead, remove from cache
            print(f"  Cached connection dead, creating new one")
            del connection_cache[db_path]
    
    # Create new connection
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={db_path};'
    )
    conn = pyodbc.connect(conn_str)
    connection_cache[db_path] = conn
    print(f"  Created new connection for {db_path}")
    return conn

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_multi_values(value_string):
    """Parse comma-separated or newline-separated values."""
    if not value_string or not value_string.strip():
        return []
    
    # Split by commas and newlines, then clean up
    values = []
    for line in value_string.split('\n'):
        for val in line.split(','):
            val = val.strip()
            if val:
                values.append(val)
    return list(set(values))  # Remove duplicates

def process_excel_file(file_path):
    """Process Excel file and return list of search criteria."""
    try:
        df = pd.read_excel(file_path)
        
        # Map Arabic column names to English field names
        arabic_to_english = {
            'الرقم الشخصي': 'perid',
            'رقم العائلة': 'perfamno',
            'الرقم الفرعي': 'persubno',
            'الاسم الأول': 'perfirst',
            'اسم الأب': 'perfather',
            'اسم الجد': 'pergrand',
            'الاسم الرابع': 'perfourth_name',
            'اسم الأم': 'mothername',
            'الجنس': 'pergender',
            'تاريخ الميلاد': 'perdob',
            'بلد الولادة': 'perbthcountry',
            'رقم الهاتف': 'pertelno',
            'مركز الاقتراع': 'vrcname',
            'محطة الاقتراع': 'pcname',
            'العنوان': 'peraddress',
            'الجنسية': 'pernationality'
        }
        
        criteria_list = []
        
        for _, row in df.iterrows():
            criteria = {}
            for arabic_col, english_col in arabic_to_english.items():
                if arabic_col in df.columns:
                    value = row[arabic_col]
                    if pd.notna(value):  # Check if not NaN
                        # Convert numeric values back to proper format
                        if english_col in ['perid', 'perfamno', 'persubno', 'pertelno']:
                            try:
                                # Remove .0 from numeric strings
                                value = str(int(float(value)))
                            except (ValueError, TypeError):
                                value = str(value).strip()
                        elif english_col == 'pergender':
                            try:
                                value = str(int(float(value)))
                            except (ValueError, TypeError):
                                value = str(value).strip()
                        else:
                            value = str(value).strip()
                        
                        criteria[english_col] = value
            
            if criteria:  # Only add if has at least one value
                criteria_list.append(criteria)
        
        return criteria_list
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return []

def search_databases_optimized(criteria, selected_table=None, search_mode='partial'):
    """Optimized search using UNION ALL and type-aware queries.

    Args:
        criteria: Dictionary of search criteria
        selected_table: Specific table to search (optional)
        search_mode: 'exact' for = comparison or 'partial' for LIKE
    """
    # Create cache key
    import hashlib
    import json
    cache_key = hashlib.md5(
        json.dumps({'criteria': criteria, 'table': selected_table, 'mode': search_mode}, sort_keys=True).encode()
    ).hexdigest()

    # Check cache
    if cache_key in results_cache:
        print(f"⚡ CACHE HIT! Returning cached results")
        return results_cache[cache_key]

    results = []
    max_results = 1000

    # Determine which tables to search
    if selected_table and selected_table in table_to_db:
        tables_to_search = [selected_table]
        db_file = table_to_db[selected_table]['db']
        db_files_to_search = [db_file]
    else:
        tables_to_search = list(table_to_db.keys())
        db_files_to_search = list(set(table_to_db[t]['db'] for t in tables_to_search))

    # Group tables by database for efficient querying
    tables_by_db = {}
    for table in tables_to_search:
        db = table_to_db[table]['db']
        if db not in tables_by_db:
            tables_by_db[db] = []
        tables_by_db[db].append(table)

    for db_file in db_files_to_search:
        if len(results) >= max_results:
            break

        db_path = os.path.join(os.getcwd(), db_file)
        if not os.path.exists(db_path):
            continue

        try:
            conn = get_connection(db_path)
            cursor = conn.cursor()

            tables_in_db = tables_by_db[db_file]

            # Build optimized UNION ALL query
            union_queries = []
            all_params = []

            for table_name in tables_in_db:
                if len(results) >= max_results:
                    break

                # Get cached schema
                schema_key = f"{db_file}:{table_name}"
                if schema_key not in table_schema_cache:
                    try:
                        # Load column names from database
                        cursor.columns(table=table_name)
                        columns_info = cursor.fetchall()
                        columns = [row.column_name for row in columns_info if row.column_name]
                        table_schema_cache[schema_key] = columns
                        print(f"  Loaded schema for {table_name}: {len(columns)} columns - {columns}")
                    except Exception as e:
                        print(f"  Failed to load schema for {table_name}: {e}")
                        continue
                columns = table_schema_cache[schema_key]

                if not columns:
                    continue

                # Build WHERE clause with type awareness
                where_parts = []
                params = []

                for key, value in criteria.items():
                    if not value or not value.strip():
                        continue

                    matching_col = next((col for col in columns if col.upper() == key.upper()), None)
                    if not matching_col:
                        continue

                    # Type-aware search logic
                    if search_mode == 'exact':
                        # For exact matches, use type-appropriate comparison
                        if key.upper() in ['PERID', 'PERFAMNO', 'PERSUBNO']:
                            # Numeric fields - direct numeric comparison
                            where_parts.append(f"([{matching_col}] IS NOT NULL AND [{matching_col}] = ?)")
                            try:
                                params.append(int(value.strip()))
                            except ValueError:
                                params.append(value.strip())  # fallback to string if not numeric
                        else:
                            # Text fields - string comparison
                            where_parts.append(f"([{matching_col}] IS NOT NULL AND LTrim(RTrim(CStr([{matching_col}]))) = ?)")
                            params.append(value.strip())
                    else:
                        # Partial search - optimize LIKE patterns
                        if key.upper() in ['PERID', 'PERFAMNO', 'PERSUBNO']:
                            # For numeric fields in partial mode, still use exact match
                            where_parts.append(f"([{matching_col}] IS NOT NULL AND CStr([{matching_col}]) LIKE ?)")
                            params.append(f'%{value.strip()}%')
                        else:
                            # For text fields, use optimized LIKE
                            where_parts.append(f"([{matching_col}] IS NOT NULL AND CStr([{matching_col}]) LIKE ?)")
                            params.append(f'%{value.strip()}%')

                if not where_parts:
                    continue

                # Add table identifier and build query
                where_clause = " AND ".join(where_parts)
                remaining = max_results - len(results)

                # Add a column to identify which table the result came from
                select_cols = ", ".join([f"[{col}]" for col in columns])
                query = f"SELECT TOP {remaining} '{table_name}' AS source_table, {select_cols} FROM [{table_name}] WHERE {where_clause}"

                union_queries.append(f"({query})")
                all_params.extend(params)

            if not union_queries:
                continue

            # Execute UNION ALL query
            final_query = " UNION ALL ".join(union_queries)
            print(f"Executing optimized UNION query with {len(union_queries)} table(s)")
            print(f"Query: {final_query}")
            print(f"Params: {all_params}")

            cursor.execute(final_query, all_params)
            rows = cursor.fetchall()

            # Process results
            for row in rows:
                if len(results) >= max_results:
                    break

                # First column is source_table, rest are data
                source_table = row[0]
                row_data = dict(zip(columns, row[1:]))

                results.append({
                    'db': db_file,
                    'table': source_table,
                    'data': row_data
                })

            print(f"Found {len(rows)} results in {db_file}")

        except Exception as e:
            print(f"ERROR in {db_file}: {str(e)[:200]}")
            continue

    # Cache and return results
    results_cache[cache_key] = results
    if len(results_cache) > cache_max_size:
        oldest_key = next(iter(results_cache))
        del results_cache[oldest_key]

    return results

@app.route('/', methods=['GET', 'POST'])
def home():
    """Home page with search form."""
    if request.method == 'POST':
        # Redirect POST requests to search endpoint
        return search()
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    """Search endpoint that returns results page."""
    import time
    start_time = time.time()
    
    print("\n" + "="*50)
    print("POST request received")
    print("Form data:", request.form)
    
    # Get search type
    search_type = request.form.get('search_type', 'single')
    print(f"Search type: {search_type}")
    
    # Get selected table and search mode
    selected_table = request.form.get('table', '')
    search_mode = request.form.get('search_mode', 'partial')
    print(f"Selected table: {selected_table}")
    print(f"Search mode: {search_mode}")
    
    results = []
    all_criteria_list = []
    
    if search_type == 'single':
        # Single search - same as before
        criteria = {
            'perid': request.form.get('perid', ''),
            'perfamno': request.form.get('perfamno', ''),
            'persubno': request.form.get('persubno', ''),
            'perfirst': request.form.get('perfirst', ''),
            'perfather': request.form.get('perfather', ''),
            'pergrand': request.form.get('pergrand', ''),
            'perfourth_name': request.form.get('perfourth_name', ''),
            'mothername': request.form.get('mothername', ''),
            'pergender': request.form.get('pergender', ''),
            'perdob': request.form.get('perdob', ''),
            'perbthcountry': request.form.get('perbthcountry', ''),
            'pertelno': request.form.get('pertelno', ''),
            'vrcname': request.form.get('vrcname', ''),
            'pcname': request.form.get('pcname', ''),
            'peraddress': request.form.get('peraddress', ''),
            'pernationality': request.form.get('pernationality', '')
        }
        
        # Filter out empty criteria
        criteria = {k: v for k, v in criteria.items() if v and v.strip()}
        if criteria:
            all_criteria_list = [criteria]
    
    elif search_type == 'multi':
        # Multi search - parse multiple values
        multi_criteria = {}
        
        # Parse each multi field
        for field in ['perid', 'perfamno', 'persubno', 'perfirst', 'perfather', 'pergrand']:
            multi_values = parse_multi_values(request.form.get(f'{field}_multi', ''))
            if multi_values:
                multi_criteria[field] = multi_values
        
        # Generate all combinations of criteria
        if multi_criteria:
            from itertools import product
            keys = list(multi_criteria.keys())
            values = list(multi_criteria.values())
            
            for combo in product(*values):
                criteria = dict(zip(keys, combo))
                all_criteria_list.append(criteria)
    
    elif search_type == 'excel':
        # Excel upload
        if 'excel_file' not in request.files:
            return render_template('results.html', results=[], criteria={'error': 'لم يتم رفع ملف'})
        
        file = request.files['excel_file']
        if file.filename == '' or not allowed_file(file.filename):
            return render_template('results.html', results=[], criteria={'error': 'نوع الملف غير مدعوم'})
        
        # Save file
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Process Excel file
        all_criteria_list = process_excel_file(file_path)
        
        # Clean up file
        try:
            os.remove(file_path)
        except:
            pass
    
    print(f"Total search criteria to process: {len(all_criteria_list)}")
    
    # Process all criteria
    for i, criteria in enumerate(all_criteria_list):
        if not criteria:
            continue
            
        print(f"Processing criteria {i+1}/{len(all_criteria_list)}: {criteria}")
        
        try:
            batch_results = search_databases_optimized(criteria, selected_table, search_mode)
            results.extend(batch_results)
            
            # Limit total results to prevent memory issues
            if len(results) >= 5000:
                results = results[:5000]
                break
                
        except Exception as e:
            print(f"Error processing criteria {criteria}: {e}")
            continue
    
    elapsed_time = time.time() - start_time
    print(f"Search completed in {elapsed_time:.2f} seconds")
    print(f"Found {len(results)} results")
    
    # Prepare display criteria
    search_criteria = {
        'search_type': search_type,
        'search_mode': search_mode,
        'total_criteria': len(all_criteria_list)
    }
    
    if selected_table:
        search_criteria['table'] = selected_table
        search_criteria['table_name'] = table_to_db.get(selected_table, {}).get('name', selected_table)
    
    print("="*50 + "\n")
    
    return render_template('results.html', results=results, criteria=search_criteria)

@app.route('/download_template')
def download_template():
    """Download the Excel search template."""
    template_path = os.path.join(os.getcwd(), 'template.xlsx')
    
    if os.path.exists(template_path):
        return send_file(
            template_path,
            as_attachment=True,
            download_name='template.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        return "Template file not found", 404

if __name__ == '__main__':
    app.run(debug=True)