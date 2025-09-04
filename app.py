from flask import Flask, render_template, request, jsonify, redirect, url_for, session, flash, send_file
from flask_mysqldb import MySQL
import MySQLdb.cursors
import config
from decimal import Decimal
from functools import wraps
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config.from_object(config.Config)
app.secret_key = 'your-secret-key-here'  # Ganti dengan secret key yang aman
mysql = MySQL(app)

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def get_summary_stats():
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
    
    # Get total realisasi and outstanding
    cursor.execute("""
        SELECT 
            COUNT(*) as total_mitra,
            SUM(Jumlah_Realisasi_Rp) as total_realisasi,
            SUM(Outstanding_Rp) as total_outstanding
        FROM mitra_binaan
    """)
    stats = cursor.fetchone()
    
    # Get kolektabilitas distribution
    cursor.execute("""
        SELECT 
            Kolektabilitas_BUMN,
            COUNT(*) as count
        FROM mitra_binaan
        GROUP BY Kolektabilitas_BUMN
    """)
    kolektabilitas = cursor.fetchall()
    
    return stats, kolektabilitas

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
        
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        remember = request.form.get('remember_me')
        
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute('SELECT * FROM users WHERE email = %s AND password = %s', (email, password))
        user = cursor.fetchone()
        
        if user:
            session['user_id'] = user['id']
            session['name'] = user['name']
            session.permanent = True if remember else False
            return redirect(url_for('dashboard'))
        else:
            flash('Email atau password salah!')
    
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
        
    if request.method == 'POST':
        name = f"{request.form['first_name']} {request.form['last_name']}"
        email = request.form['email']
        password = request.form['password']
        confirm_password = request.form['confirm_password']
        terms = request.form.get('terms')
        
        if not terms:
            flash('Anda harus menyetujui Syarat dan Ketentuan!')
            return render_template('register.html')
            
        if password != confirm_password:
            flash('Password tidak cocok!')
            return render_template('register.html')
        
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        try:
            # Check if email already exists
            cursor.execute('SELECT * FROM users WHERE email = %s', (email,))
            if cursor.fetchone():
                flash('Email sudah terdaftar!')
                return render_template('register.html')
                
            cursor.execute('INSERT INTO users (name, email, password) VALUES (%s, %s, %s)',
                         (name, email, password))
            mysql.connection.commit()
            flash('Registrasi berhasil! Silakan login.')
            return redirect(url_for('login'))
        except Exception as e:
            print(f"Error during registration: {str(e)}")
            flash('Terjadi kesalahan saat mendaftar. Silakan coba lagi.')
            return render_template('register.html')
        finally:
            cursor.close()
    
    return render_template('register.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/export_excel')
def export_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
        
    try:
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        
        # Ambil data mitra
        cursor.execute('SELECT * FROM mitra_binaan ORDER BY No')
        data = cursor.fetchall()
        
        # Buat Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data PUMK"
        
        # Tulis header
        headers = [
            'No', 'Nama Mitra Binaan', 'Alamat', 'Provinsi', 'Kabupaten/Kota',
            'Jumlah Realisasi (Rp)', 'Outstanding (Rp)', 'Kolektabilitas'
        ]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
            
        # Style header
        header_font = Font(bold=True)
        header_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
        
        # Tulis data
        for row, item in enumerate(data, 2):
            ws.cell(row=row, column=1, value=item['No'])
            ws.cell(row=row, column=2, value=item['Nama_Mitra_Binaan'])
            ws.cell(row=row, column=3, value=item['Alamat'])
            ws.cell(row=row, column=4, value=item['Provinsi'])
            ws.cell(row=row, column=5, value=item['Kabupaten_Kota'])
            ws.cell(row=row, column=6, value=item['Jumlah_Realisasi_Rp'])
            ws.cell(row=row, column=7, value=item['Outstanding_Rp'])
            ws.cell(row=row, column=8, value=item['Kolektabilitas_BUMN'])
            
        # Format currency columns
        currency_format = '#,##0'
        for row in range(2, len(data) + 2):
            ws.cell(row=row, column=6).number_format = currency_format
            ws.cell(row=row, column=7).number_format = currency_format
            
        # Adjust column widths
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            ws.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 2
            
        # Create response
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'PUMK_Data_{datetime.now().strftime("%Y%m%d")}.xlsx'
        )
        
    except Exception as e:
        print(f"Error exporting Excel: {str(e)}")
        flash('Terjadi kesalahan saat mengexport data', 'error')
        return redirect(url_for('dashboard'))
        
    finally:
        cursor.close()

@app.route('/')
@login_required
def index():
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
@login_required
def dashboard():
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
    
    # Get base data
    query = """
        SELECT * FROM mitra_binaan
        ORDER BY No
        LIMIT 1000
    """
    cursor.execute(query)
    data = cursor.fetchall()
    
    # Get summary statistics
    stats, kolektabilitas = get_summary_stats()
    
    return render_template('dashboard.html', 
                         mitra=data, 
                         stats=stats,
                         kolektabilitas=kolektabilitas)

@app.route('/mitra-binaan')
@login_required
def mitra_binaan():
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
    cursor.execute('SELECT * FROM mitra_binaan ORDER BY No')
    mitra = cursor.fetchall()
    return render_template('mitra_binaan.html', mitra=mitra)

@app.route('/keuangan')
@login_required
def keuangan():
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
    # Get summary statistics
    stats, kolektabilitas = get_summary_stats()
    return render_template('keuangan.html', stats=stats, kolektabilitas=kolektabilitas)

@app.route('/laporan')
@login_required
def laporan():
    return render_template('laporan.html')

@app.route('/api/filter')
def filter_data():
    provinsi = request.args.get('provinsi')
    kolektabilitas = request.args.get('kolektabilitas')
    
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
    
    query = "SELECT * FROM mitra_binaan WHERE 1=1"
    if provinsi:
        query += f" AND Provinsi = '{provinsi}'"
    if kolektabilitas:
        query += f" AND Kolektabilitas_BUMN = '{kolektabilitas}'"
    
    cursor.execute(query)
    data = cursor.fetchall()
    return jsonify(data)

if __name__ == '__main__':
    app.run(debug=True)
