import sqlite3
import pandas as pd

DB_NAME = "indonesia_companies.db"
EXCEL_OUTPUT = "Indonesia_FnB_OilGas_Directory.xlsx"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT UNIQUE,
            sector TEXT,
            sub_industry TEXT,
            address TEXT,
            website TEXT,
            phone TEXT,
            email TEXT,
            contact_person_name TEXT,
            contact_person_title TEXT,
            contact_person_linkedin TEXT,
            last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def seed_initial_data():
    initial_companies = [
        ("PT Indofood Sukses Makmur Tbk", "Food & Beverage", "Consumer Goods / Packaged Food", "Sudirman Plaza, Indofood Tower, Jakarta", "https://www.indofood.com", "+62 21 57958822", "corporate@indofood.co.id", "Corporate Secretary", "Investor Relations / Procurement", "https://linkedin.com/company/indofood"),
        ("PT Mayora Indah Tbk", "Food & Beverage", "Confectionery / Biscuits", "Gedung Mayora, Jl. Tomang Raya, Jakarta", "https://www.mayoraindah.co.id", "+62 21 5655322", "consumer@mayora.co.id", "Procurement Dept", "Purchasing Manager", ""),
        ("PT Cisarua Mountain Dairy Tbk (Cimory)", "Food & Beverage", "Dairy & Beverages", "Bogor, West Java", "https://cimory.com", "+62 21 5818288", "contact@cimory.com", "Sales & Business Dev", "Commercial Lead", ""),
        ("PT Ultra Jaya Milk Industry Tbk", "Food & Beverage", "Dairy & Beverages", "Padalarang, Bandung", "https://www.ultrajaya.co.id", "+62 22 86700700", "investor@ultrajaya.co.id", "Supplier Desk", "Procurement Specialist", ""),
        ("PT Pertamina (Persero)", "Oil & Gas", "National Oil Company (Upstream/Downstream)", "Jl. Medan Merdeka Timur No.1A, Jakarta", "https://www.pertamina.com", "+62 21 135", "pcc135@pertamina.com", "Vendor Care / Procurement", "Vendor Relations Manager", "https://linkedin.com/company/pertamina"),
        ("PT Medco Energi Internasional Tbk", "Oil & Gas", "E&P / Upstream Exploration", "The Energy Building, SCBD, Jakarta", "https://www.medcoenergi.com", "+62 21 52963000", "investor.relations@medcoenergi.com", "Supply Chain Mgmt", "Procurement Manager", "https://linkedin.com/company/medco-energi"),
        ("BP Indonesia", "Oil & Gas", "LNG / Upstream Operator", "Perkantoran Hijau Arkadia, Jakarta", "https://www.bp.com/indonesia", "+62 21 78838000", "indonesiaenquiries@bp.com", "Supply Chain Dept", "PSC Procurement Specialist", ""),
        ("PT Elnusa Tbk", "Oil & Gas Services", "Seismic & Drilling Services", "Gedung Graha Elnusa, Jakarta", "https://www.elnusa.co.id", "+62 21 78830850", "corporate@elnusa.co.id", "Commercial Team", "Business Development Manager", "")
    ]
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    for item in initial_companies:
        cursor.execute('''
            INSERT OR IGNORE INTO companies (
                company_name, sector, sub_industry, address, website, phone, email, 
                contact_person_name, contact_person_title, contact_person_linkedin
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', item)
    conn.commit()
    conn.close()

def export_to_excel():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM companies", conn)
    conn.close()

    if df.empty:
        return

    df = df.drop(columns=['id'])
    df.columns = [
        'Company Name', 'Sector', 'Sub-Industry', 'HQ Address', 'Website',
        'Company Phone', 'Company Email', 'Contact Person Name/Dept', 
        'Contact Title', 'Contact LinkedIn', 'Last Updated'
    ]

    with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ID_FnB_OilGas_Directory', index=False)
        worksheet = writer.sheets['ID_FnB_OilGas_Directory']
        
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

if __name__ == "__main__":
    init_db()
    seed_initial_data()
    export_to_excel()
