import os
import shutil
import pandas as pd
import sqlite3
import subprocess
from database import init_db, DB_PATH

ORIGINAL_EXCEL_PATH = r"C:\Users\serhat.cakir\OneDrive\Belgeler\AFP.xlsx"
TEMP_EXCEL_PATH = os.path.join(os.path.dirname(__file__), "AFP_temp.xlsx")

def run_migration(excel_url=None):
    if excel_url:
        print(f"Cloud mode: downloading from URL...")
        import requests
        try:
            r = requests.get(excel_url, timeout=30)
            r.raise_for_status()
            with open(TEMP_EXCEL_PATH, 'wb') as f:
                f.write(r.content)
            copied = True
            print("Successfully downloaded Excel from URL.")
        except Exception as e:
            print(f"Failed to download Excel from URL: {e}")
            raise Exception(f"Buluttan (OneDrive) Excel dosyası indirilemedi: {e}")
    else:
        print(f"Checking original Excel file at: {ORIGINAL_EXCEL_PATH}")
        if not os.path.exists(ORIGINAL_EXCEL_PATH):
            print(f"Error: Original Excel file not found at {ORIGINAL_EXCEL_PATH}")
            return
            
        try:
            print("Creating temporary copy of Excel file...")
            copied = False
            
            # 1. Try win32com first
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                excel = win32com.client.GetActiveObject("Excel.Application")
                for wb in excel.Workbooks:
                    if wb.Name == os.path.basename(ORIGINAL_EXCEL_PATH):
                        wb.SaveCopyAs(TEMP_EXCEL_PATH)
                        copied = True
                        print("Copied successfully using COM (live Excel data).")
                        break
            except Exception as e:
                print(f"COM copy failed: {e}")
            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
                
            # 2. Shell copy fallback
            if not copied:
                cmd = f'cmd.exe /c copy /y "{ORIGINAL_EXCEL_PATH}" "{TEMP_EXCEL_PATH}"'
                res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
                if os.path.exists(TEMP_EXCEL_PATH):
                    copied = True
                    print("Copied using shell copy.")
                    
            # 3. Shutil Fallback
            if not copied:
                shutil.copy2(ORIGINAL_EXCEL_PATH, TEMP_EXCEL_PATH)
                print("Copied using shutil.")
        except Exception as e:
            raise Exception(f"Excel kopyalanamadı: {e}")
            
    try:
        print("Reading Excel sheets (this may take a few seconds)...")
        # Load the spreadsheet
        xl = pd.ExcelFile(TEMP_EXCEL_PATH)
        df = xl.parse("PROGRAMLAR")
        
        # Try to find GENEL MAKINE sheet
        genel_sheet_name = next((s for s in xl.sheet_names if s.upper().startswith("GENEL MAK")), None)
        if genel_sheet_name:
            print(f"Reading '{genel_sheet_name}' sheet...")
            df_genel = xl.parse(genel_sheet_name)
            # Clean column names
            df_genel.columns = [str(c).strip() for c in df_genel.columns]
        else:
            df_genel = None
            
        print(f"Loaded {len(df)} rows from PROGRAMLAR.")
        
        # Data cleaning:
        # 1. Clean LOT column (convert to string, drop NaN rows, strip whitespace)
        print("Cleaning LOT numbers...")
        df = df.dropna(subset=["LOT"])
        df["LOT"] = df["LOT"].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x).str.strip()
        
        # 2. Clean MAKİNE column
        if "MAKİNE" in df.columns:
            df["MAKİNE"] = df["MAKİNE"].astype(str).str.strip()
            # Remove rows where MAKİNE contains 'ÜRETİLEMİYOR'
            df = df[~df["MAKİNE"].str.upper().str.contains("ÜRETİLEMİYOR", na=False)]
            df.loc[df["MAKİNE"].isin(["nan", "None", ""]), "MAKİNE"] = None
            
        # 3. Clean SİPARİŞ DURUMU column (strip whitespace)
        status_col = [c for c in df.columns if "durum" in c.lower()][0]
        df[status_col] = df[status_col].astype(str).str.strip()
        df.loc[df[status_col].isin(["nan", "None", ""]), status_col] = None
        
        # 4. Clean column names
        df.columns = [str(c).strip() for c in df.columns]
        
        # 5. Handle datetime columns to strings for SQLite compatibility
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                
        print(f"Total cleaned rows to migrate: {len(df)}")
        
        # Load into SQLite
        print("Migrating to SQLite...")
        init_db(df, df_genel=df_genel)
        
        print("Parsing X cutoffs from machine sheets...")
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(programs)")
        cols = [c[1] for c in cursor.fetchall()]
        if "X_CUTOFF" not in cols:
            cursor.execute("ALTER TABLE programs ADD COLUMN X_CUTOFF INTEGER DEFAULT 0")
            
        for s in xl.sheet_names:
            if s[0] in ['B', 'V', 'N'] and len(s) <= 3:
                try:
                    df_m = xl.parse(s, header=None, usecols=[0])
                    x_idx = df_m.index[df_m[0].astype(str).str.upper() == 'X'].tolist()
                    if x_idx:
                        idx = x_idx[0]
                        prev_val = str(df_m.iloc[idx-1, 0])
                        print(f"Machine {s}: Found X at index {idx}, prev_val: {repr(prev_val)}")
                        if prev_val != 'nan' and prev_val != 'LOT NO':
                            clean_prev_val = prev_val.split('.')[0].strip()
                            print(f"Machine {s}: Updating LOT {repr(clean_prev_val)}")
                            cursor.execute("UPDATE programs SET X_CUTOFF = 0 WHERE MAKİNE = ?", (s,))
                            cursor.execute("UPDATE programs SET X_CUTOFF = 1 WHERE MAKİNE = ? AND LOT = ?", (s, clean_prev_val))
                            print(f"Updated rows: {cursor.rowcount}")
                        elif prev_val == 'LOT NO':
                            cursor.execute("UPDATE programs SET X_CUTOFF = 0 WHERE MAKİNE = ?", (s,))
                            cursor.execute("UPDATE programs SET X_CUTOFF = 1 WHERE MAKİNE = ? AND [sıra] = (SELECT MIN([sıra]) FROM programs WHERE MAKİNE = ?)", (s, s))
                            print(f"Machine {s}: Updated top row. Rows affected: {cursor.rowcount}")
                            
                    # Check for manual additions, reorderings, and deletions in the machine sheet
                    sheet_lots_raw = df_m[0].astype(str).tolist()
                    sheet_lots = []
                    for raw_l in sheet_lots_raw:
                        l_clean = str(raw_l).split('.')[0].strip()
                        if l_clean and l_clean.upper() not in ['NAN', 'NONE', 'LOT NO', 'X']:
                            if l_clean not in sheet_lots:
                                sheet_lots.append(l_clean)
                    
                    if len(sheet_lots) > 0:
                        # 1. Sync additions and reordering: Update existing LOTs to match this machine and exact row order
                        for index, lot_num in enumerate(sheet_lots):
                            cursor.execute('''
                                UPDATE programs 
                                SET MAKİNE = ?, [SİPARİŞ DURUMU] = 'BİTMEDİ', [sıra] = ? 
                                WHERE LOT = ?
                            ''', (s, index, lot_num))
                            
                        # 2. Sync deletions: Find LOTs in database that should be here but aren't
                        cursor.execute("SELECT LOT FROM programs WHERE MAKİNE = ? AND [SİPARİŞ DURUMU] = 'BİTMEDİ'", (s,))
                        db_lots = [str(row[0]).strip() for row in cursor.fetchall()]
                        deleted_count = 0
                        for db_lot in db_lots:
                            if db_lot not in sheet_lots:
                                cursor.execute("UPDATE programs SET [SİPARİŞ DURUMU] = 'BİTTİ' WHERE MAKİNE = ? AND LOT = ?", (s, db_lot))
                                deleted_count += 1
                        if deleted_count > 0:
                            print(f"Machine {s}: Marked {deleted_count} manually deleted LOTs as 'BİTTİ'")
                            
                except Exception as e:
                    print(f"Machine {s} parser error: {e}")
                    pass
        conn.commit()
        conn.close()
        
        # Verify
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM programs")
        db_count = cursor.fetchone()[0]
        conn.close()
        
        print("\n" + "="*40)
        print(f"Migration completed successfully!")
        print(f"Excel row count: {len(df)}")
        print(f"Database row count: {db_count}")
        print("="*40)
        return True
        
    except Exception as e:
        print(f"\nMigration failed with error: {e}")
        import traceback
        traceback.print_exc()
        raise Exception(f"Excel dosyası okunamadı veya başka bir hata oluştu: {e}. Lütfen Excel dosyasını KAPATIP tekrar deneyin.")
        
    finally:
        # Clean up temp file
        if os.path.exists(TEMP_EXCEL_PATH):
            try:
                os.remove(TEMP_EXCEL_PATH)
                print("Temporary copy cleaned up.")
            except Exception:
                pass

if __name__ == "__main__":
    run_migration()
