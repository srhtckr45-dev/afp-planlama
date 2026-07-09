import sqlite3
import os
import pandas as pd

DB_PATH = os.path.join(os.path.dirname(__file__), "afp_database.db")

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db(df=None, df_genel=None):
    """
    Initializes the database. If a pandas DataFrame is provided (from migration),
    it will create the tables and load the initial data.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if df is not None:
        # Save dataframe to SQLite table 'programs'
        # We replace the table if it exists during migration
        df.to_sql("programs", conn, if_exists="replace", index=False)
        print("Initial data loaded into 'programs' table.")
        
        # Ensure we have indexes for fast lookups
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_programs_lot ON programs (LOT)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_programs_makine ON programs (MAKİNE)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_programs_durum ON programs ([SİPARİŞ DURUMU])")
        
    if df_genel is not None:
        df_genel.to_sql("genel_makine", conn, if_exists="replace", index=False)
        print("Initial data loaded into 'genel_makine' table.")
        
    conn.commit()
    conn.close()

def get_genel_makine():
    """
    Returns the GENEL MAKINE data as a DataFrame.
    """
    conn = get_db_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM genel_makine", conn)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df

def update_genel_makine(df):
    """
    Updates the genel_makine table with the provided DataFrame.
    """
    conn = get_db_connection()
    df.to_sql("genel_makine", conn, if_exists="replace", index=False)
    conn.commit()
    conn.close()

def get_machines():
    """
    Returns a sorted list of unique machines available in the database.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT MAKİNE FROM programs WHERE MAKİNE IS NOT NULL AND MAKİNE != '' ORDER BY MAKİNE")
    machines = [row[0] for row in cursor.fetchall()]
    conn.close()
    return machines

def get_active_queue(machine_name):
    """
    Returns all active (NOT FINISHED) lots for a machine, sorted by order ('sıra').
    """
    conn = get_db_connection()
    # We query using pandas for easy table visualization in streamlit
    query = """
        SELECT * FROM programs 
        WHERE MAKİNE = ? AND ([SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = '')
        ORDER BY CAST([sıra] AS INTEGER) ASC, LOT ASC
    """
    df = pd.read_sql_query(query, conn, params=(machine_name,))
    conn.close()
    return df

def get_completed_lots(machine_name=None):
    """
    Returns completed (BİTTİ) lots, optionally filtered by machine.
    """
    conn = get_db_connection()
    if machine_name:
        query = "SELECT * FROM programs WHERE MAKİNE = ? AND [SİPARİŞ DURUMU] = 'BİTTİ' ORDER BY LOT DESC"
        df = pd.read_sql_query(query, conn, params=(machine_name,))
    else:
        query = "SELECT * FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTTİ' ORDER BY LOT DESC LIMIT 500"
        df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_lot_details(lot_no):
    """
    Returns the details of a specific lot as a dictionary.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM programs WHERE LOT = ?", (str(lot_no),))
    row = cursor.fetchone()
    conn.close()
    if row:
        return dict(row)
    return None

def update_lot_sequence(lot_order_list):
    """
    Updates the queue order ('sıra') for a list of lot numbers.
    lot_order_list should be a list of lot numbers in the desired sequence.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    for idx, lot in enumerate(lot_order_list, 1):
        cursor.execute("UPDATE programs SET [sıra] = ? WHERE LOT = ?", (float(idx), str(lot)))
    conn.commit()
    conn.close()

def update_lot_fields(lot_no, fields_dict):
    """
    Updates specific columns of a lot record in the database.
    fields_dict: dict of {column_name: new_value}
    """
    if not fields_dict:
        return
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Construct dynamic SQL update
    set_clause = ", ".join([f"[{k}] = ?" for k in fields_dict.keys()])
    values = list(fields_dict.values())
    values.append(str(lot_no))
    
    query = f"UPDATE programs SET {set_clause} WHERE LOT = ?"
    cursor.execute(query, values)
    conn.commit()
    conn.close()

def add_new_lot(fields_dict):
    """
    Inserts a new lot into the programs table.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Filter out empty keys
    fields_dict = {k: v for k, v in fields_dict.items() if k}
    
    columns = ", ".join([f"[{k}]" for k in fields_dict.keys()])
    placeholders = ", ".join(["?" for _ in fields_dict])
    values = list(fields_dict.values())
    
    query = f"INSERT INTO programs ({columns}) VALUES ({placeholders})"
    cursor.execute(query, values)
    conn.commit()
    conn.close()

def delete_lot(lot_no):
    """
    Deletes a lot from the database.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM programs WHERE LOT = ?", (str(lot_no),))
    conn.commit()
    conn.close()

def sync_db_to_excel(excel_path=None, modified_machines=None):
    """
    Saves the entire SQLite database back to the original Excel file using Excel COM.
    Preserves 100% of formatting, conditional formatting, slicers, and gridlines.
    """
    import shutil
    import os
    import sys
    import ctypes
    
    # Dynamically inject DLL search paths and manually load pywin32 DLLs on Windows
    if sys.platform == 'win32':
        py_root = os.path.dirname(sys.executable)
        pywin32_sys32 = os.path.join(py_root, 'Lib', 'site-packages', 'pywin32_system32')
        if os.path.exists(pywin32_sys32):
            try:
                os.add_dll_directory(pywin32_sys32)
            except Exception:
                pass
        if os.path.exists(py_root):
            try:
                os.add_dll_directory(py_root)
            except Exception:
                pass
                
        # Manually load DLLs using ctypes to bypass loader path errors
        dll_dirs = [py_root, pywin32_sys32]
        loaded_pw = False
        loaded_pc = False
        for d in dll_dirs:
            if os.path.exists(d):
                try:
                    for f in os.listdir(d):
                        if f.startswith("pywintypes") and f.endswith(".dll") and not loaded_pw:
                            ctypes.CDLL(os.path.join(d, f))
                            loaded_pw = True
                        if f.startswith("pythoncom") and f.endswith(".dll") and not loaded_pc:
                            ctypes.CDLL(os.path.join(d, f))
                            loaded_pc = True
                except Exception:
                    pass
                
    import win32com.client
    
    if excel_path is None:
        excel_path = r"C:\Users\serhat.cakir\OneDrive\Belgeler\AFP.xlsx"
        
    excel_path = os.path.abspath(excel_path)
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Orijinal Excel dosyası bulunamadı: {excel_path}")
        
    # Helper to convert python/numpy types to Excel compatible types
    def clean_val(val):
        if pd.isna(val) or val is None:
            return None
        val_str = str(val).strip()
        if not val_str or val_str.lower() in ["nan", "none"]:
            return None
        if val_str.isdigit():
            return int(val_str)
        try:
            return float(val_str)
        except ValueError:
            return val_str
            
    # 1. Create a backup of the Excel file
    backup_path = excel_path + ".bak"
    shutil.copy2(excel_path, backup_path)
    
    # 2. Fetch current database table
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM programs", conn)
    conn.close()
    
    # 3. Connect to Excel Application via COM
    excel = None
    wb = None
    is_attached = False
    
    file_name = os.path.basename(excel_path)
    
    try:
        # Try to attach to already running Excel application instance
        excel = win32com.client.GetActiveObject("Excel.Application")
        for open_wb in excel.Workbooks:
            if open_wb.Name == file_name:
                wb = open_wb
                is_attached = True
                break
    except Exception:
        pass
        
    if excel is None:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
    try:
        if wb is None:
            wb = excel.Workbooks.Open(excel_path)
            
        # Update 'PROGRAMLAR' sheet
        if "PROGRAMLAR" in [s.Name for s in wb.Sheets]:
            ws_prog = wb.Sheets("PROGRAMLAR")
            
            # Fetch headers from row 1
            num_cols = 1
            headers = []
            while True:
                val = ws_prog.Cells(1, num_cols).Value
                if not val:
                    break
                headers.append(val)
                num_cols += 1
            num_cols -= 1
            
            # Identify formula columns in Tablo1 if it exists
            formula_cols = set()
            is_table = False
            
            if "Tablo1" in [t.Name for t in ws_prog.ListObjects]:
                is_table = True
                tbl = ws_prog.ListObjects("Tablo1")
                if tbl.DataBodyRange:
                    # Check first row of data for formulas
                    for c in range(1, tbl.ListColumns.Count + 1):
                        try:
                            if tbl.DataBodyRange.Cells(1, c).HasFormula:
                                formula_cols.add(c)
                        except Exception:
                            pass
                    # Delete the existing rows to reset the table
                    tbl.DataBodyRange.Delete()
            else:
                # Fallback if it's not a table
                ws_prog.Rows("2:10000").ClearContents()
                
            num_rows = len(df)
            if num_rows > 0:
                # If it's a table, resize it first so formulas autofill down
                if is_table:
                    tbl.Resize(ws_prog.Range(ws_prog.Cells(1, 1), ws_prog.Cells(num_rows + 1, num_cols)))
                    
                # Find contiguous column chunks that do NOT have formulas
                chunks = []
                current_chunk = []
                for c in range(1, num_cols + 1):
                    if c not in formula_cols:
                        current_chunk.append(c)
                    else:
                        if current_chunk:
                            chunks.append(current_chunk)
                            current_chunk = []
                if current_chunk:
                    chunks.append(current_chunk)
                    
                # Prepare matrices for each chunk in a single pass
                chunk_matrices = {i: [] for i in range(len(chunks))}
                chunk_headers = {i: [headers[c - 1] for c in chunk] for i, chunk in enumerate(chunks)}
                
                for idx, row in df.iterrows():
                    for i, chunk in enumerate(chunks):
                        row_vals = []
                        for h in chunk_headers[i]:
                            val = row[h] if h in df.columns else None
                            row_vals.append(clean_val(val))
                        chunk_matrices[i].append(row_vals)
                        
                # Write each chunk to Excel
                for i, chunk in enumerate(chunks):
                    start_col = chunk[0]
                    end_col = chunk[-1]
                    write_range = ws_prog.Range(ws_prog.Cells(2, start_col), ws_prog.Cells(num_rows + 1, end_col))
                    write_range.Value = chunk_matrices[i]
                
        # Update 'GENEL MAKİNE' sheet
        target_sheet = None
        for s in wb.Sheets:
            if s.Name.startswith("GENEL MAK"):
                target_sheet = s
                break
                
        if target_sheet and (modified_machines is None or "GENEL MAKİNE" in modified_machines):
            df_genel = get_genel_makine()
            if not df_genel.empty:
                ws_genel = target_sheet
                num_cols_g = 1
                headers_g = []
                while True:
                    try:
                        val = ws_genel.Cells(1, num_cols_g).Value
                        if not val:
                            break
                        headers_g.append(val)
                        num_cols_g += 1
                    except Exception:
                        break
                num_cols_g -= 1
                
                formula_cols_g = set()
                is_table_g = False
                
                if len([t.Name for t in ws_genel.ListObjects]) > 0:
                    is_table_g = True
                    tbl_g = ws_genel.ListObjects(1)
                    if tbl_g.DataBodyRange:
                        for c in range(1, tbl_g.ListColumns.Count + 1):
                            try:
                                if tbl_g.DataBodyRange.Cells(1, c).HasFormula:
                                    formula_cols_g.add(c)
                            except Exception:
                                pass
                        tbl_g.DataBodyRange.Delete()
                else:
                    for c in range(1, num_cols_g + 1):
                        try:
                            if ws_genel.Cells(2, c).HasFormula:
                                formula_cols_g.add(c)
                        except Exception:
                            pass
                    ws_genel.Rows("2:10000").ClearContents()
                
                num_rows_g = len(df_genel)
                if num_rows_g > 0:
                    if is_table_g:
                        tbl_g.Resize(ws_genel.Range(ws_genel.Cells(1, 1), ws_genel.Cells(num_rows_g + 1, num_cols_g)))
                        
                    chunks_g = []
                    current_chunk_g = []
                    for c in range(1, num_cols_g + 1):
                        if c not in formula_cols_g:
                            current_chunk_g.append(c)
                        else:
                            if current_chunk_g:
                                chunks_g.append(current_chunk_g)
                                current_chunk_g = []
                    if current_chunk_g:
                        chunks_g.append(current_chunk_g)
                        
                    chunk_matrices_g = {i: [] for i in range(len(chunks_g))}
                    chunk_headers_g = {i: [headers_g[c - 1] for c in chunk] for i, chunk in enumerate(chunks_g)}
                    
                    for idx, row in df_genel.iterrows():
                        for i, chunk in enumerate(chunks_g):
                            row_vals = []
                            for h in chunk_headers_g[i]:
                                val = row[h] if h in df_genel.columns else None
                                row_vals.append(clean_val(val))
                            chunk_matrices_g[i].append(row_vals)
                            
                    for i, chunk in enumerate(chunks_g):
                        start_col = chunk[0]
                        end_col = chunk[-1]
                        write_range = ws_genel.Range(ws_genel.Cells(2, start_col), ws_genel.Cells(num_rows_g + 1, end_col))
                        write_range.Value = chunk_matrices_g[i]

        # Get active machine list
        machines = get_machines()
        
        # If user passed a specific list of modified machines, ONLY sync those sekmeler!
        if modified_machines is not None and len(modified_machines) > 0:
            machines = [m for m in machines if m in modified_machines]
        
        # Update each machine sheet
        for m in machines:
            if m in [s.Name for s in wb.Sheets]:
                ws_mach = wb.Sheets(m)
                
                # Fetch active queue lots from SQLite
                active_q = get_active_queue(m)
                active_lots = active_q["LOT"].tolist()
                
                # Clear Column A from row 3 to 500 in one bulk call (preserves formulas in columns B-AH!)
                ws_mach.Range("A3:A500").ClearContents()
                
                if len(active_lots) > 0:
                    # Write all lot numbers in ONE single COM call (100x faster!)
                    lot_matrix = [[clean_val(lot)] for lot in active_lots]
                    ws_mach.Range(f"A3:A{2 + len(active_lots)}").Value = lot_matrix
                    
                    # Read formulas for Column B in bulk to check if we need to copy VLOOKUP formulas
                    b_formulas = ws_mach.Range(f"B3:B{2 + len(active_lots)}").Formula
                    if not isinstance(b_formulas, tuple):
                        b_formulas = ((b_formulas,),)
                        
                    formula_b3 = ws_mach.Range("B3:AI3").FormulaArray
                    
                    for idx in range(1, len(active_lots)):
                        r = 3 + idx
                        current_f = b_formulas[idx][0]
                        # Only copy if the cell doesn't have a formula starting with "="
                        if not current_f or not str(current_f).startswith("="):
                            new_formula = formula_b3.replace("A3", f"A{r}")
                            ws_mach.Range(f"B{r}:AI{r}").FormulaArray = new_formula
                            
        wb.Save()
        
        # Only close and quit if we opened a background instance
        if not is_attached:
            wb.Close()
            excel.Quit()
            
    except Exception as e:
        raise e
    finally:
        # If we opened in background and didn't close yet due to error
        if not is_attached:
            try:
                excel.Quit()
            except Exception:
                pass
