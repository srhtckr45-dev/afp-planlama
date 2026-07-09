import sqlite3
import pandas as pd

def run_autonomous_planning(db_path, locked_machines, urgent_lots_text, target_machine=None):
    """
    Executes the basic autonomous planning algorithm.
    """
    try:
        conn = sqlite3.connect(db_path)
        
        # 1. Parse urgent lots
        urgent_lots = []
        if urgent_lots_text:
            lines = urgent_lots_text.replace(',', '\\n').split('\\n')
            urgent_lots = [str(l).strip() for l in lines if str(l).strip()]
            
        # 2. Fetch all active orders
        query = "SELECT * FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = ''"
        df = pd.read_sql(query, conn)
        
        if df.empty:
            conn.close()
            return True, "Kuyrukta aktif iş bulunamadı."
            
        # Prepare for sorting
        if 'BOY' in df.columns:
            df['BOY_num'] = pd.to_numeric(df['BOY'], errors='coerce').fillna(0)
        else:
            df['BOY_num'] = 0
            
        if 'TERMİN TARİHİ' in df.columns:
            df['Termin_Date'] = pd.to_datetime(df['TERMİN TARİHİ'], format='%d.%m.%Y', errors='coerce')
        else:
            df['Termin_Date'] = pd.NaT
            
        # 2.5 Define Standard Equivalencies
        def map_standard_group(row):
            std = str(row.get('STANDART', '')).strip().upper()
            cap = str(row.get('ÇAP', '')).strip().upper()
            
            # Group 1: DIN933 and ISO4017 are identical except for M10, M12, M14
            if std in ['DIN933', 'ISO4017'] and cap not in ['M10', 'M12', 'M14']:
                return 'DIN933_ISO4017_GROUP'
                
            # Group 2: DIN931 and ISO4014 are identical except for M10, M12, M14
            if std in ['DIN931', 'ISO4014'] and cap not in ['M10', 'M12', 'M14']:
                return 'DIN931_ISO4014_GROUP'
                
            return std
            
        df['STANDART_GRUP'] = df.apply(map_standard_group, axis=1)
            
        # 3. Process machine by machine
        machines = df['MAKİNE'].dropna().unique()
        machines = [m for m in machines if str(m).strip().upper() not in ['B', 'N', 'V'] and 'ÜRETİLEMİYOR' not in str(m).strip().upper()]
        if target_machine and target_machine != 'Tüm Makineler':
            machines = [m for m in machines if m == target_machine]
        
        # We will keep a list of tuples (LOT, MAKİNE, new_sira)
        updates = []
        
        for mac in machines:
            df_m = df[df['MAKİNE'] == mac].copy()
            # Sort by current sequence to ensure we identify the X mark correctly
            if 'sıra' in df_m.columns:
                df_m['sıra_num'] = pd.to_numeric(df_m['sıra'], errors='coerce').fillna(999999)
                df_m = df_m.sort_values(by='sıra_num', na_position='last').drop(columns=['sıra_num'])
                
            locked_queue = pd.DataFrame()
            open_queue = pd.DataFrame()
            
            is_locked = any(str(m).strip().upper() == str(mac).strip().upper() for m in locked_machines)
            
            if is_locked and 'X_CUTOFF' in df_m.columns:
                cutoff_col = pd.to_numeric(df_m['X_CUTOFF'], errors='coerce').fillna(0)
                x_rows = df_m[cutoff_col == 1]
                if not x_rows.empty:
                    # Find the physical index of the first X mark
                    x_idx = x_rows.index[0]
                    # We need the positional index
                    pos_idx = df_m.index.get_loc(x_idx)
                    locked_queue = df_m.iloc[:pos_idx+1]
                    open_queue = df_m.iloc[pos_idx+1:]
                else:
                    open_queue = df_m
            else:
                open_queue = df_m
                
            # Extract urgent lots from open queue
            urgent_queue = pd.DataFrame()
            if urgent_lots and not open_queue.empty:
                urgent_mask = open_queue['LOT'].isin(urgent_lots)
                urgent_queue = open_queue[urgent_mask]
                open_queue = open_queue[~urgent_mask]
                
            # Sort urgent queue using traditional sort (urgent is urgent, just sort it basically)
            sort_cols = ['STANDART_GRUP', 'ÇAP', 'BOY_num', 'Termin_Date']
            sort_asc = [True, True, True, True]
            existing_cols = [c for c in sort_cols if c in open_queue.columns]
            existing_asc = [sort_asc[i] for i, c in enumerate(sort_cols) if c in open_queue.columns]
            
            if not urgent_queue.empty and existing_cols:
                urgent_queue = urgent_queue.sort_values(by=existing_cols, ascending=existing_asc, na_position='last')
                
            # Nearest Neighbor Optimization for open_queue
            optimized_open_queue = []
            
            if not open_queue.empty:
                # Determine current state to start NN
                current_cap = None
                current_std = None
                current_boy = None
                
                # Try to get state from urgent queue first, then locked queue
                last_known = None
                if not urgent_queue.empty:
                    last_known = urgent_queue.iloc[-1]
                elif not locked_queue.empty:
                    last_known = locked_queue.iloc[-1]
                    
                if last_known is not None:
                    current_cap = str(last_known.get('ÇAP', '')).strip().upper()
                    current_std = last_known.get('STANDART_GRUP', '')
                    current_boy = last_known.get('BOY_num', 0)
                
                from datetime import datetime
                today = pd.Timestamp(datetime.now().date())
                
                # Split into delayed and normal to strictly prioritize delayed
                is_delayed_mask = open_queue['Termin_Date'] < today
                delayed_df = open_queue[is_delayed_mask].copy()
                normal_df = open_queue[~is_delayed_mask].copy()
                
                def run_nn(df_to_process, curr_cap, curr_std, curr_boy):
                    result_list = []
                    # Convert to list of dicts for faster iteration
                    records = df_to_process.to_dict('records')
                    
                    while records:
                        best_idx = -1
                        min_cost = float('inf')
                        
                        for i, rec in enumerate(records):
                            cand_cap = str(rec.get('ÇAP', '')).strip().upper()
                            cand_std = rec.get('STANDART_GRUP', '')
                            cand_boy = rec.get('BOY_num', 0)
                            
                            cost = 0
                            
                            # Only apply penalties if we have a current state
                            if curr_cap is not None:
                                if cand_cap != curr_cap:
                                    cost += 10000  # Cap degisimi cok agir ceza
                                if cand_std != curr_std:
                                    cost += 1000   # Standart degisimi agir ceza
                                cost += abs(curr_boy - cand_boy) # Boy farki kadar ceza
                            else:
                                # No current state (first item), just pick the one with smallest boy to start nicely
                                cost = cand_boy
                                
                            # Tie breaker: Termin Date
                            # We add a tiny fraction based on termin date to prioritize earlier termins among identical items
                            term = rec.get('Termin_Date')
                            if pd.notnull(term):
                                # Number of days from today
                                days = (term - today).days
                                cost += (days * 0.001)
                                
                            if cost < min_cost:
                                min_cost = cost
                                best_idx = i
                                
                        # Pick the best
                        best_rec = records.pop(best_idx)
                        result_list.append(best_rec)
                        
                        # Update current state
                        curr_cap = str(best_rec.get('ÇAP', '')).strip().upper()
                        curr_std = best_rec.get('STANDART_GRUP', '')
                        curr_boy = best_rec.get('BOY_num', 0)
                        
                    return pd.DataFrame(result_list), curr_cap, curr_std, curr_boy

                # Process delayed first
                if not delayed_df.empty:
                    delayed_df, current_cap, current_std, current_boy = run_nn(delayed_df, current_cap, current_std, current_boy)
                    optimized_open_queue.append(delayed_df)
                    
                # Then normal
                if not normal_df.empty:
                    normal_df, current_cap, current_std, current_boy = run_nn(normal_df, current_cap, current_std, current_boy)
                    optimized_open_queue.append(normal_df)
                    
            if optimized_open_queue:
                open_queue = pd.concat(optimized_open_queue)
            else:
                open_queue = pd.DataFrame(columns=df_m.columns)
            
            # Recombine
            combined = pd.concat([locked_queue, urgent_queue, open_queue])
            
            # Record the new sequence
            for i, row in enumerate(combined.itertuples()):
                # Ensure LOT is string to match DB perfectly without '.0'
                lot_str = str(row.LOT).split('.')[0].strip()
                updates.append((i, lot_str, str(mac)))
                
        # 4. Update the database
        cursor = conn.cursor()
        # Ensure ai_sıra column exists
        cursor.execute("PRAGMA table_info(programs)")
        cols = [c[1] for c in cursor.fetchall()]
        if "ai_sıra" not in cols:
            cursor.execute("ALTER TABLE programs ADD COLUMN ai_sıra INTEGER")
            
        cursor.executemany("UPDATE programs SET ai_sıra = ? WHERE LOT = ? AND MAKİNE = ?", updates)
        conn.commit()
        conn.close()
        
        return True, "Planlama başarıyla tamamlandı!"
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, f"Hata oluştu: {str(e)}"
