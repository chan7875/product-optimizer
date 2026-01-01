import pandas as pd
import argparse
import os
import sys
import re

def parse_arguments():
    parser = argparse.ArgumentParser(description='Calculate production schedule from Excel.')
    parser.add_argument('--file', type=str, required=True, help='Path to the schedule Excel file.')
    parser.add_argument('--date', type=str, required=True, help='Target date (YYYY-MM-DD) to read quantities from.')
    parser.add_argument('--setup-times', type=str, default="", help='Setup times mapping (e.g. S01:40,S02:13)')
    return parser.parse_args()

def calculate_time_for_row(row, target_qty, setup_time_map, line_col_idx):
    """
    Calculates production time based on formula and line-specific setup time.
    """
    item_code = str(row.iloc[0]).strip() # Col A
    layer_info = str(row.iloc[2]).strip() # Col C
    tt_info = str(row.iloc[5]).strip()    # Col F
    array_val = row.iloc[8]               # Col I
    
    # Get Line Info
    line_val = "Unknown"
    if line_col_idx is not None and line_col_idx < len(row):
        val = str(row.iloc[line_col_idx]).strip()
        if val and val != 'nan':
            line_val = val
            
    # Determine Setup Time
    # Default 13
    setup_time = 13
    if line_val in setup_time_map:
        setup_time = setup_time_map[line_val]
    
    # Validation
    if pd.isna(item_code) or item_code == 'nan' or not item_code:
        return []
    
    try:
        # Extract number from string like "(적층) 1"
        s_val = str(array_val)
        match = re.search(r"(\d+(\.\d+)?)", s_val)
        if match:
            array_count = float(match.group(1))
        else:
            return []
    except (ValueError, TypeError):
        return []
        
    try:
        qty = float(target_qty)
        if qty <= 0:
            return []
    except (ValueError, TypeError):
        return []

    # Parse Layer
    layers = []
    if '/' in layer_info:
        layers = [L.strip().upper() for L in layer_info.split('/')]
    else:
        layers = [layer_info.strip().upper()]
        
    std_layers = []
    for L in layers:
        if L == 'B': std_layers.append('Bottom')
        elif L == 'T': std_layers.append('Top')
        else: std_layers.append(L)
        
    # Parse T/T (Cycle Time)
    # Rule: "Left, Right" -> Left=Bottom, Right=Top
    b_cycle = 0.0
    t_cycle = 0.0
    
    # Clean and split
    tt_clean = tt_info.replace(' ', '')
    
    if ',' in tt_info:
        parts = tt_info.split(',')
        try:
            b_cycle = float(parts[0].strip())
        except: b_cycle = 0.0
        
        if len(parts) > 1:
            try:
                t_cycle = float(parts[1].strip())
            except: t_cycle = 0.0
        else:
            t_cycle = b_cycle # Should not happen if ',' exists but just in case
            
    else:
        # Single value case
        try:
            val = float(tt_info.strip())
            b_cycle = val
            t_cycle = val
        except:
            pass # 0.0
            
    if b_cycle == 0 and t_cycle == 0:
        return []

    results = []
    
    def add_result(l_name, c_time):
        if c_time <= 0: return # Skip if no time
        
        # Formula: (TT * Array * Qty) / 60 + SetupTime
        prod_time_mins = ((c_time * array_count * qty) / 60) + setup_time
        results.append({
            'Item_Code': item_code,
            'Layer': l_name,
            'Qty': int(qty),
            'Cycle_Time': c_time,
            'Line': line_val,
            'Setup_Time': setup_time,
            'Prod_Time': round(prod_time_mins, 2)
        })

    for layer_name in std_layers:
        if layer_name == 'Bottom':
            add_result(layer_name, b_cycle)
        elif layer_name == 'Top':
            add_result(layer_name, t_cycle)
        else:
            # Unknown layer? Use B cycle or T?
            # Fallback to B (first val)
            add_result(layer_name, b_cycle)
            
    return results

def main():
    args = parse_arguments()
    
    if not os.path.exists(args.file):
        print(f"Error: File not found: {args.file}")
        sys.exit(1)
        
    print(f"Loading schedule from {args.file}...")
    
    try:
        # Load Excel, header is at row 37 (index 36)
        # Verify if row 37 is indeed the header or if we just start reading data from there?
        # User: "엑셀 파일의 37번행에 ... 정보가 입력되어 있고" -> This usually means Row 37 is the Header.
        df = pd.read_excel(args.file, header=36) 
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Clean columns check
    # We need to find the column for the specific date.
    # The user says "J열부터는 날짜별...". J is index 9 (0-based A=0).
    
    target_col = None
    
    # Look for the date column
    # The header might comprise dates. We need to match string representation.
    # Check if args.date is in columns
    
    # Convert args.date to datetime if columns are datetime objects?
    # Or simple string match.
    
    found_col = False
    for col in df.columns:
        col_str = str(col).split(' ')[0] # Handle timestamps "2024-01-01 00:00:00"
        if args.date in col_str:
            target_col = col
            found_col = True
            break
            
    if not found_col:
        print(f"Error: Could not find column for date {args.date} in Excel header.")
        print("Available columns (sample):", list(df.columns)[9:15]) # Show some date columns
        sys.exit(1)
        
    print(f"Found target date column: {target_col}")

    # Parse Setup Times
    setup_time_map = {}
    if args.setup_times:
        try:
            # Format: S01:40,S02:13
            pairs = args.setup_times.split(',')
            for p in pairs:
                k, v = p.split(':')
                setup_time_map[k.strip()] = float(v.strip())
            print(f"Setup Times: {setup_time_map}")
        except Exception as e:
            print(f"Error parsing setup times: {e}")

    # Find Line Column
    line_col_idx = None
    for i, col in enumerate(df.columns):
        c_str = str(col).lower()
        if "line" in c_str or "생산라인" in c_str:
            line_col_idx = i
            break
            
    print(f"Found Line Column Index: {line_col_idx} (Name: {df.columns[line_col_idx] if line_col_idx is not None else 'None'})")

    all_production_items = []
    
    # Iterate rows
    # df row indices start from 0 (which is Excel row 38 data, since 37 was header)
    for index, row in df.iterrows():
        # Get Qty from target column
        try:
            qty_val = row[target_col]
            if pd.isna(qty_val) or qty_val == 0:
                continue
        except KeyError:
            continue
            
        # Parse and Calculate
        items = calculate_time_for_row(row, qty_val, setup_time_map, line_col_idx)
        all_production_items.extend(items)
        
    # Detailed Output
    print("\n" + "=" * 100)
    print(f"{'Line':<6} | {'Item Code':<20} | {'Layer':<8} | {'Qty':<8} | {'T/T':<8} | {'Setup':<6} | {'Time (min)':<10}")
    print("-" * 100)
    for item in all_production_items:
        print(f"{item['Line']:<6} | {item['Item_Code']:<20} | {item['Layer']:<8} | {item['Qty']:<8} | {item['Cycle_Time']:<8} | {item['Setup_Time']:<6} | {item['Prod_Time']:.2f}")
    print("=" * 100 + "\n")

    # Summary Output
    total_time_mins = sum(item['Prod_Time'] for item in all_production_items)
    operation_rate = (total_time_mins / 480) * 100
    
    # Group by Line
    line_totals = {}
    for item in all_production_items:
        ln = item['Line']
        if ln not in line_totals:
            line_totals[ln] = 0
        line_totals[ln] += item['Prod_Time']
    
    print("-" * 50)
    print(f"Date: {args.date}")
    print(f"Total Production Count (Items): {len(all_production_items)}")
    print(f"Total Production Time: {total_time_mins:.0f} minutes")
    print(f"Operation Rate (vs 480min): {operation_rate:.1f}%")
    print("-" * 20)
    print("Time per Line:")
    for ln, t_min in line_totals.items():
        print(f"  {ln}: {t_min:.0f} min ({(t_min/480)*100:.1f}%)")
    print("-" * 50)
    
    # Save to CSV (item_list_from_excel.txt)
    # Format: Item_Code,T_B,Qty,Prod_Time
    # Note: T_B in item_list.txt expected 'T' or 'B'. 
    # My logic produced 'Top' or 'Bottom'. I should output 'T' or 'B' for compatibility.
    
    output_path = os.path.join(os.path.dirname(args.file), "item_list_from_excel.txt")
    
    try:
        with open(output_path, 'w', encoding='utf-8-sig') as f:
            f.write("Item_Code,T_B,Qty,Prod_Time\n")
            for item in all_production_items:
                # Map Layer back to T/B char
                tb_char = 'T' if item['Layer'] == 'Top' else 'B'
                f.write(f"{item['Item_Code']},{tb_char},{item['Qty']},{item['Prod_Time']}\n")
        
        print(f"Saved production plan to: {output_path}")
    except Exception as e:
        print(f"Error saving output file: {e}")

if __name__ == "__main__":
    main()
