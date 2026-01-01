import pandas as pd
import argparse
import os
import sys

def parse_arguments():
    parser = argparse.ArgumentParser(description='Calculate production schedule from Excel.')
    parser.add_argument('--file', type=str, required=True, help='Path to the schedule Excel file.')
    parser.add_argument('--date', type=str, required=True, help='Target date (YYYY-MM-DD) to read quantities from.')
    return parser.parse_args()

def calculate_time_for_row(row, target_qty):
    """
    Calculates production time for a single row based on the formula:
    Time = (Layer_TT * Array * Qty) / 60
    Returns a list of dicts: [{'Item_Code': ..., 'Layer': ..., 'Prod_Time': ..., 'Qty': ...}]
    """
    item_code = str(row.iloc[0]).strip() # Col A: Item Code
    layer_info = str(row.iloc[2]).strip() # Col C: Layer (B, T, B/T, T/B)
    tt_info = str(row.iloc[5]).strip()    # Col F: T/T (e.g. "42", "42,55")
    array_val = row.iloc[8]               # Col I: Array
    
    # Validation
    if pd.isna(item_code) or item_code == 'nan' or not item_code:
        return []
    
    try:
        array_count = float(array_val)
    except (ValueError, TypeError):
        return []
        
    try:
        qty = float(target_qty)
        if qty <= 0:
            return []
    except (ValueError, TypeError):
        return []

    # Parse Layer and TT
    # Example: Layer="B/T", TT="42,55" -> Bottom=42, Top=55
    # Example: Layer="B", TT="42" -> Bottom=42
    
    layers = []
    if '/' in layer_info:
        layers = [L.strip().upper() for L in layer_info.split('/')]
    else:
        layers = [layer_info.strip().upper()]
        
    # Standardize Layer Names
    std_layers = []
    for L in layers:
        if L == 'B': std_layers.append('Bottom')
        elif L == 'T': std_layers.append('Top')
        else: std_layers.append(L) # Should probably be Error or keep as is? User said B or T.
        
    tts = []
    # TT might be comma separated or just space separated? User said "42,55"
    # User said "F 열에는 T/T 정보가 있는데... 셀에 42,55 가 있으면"
    if ',' in tt_info:
        tts = [float(t.strip()) for t in tt_info.split(',')]
    else:
        try:
            tts = [float(tt_info)]
        except ValueError:
            pass # Handle empty or invalid TT
            
    if not tts:
        return []

    results = []
    
    # Mapping Logic
    # Case 1: Simple 1:1 match
    if len(std_layers) == len(tts):
        for i, layer_name in enumerate(std_layers):
            cycle_time = tts[i]
            # Formula: (TT * Array * Qty) / 60
            # Result is in minutes. Add 13 minutes setup time per layer.
            prod_time_mins = ((cycle_time * array_count * qty) / 60) + 13
            
            results.append({
                'Item_Code': item_code,
                'Layer': layer_name,
                'Qty': int(qty),
                'Cycle_Time': cycle_time,
                'Prod_Time': round(prod_time_mins, 2)
            })
            
    # Case 2: Mismatch (Robustness)
    # If 2 layers but 1 TT? Assume same TT? Or Error?
    # Taking safe approach: if 1 TT provided but 2 layers, apply to both?
    # Or strict mapping. Let's try strict mapping first, but fallback to 1st TT if mismatch.
    elif len(tts) == 1 and len(std_layers) > 1:
         for layer_name in std_layers:
            cycle_time = tts[0]
            prod_time_mins = ((cycle_time * array_count * qty) / 60) + 13
            results.append({
                'Item_Code': item_code,
                'Layer': layer_name,
                'Qty': int(qty),
                'Cycle_Time': cycle_time,
                'Prod_Time': round(prod_time_mins, 2)
            })
            
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
        items = calculate_time_for_row(row, qty_val)
        all_production_items.extend(items)
        
    # Detailed Output
    print("\n" + "=" * 80)
    print(f"{'Item Code':<20} | {'Layer':<8} | {'Qty':<8} | {'T/T':<8} | {'Time (min)':<10}")
    print("-" * 80)
    for item in all_production_items:
        print(f"{item['Item_Code']:<20} | {item['Layer']:<8} | {item['Qty']:<8} | {item['Cycle_Time']:<8} | {item['Prod_Time']:.2f}")
    print("=" * 80 + "\n")

    # Summary Output
    total_time_mins = sum(item['Prod_Time'] for item in all_production_items)
    operation_rate = (total_time_mins / 480) * 100
    
    print("-" * 50)
    print(f"Date: {args.date}")
    print(f"Total Production Count (Items): {len(all_production_items)}")
    print(f"Total Production Time: {total_time_mins:.0f} minutes")
    print(f"Operation Rate (vs 480min): {operation_rate:.1f}%")
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
