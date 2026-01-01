import csv
import os

import argparse

import csv
import os
import argparse

import csv
import os
import argparse

def main():
    parser = argparse.ArgumentParser(description='Analyze BOM data.')
    parser.add_argument('--extract-common', action='store_true', help='Extract common parts used by ALL items.')
    parser.add_argument('--items', type=str, help='Comma-separated list of specific Item_Codes to analyze for common parts.')
    args = parser.parse_args()

    # File Paths
    base_dir = "."
    bom_path = os.path.join(base_dir, "Input", "BOM.txt")
    common_list_path = os.path.join(base_dir, "Input", "common_material_list.csv")
    output_path = os.path.join(base_dir, "Output", "optimization_result.csv")
    common_extract_path = os.path.join(base_dir, "Output", "common_part.csv")

    # --- Mode 1: Extract Common Parts (Dynamic Intersection) ---
    if args.extract_common:
        print("Extracting common parts (Dynamic Intersection)...")
        print("Loading BOM data...")
        item_materials = {} # Item_Code -> Set of Material_Code
        
        try:
            with open(bom_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    item_code = row.get('Item_Code', '').strip()
                    material_code = row.get('Material_Code', '').strip()
                    
                    if item_code and material_code:
                        if item_code not in item_materials:
                            item_materials[item_code] = set()
                        item_materials[item_code].add(material_code)
            
            if not item_materials:
                print("No data found in BOM.")
                return

            # Filter by specific items if requested
            if args.items:
                target_items = [x.strip() for x in args.items.split(',') if x.strip()]
                print(f"Filtering for items: {target_items}")
                
                filtered_materials = {k: v for k, v in item_materials.items() if k in target_items}
                
                if not filtered_materials:
                    print(f"Error: None of the requested items {target_items} were found in the BOM.")
                    return
                
                # Check if all requested items were found
                found_keys = list(filtered_materials.keys())
                if len(found_keys) < len(target_items):
                    missing = set(target_items) - set(found_keys)
                    print(f"Warning: The following items were not found in BOM: {missing}")
                
                item_materials = filtered_materials

            # Calculate Intersection across selected items
            first_item = next(iter(item_materials))
            common_set = item_materials[first_item].copy()
            
            for item, mats in item_materials.items():
                common_set.intersection_update(mats)
            
            print(f"Found {len(common_set)} common parts across {len(item_materials)} items: {list(item_materials.keys())}")
            
            # Save to CSV
            os.makedirs(os.path.dirname(common_extract_path), exist_ok=True)
            with open(common_extract_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['Common_Material_Code'])
                for mat in sorted(common_set):
                    writer.writerow([mat])
            
            print(f"Successfully saved common parts to: {common_extract_path}")
            
        except Exception as e:
            print(f"Error processing BOM for common parts: {e}")
        
        return

    # --- Mode 2: Standard Analysis (Using common_material_list.csv) ---
    print("Loading data for analysis...")

    # 1. Load Common Materials from File
    common_materials = set()
    try:
        if os.path.exists(common_list_path):
            with open(common_list_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if 'Material_Code' in row and row['Material_Code']:
                        common_materials.add(row['Material_Code'].strip())
            print(f"Loaded {len(common_materials)} common materials from {common_list_path}.")
        else:
            print(f"Warning: Common material file not found: {common_list_path}. Treating all materials as individual.")
    except Exception as e:
        print(f"Error reading common material list: {e}")
        return

    # 2. Load BOM Data
    item_layer_materials = {} # (Item_Code, Layer) -> Set of Material_Code
    
    try:
        with open(bom_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                item_code = row.get('Item_Code', '').strip()
                material_code = row.get('Material_Code', '').strip()
                t_b_code = row.get('T_B', '').strip().upper()
                
                # Determine Layer
                layer = "Unknown"
                if t_b_code == "SB":
                    layer = "Bottom"
                elif t_b_code == "ST":
                    layer = "Top"
                
                if item_code and material_code:
                    key = (item_code, layer)
                    if key not in item_layer_materials:
                        item_layer_materials[key] = set()
                    item_layer_materials[key].add(material_code)
        
        print(f"Loaded BOM data for {len(item_layer_materials)} item-layer combinations.")
    except Exception as e:
        print(f"Error reading BOM file: {e}")
        return

    # 3. Analyze and Output
    results = []
    print("\nAnalysis Result (Preview):")
    print("-" * 100)
    print(f"{'Item_Code':<20} | {'Layer':<10} | {'Common':<10} | {'Individual':<10}")
    print("-" * 100)

    # Sort keys for consistent output
    sorted_keys = sorted(item_layer_materials.keys())

    for item, layer in sorted_keys:
        materials = item_layer_materials[(item, layer)]
        common_in_item = {m for m in materials if m in common_materials}
        individual_in_item = materials - common_in_item
        
        results.append({
            'Item_Code': item,
            'Layer': layer,
            'Common_Count': len(common_in_item),
            'Individual_Count': len(individual_in_item),
            'Common_Materials': ','.join(sorted(common_in_item)),
            'Individual_Materials': ','.join(sorted(individual_in_item))
        })

        print(f"{item:<20} | {layer:<10} | {len(common_in_item):<10} | {len(individual_in_item):<10}")

    # 4. Save to CSV
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            fieldnames = ['Item_Code', 'Layer', 'Common_Count', 'Individual_Count', 'Common_Materials', 'Individual_Materials']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)
        print(f"\nSuccessfully saved detailed results to: {output_path}")
    except Exception as e:
        print(f"Error saving output file: {e}")

if __name__ == "__main__":
    main()

