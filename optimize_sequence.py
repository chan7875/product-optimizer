import csv
import os
import sys
import argparse
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp

def load_production_data(file_path):
    """Loads production data (Qty, Prod_Time) from item_list.txt."""
    prod_data = {} # (Item_Code, Layer) -> {'Qty': ..., 'Prod_Time': ...}
    
    if not os.path.exists(file_path):
        print(f"Warning: Production data file not found: {file_path}")
        return prod_data

    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            # Read header
            header_line = f.readline()
            if not header_line:
                return prod_data
            headers = [h.strip() for h in header_line.split(',')]
            
            # Map headers to indices
            try:
                idx_item = headers.index('Item_Code')
                idx_tb = headers.index('T_B')
                idx_qty = headers.index('Qty')
                idx_time = headers.index('Prod_Time')
            except ValueError as e:
                print(f"Error parsing headers in item_list.txt: {e}")
                return prod_data

            for line in f:
                if not line.strip():
                    continue
                parts = [p.strip() for p in line.split(',')]
                if len(parts) <= max(idx_item, idx_tb, idx_qty, idx_time):
                    continue
                
                item_code = parts[idx_item]
                tb_val = parts[idx_tb].upper()
                qty = parts[idx_qty]
                prod_time = parts[idx_time]
                
                layer = "Unknown"
                if tb_val == 'B':
                    layer = 'Bottom'
                elif tb_val == 'T':
                    layer = 'Top'
                
                if item_code and layer != "Unknown":
                    prod_data[(item_code, layer)] = {
                        'Qty': qty,
                        'Prod_Time': prod_time
                    }
        print(f"Loaded production data for {len(prod_data)} items.")
    except Exception as e:
        print(f"Error reading item_list.txt: {e}")
    
    return prod_data

def parse_arguments():
    parser = argparse.ArgumentParser(description='Optimize production sequence.')
    parser.add_argument('--priority', type=str, help='Comma-separated list of Item_Codes to prioritize.')
    parser.add_argument('--manual', type=str, help='Manual sequence definition (e.g. "(ItemA,Top), (ItemB,Bottom)")')
    parser.add_argument('--layer', type=str, choices=['TB', 'BT'], help='Prioritize specific layer order: TB (Top then Bottom) or BT (Bottom then Top).')
    return parser.parse_args()

def parse_manual_sequence(manual_str):
    """
    Parses string like "(EP94-04976A,Top), (EP94-04820A,Top)"
    Returns list of tuples: [('EP94-04976A', 'Top'), ('EP94-04820A', 'Top')]
    """
    items = []
    # Remove parens and split
    raw_str = manual_str.replace("(", "").replace(")", "")
    parts = raw_str.split(',')
    
    # Filter out empty strings
    parts = [p.strip() for p in parts if p.strip()]
    
    # parts should be [Item1, Layer1, Item2, Layer2, ...]
    if len(parts) % 2 != 0:
        print("Error: Manual sequence format must be pairs of (Item, Layer).")
        return []
    
    for i in range(0, len(parts), 2):
        item_code = parts[i].strip()
        layer = parts[i+1].strip()
        items.append((item_code, layer))
        
    return items

def solve_tsp(jobs, start_ref_job=None):
    """
    Solves TSP for a list of jobs.
    """
    if not jobs:
        return []
    
    if len(jobs) == 1:
        return [0]

    # Calculate distance matrix (symmetric difference)
    num_jobs = len(jobs)
    # 0 is the depot (dummy or start ref), 1..n are actual jobs
    num_nodes = num_jobs + 1 
    distance_matrix = [[0] * num_nodes for _ in range(num_nodes)]

    for i in range(num_jobs):
        # Distance from Depot to Job i
        if start_ref_job:
            cost = len(start_ref_job['Individual_Set'].symmetric_difference(jobs[i]['Individual_Set']))
            distance_matrix[0][i+1] = cost
        else:
            distance_matrix[0][i+1] = 0
            
        # Distance from Job i to Depot
        distance_matrix[i+1][0] = 0

        for j in range(num_jobs):
            if i == j:
                distance_matrix[i+1][j+1] = 0
            else:
                s1 = jobs[i]['Individual_Set']
                s2 = jobs[j]['Individual_Set']
                cost = len(s1.symmetric_difference(s2))
                distance_matrix[i+1][j+1] = cost

    # Create Data Model
    data = {}
    data['distance_matrix'] = distance_matrix
    data['num_vehicles'] = 1
    data['depot'] = 0

    manager = pywrapcp.RoutingIndexManager(len(data['distance_matrix']),
                                           data['num_vehicles'], data['depot'])
    routing = pywrapcp.RoutingModel(manager)

    def distance_callback(from_index, to_index):
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return data['distance_matrix'][from_node][to_node]

    transit_callback_index = routing.RegisterTransitCallback(distance_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (
        routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)
    search_parameters.time_limit.seconds = 5

    solution = routing.SolveWithParameters(search_parameters)
    
    if solution:
        # Extract solution (indices in 'jobs' list)
        ordered_indices = []
        index = routing.Start(0)
        while not routing.IsEnd(index):
            node_index = manager.IndexToNode(index)
            if node_index != 0:
                ordered_indices.append(node_index - 1)
            index = solution.Value(routing.NextVar(index))
            
        return ordered_indices
    else:
        print("No solution found for subset.")
        return list(range(len(jobs))) # Fallback

def optimize_segment(jobs, layer_mode, start_ref_job=None):
    """
    Optimizes a segment of jobs, optionally splitting by layer.
    Returns: List of ordered jobs.
    """
    if not jobs:
        return []

    if not layer_mode:
        # No layer priority, optimize as one block
        indices = solve_tsp(jobs, start_ref_job)
        return [jobs[i] for i in indices]
    
    # Split by Layer
    tops = [j for j in jobs if j['Layer'] == 'Top']
    bottoms = [j for j in jobs if j['Layer'] == 'Bottom']
    others = [j for j in jobs if j['Layer'] not in ['Top', 'Bottom']] # Should be empty ideally
    
    ordered_segment = []
    current_ref = start_ref_job
    
    first_group = []
    second_group = []
    
    if layer_mode == 'TB':
        first_group = tops
        second_group = bottoms
    elif layer_mode == 'BT':
        first_group = bottoms
        second_group = tops
    
    # 1. Optimize First Group
    if first_group:
        indices = solve_tsp(first_group, current_ref)
        ordered_first = [first_group[i] for i in indices]
        ordered_segment.extend(ordered_first)
        if ordered_first:
            current_ref = ordered_first[-1]
            
    # 2. Optimize Second Group
    if second_group:
        indices = solve_tsp(second_group, current_ref)
        ordered_second = [second_group[i] for i in indices]
        ordered_segment.extend(ordered_second)
        if ordered_second:
            current_ref = ordered_second[-1]
            
    # 3. Optimize Others (if any, append at end)
    if others:
        indices = solve_tsp(others, current_ref)
        ordered_others = [others[i] for i in indices]
        ordered_segment.extend(ordered_others)
        
    return ordered_segment

def load_common_materials(file_path):
    """Loads common materials from csv file into a set."""
    common_set = set()
    if not os.path.exists(file_path):
        print(f"Warning: Common material file not found: {file_path}")
        return common_set
    
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            for row in reader:
                if row:
                    # Assuming strictly one column or first column is the ID
                    common_set.add(row[0].strip())
        print(f"Loaded {len(common_set)} common materials from file.")
    except Exception as e:
        print(f"Error reading common material file: {e}")
        
    return common_set

def main():
    args = parse_arguments()
    
    base_dir = r"D:\Develoment\ProductOptimize"
    input_path = os.path.join(base_dir, "Output", "optimization_result.csv")
    item_list_path = os.path.join(base_dir, "Input", "item_list.txt")
    common_mat_path = os.path.join(base_dir, "Input", "common_material_list.csv")
    output_path = os.path.join(base_dir, "Output", "optimization_sequence.csv")

    if not os.path.exists(input_path):
        print(f"Error: Input file not found: {input_path}")
        return

    # 0. Load Common Materials (to ensure exclusion)
    common_materials_set = load_common_materials(common_mat_path)

    # 1. Load Jobs
    jobs = []
    try:
        with open(input_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                ind_str = row.get('Individual_Materials', '')
                if ind_str:
                    ind_set = set(x.strip() for x in ind_str.replace('"', '').split(',') if x.strip())
                else:
                    ind_set = set()
                
                # Explicitly remove common materials if they exist in valid set
                if common_materials_set:
                    ind_set = ind_set.difference(common_materials_set)
                
                row['Individual_Set'] = ind_set
                jobs.append(row)
        print(f"Loaded {len(jobs)} jobs from optimization result.")
    except Exception as e:
        print(f"Error reading input file: {e}")
        return

    if not jobs:
        print("No jobs to process.")
        return

    # 1.5 Load and Merge Production Data
    prod_data = load_production_data(item_list_path)
    for job in jobs:
        key = (job.get('Item_Code'), job.get('Layer'))
        if key in prod_data:
            job.update(prod_data[key])
        else:
            job['Qty'] = ''
            job['Prod_Time'] = ''

    # Define final sequence
    final_sequence = []

    # 2. Logic Branch
    if args.manual:
        print(f"Processing Manual Sequence: {args.manual}")
        manual_keys = parse_manual_sequence(args.manual)
        jobs_map = {(j['Item_Code'], j['Layer']): j for j in jobs}
        
        for key in manual_keys:
            if key in jobs_map:
                job = jobs_map[key]
                job['Is_Manual'] = True 
                final_sequence.append(job)
            else:
                print(f"Warning: Manual Item {key} not found in loaded data. Skipping.")
        
        if not final_sequence:
            print("Error: No valid jobs found in manual sequence.")
            return

    else:
        # Optimization Mode (Priority + Layer)
        
        # Split by Priority Item Codes
        priority_codes = []
        if args.priority:
            priority_codes = [x.strip() for x in args.priority.split(',') if x.strip()]
            print(f"Priority Items: {priority_codes}")
        
        prio_jobs = []
        remaining_jobs = []
        
        for job in jobs:
            if job.get('Item_Code') in priority_codes:
                prio_jobs.append(job)
            else:
                remaining_jobs.append(job)
        
        current_ref_job = None
        
        # Stage 1: Priority Jobs
        if prio_jobs:
            print(f"Optimizing {len(prio_jobs)} priority jobs (LayerMode: {args.layer})...")
            # Optimize segment (handles layer split internally)
            ordered_prio = optimize_segment(prio_jobs, args.layer, start_ref_job=None)
            
            for job in ordered_prio:
                job['Is_Priority'] = True
                final_sequence.append(job)
            
            if final_sequence:
                current_ref_job = final_sequence[-1]
        
        # Stage 2: Remaining Jobs
        if remaining_jobs:
            print(f"Optimizing {len(remaining_jobs)} remaining jobs (LayerMode: {args.layer})...")
            ordered_rem = optimize_segment(remaining_jobs, args.layer, start_ref_job=current_ref_job)
            
            for job in ordered_rem:
                job['Is_Priority'] = False
                final_sequence.append(job)

    # 3. Save Result with Reasoning
    try:
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            base_fields = [k for k in jobs[0].keys() if k != 'Individual_Set' and k != 'Index' and k != 'Transition_Shared_Count' and k != 'Selection_Reason' and k != 'Is_Priority' and k != 'Is_Manual' and k != 'Total_Count'] 
            
            ordered_base_fields = []
            priority_fields_list = ['Item_Code', 'Layer', 'Qty', 'Prod_Time', 'Total_Count', 'Common_Count', 'Individual_Count', 'Transition_Shared_Count', 'Selection_Reason']
            
            for pf in priority_fields_list:
                if pf in base_fields:
                    ordered_base_fields.append(pf)
                    base_fields.remove(pf)
                else:
                    # If it's a new field not in jobs[0] (like Total_Count), we just add it to our columns list
                    # But we need to make sure it's in fieldnames
                    ordered_base_fields.append(pf)

            ordered_base_fields.extend(base_fields)

            fieldnames = ['Index'] + ordered_base_fields
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            
            for i, current_job in enumerate(final_sequence):
                # Calculate stats
                try:
                    c_count = int(current_job.get('Common_Count', 0))
                    i_count = int(current_job.get('Individual_Count', 0))
                    total_count = c_count + i_count
                except ValueError:
                    c_count = 0
                    i_count = 0
                    total_count = 0
                shared_count = 0
                reason = ""
                
                prev_job = final_sequence[i-1] if i > 0 else None
                
                if i == 0:
                    if current_job.get('Is_Manual'):
                        reason = "사용자 지정 수동 순서 (Manual Sequence)"
                    elif current_job.get('Is_Priority'):
                        reason = "사용자 요청에 의한 우선 생산 모델 (Priority)"
                    else:
                        reason = "전체 생산 일정의 자재 교체 비용을 최소화하기 위한 최적의 시작점으로 선정됨."
                    shared_count = 0
                else:
                    # Intersection (Shared Count) logic
                    current_set = current_job.get('Individual_Set', set())
                    prev_set = prev_job.get('Individual_Set', set())
                    intersection = current_set.intersection(prev_set)
                    shared_count = len(intersection)
                    
                    prev_item = prev_job.get('Item_Code', 'Unknown')
                    
                    if current_job.get('Is_Manual'):
                         # Logic asked: "Same count calculation"
                         reason = f"이전 생산 모델 ({prev_item})과 개별 자재 {shared_count}개가 동일함 (Manual Evaluation)."
                    elif current_job.get('Is_Priority'):
                         reason = f"사용자 요청에 의한 우선 생산 모델 (Priority). 이전 모델과 개별 자재 {shared_count}개 동일."
                    else:
                         reason = f"이전 생산 모델 ({prev_item})과 개별 자재 {shared_count}개가 동일하여 생산 효율성을 위해 배치함 (전체 자재 {total_count}개, 공통 자재 {c_count}개)."

                # Clean up internal keys
                out_row = {k: v for k, v in current_job.items() if k in fieldnames}
                out_row['Index'] = i + 1
                out_row['Transition_Shared_Count'] = shared_count
                out_row['Selection_Reason'] = reason
                out_row['Total_Count'] = total_count # Calculated in the previous block
                
                writer.writerow(out_row)
                
        print(f"Successfully saved sequenced results to: {output_path}")
    except Exception as e:
         print(f"Error saving output file: {e}")

if __name__ == "__main__":
    main()
