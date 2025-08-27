import json
import os
import glob
import re

def update_delivery_time(json_file_path):
    """Update delivery time in a JSON file based on supplier name."""
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Get supplier name
        supplier_value = data.get('cells', {}).get('B3', {}).get('value', '')
        
        # Get current delivery time
        current_delivery = data.get('cells', {}).get('B14', {}).get('value', '')
        
        # Determine new delivery time based on supplier
        new_days = 15  # default
        
        if '印刷厂' in supplier_value:
            new_days = 7
        elif supplier_value in [
            '宁波泰丰机械有限公司',
            '阳江骏业工贸有限公司', 
            '宁波瑾秀制刷科技有限公司',
            '宁波市海曙硕丰塑料五金制品有限公司'
        ]:
            new_days = 45
        
        # Update the delivery time value with just the number
        if 'B14' in data.get('cells', {}):
            data['cells']['B14']['value'] = str(new_days)
            
            # Save the updated file
            with open(json_file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            print(f"Updated {os.path.basename(json_file_path)}: Supplier: {supplier_value} -> Delivery time: {new_days} days")
            return True
        else:
            print(f"No B14 cell found in {os.path.basename(json_file_path)}")
            return False
            
    except Exception as e:
        print(f"Error processing {json_file_path}: {e}")
        return False

def main():
    # Get all JSON files in the json_template directory
    json_pattern = r"c:\Users\Cheng\Desktop\amazon_order\order_generation\json_template\*.json"
    json_files = glob.glob(json_pattern)
    
    print(f"Found {len(json_files)} JSON files to process...")
    
    updated_count = 0
    for json_file in json_files:
        if update_delivery_time(json_file):
            updated_count += 1
    
    print(f"\nCompleted! Updated {updated_count} out of {len(json_files)} files.")

if __name__ == "__main__":
    main()
