#!/usr/bin/env python3
"""Test parser for EmStat Pico data"""

import csv

def parse_emstat_value(value_str):
    """Parse EmStat Pico hexadecimal value with optional unit suffix"""
    if not value_str:
        return 0
        
    value_str = value_str.strip()
    
    # Separate hex value from unit suffix
    hex_part = value_str
    unit_suffix = ''
    
    # Check if last character is a unit suffix (letter)
    if value_str and value_str[-1].isalpha():
        unit_suffix = value_str[-1]
        hex_part = value_str[:-1]
    
    try:
        # Parse hexadecimal value
        raw_value = int(hex_part, 16) if hex_part else 0
        
        # Handle signed 32-bit values (two's complement)
        # 0x80000000 is the midpoint (zero)
        if raw_value >= 0x80000000:
            # Convert to signed by subtracting the offset
            signed_value = raw_value - 0x80000000
        else:
            # Values below midpoint are negative
            signed_value = raw_value - 0x80000000
        
        # Apply unit scaling based on suffix
        # These scaling factors may need adjustment based on your device settings
        if unit_suffix == 'n':  # nano range
            # Value is in small units, likely nA for current
            return float(signed_value) / 1000.0  # Adjust scaling as needed
        elif unit_suffix == 'u':  # micro range  
            # Value is in micro units
            return float(signed_value) / 100.0  # Adjust scaling as needed
        elif unit_suffix in ['a', 'f']:  # Other ranges
            # These might be auto-range indicators
            return float(signed_value) / 10.0  # Adjust scaling as needed
        else:
            # No suffix or space - base units
            return float(signed_value)
            
    except (ValueError, TypeError) as e:
        print(f"Error parsing {value_str}: {e}")
        return 0.0

def parse_data_line(line):
    """Parse a single data line from EmStat Pico"""
    try:
        # Parse EmStat Pico format: Pda[hex][unit];ba[hex][unit],14,xxx
        if 'Pda' in line and 'ba' in line:
            # Split by semicolon
            parts = line.split(';')
            if len(parts) >= 2:
                data_point = {}
                
                # Parse potential (Pda prefix)
                pda_part = parts[0].strip()
                if pda_part.startswith('Pda'):
                    pda_value = pda_part[3:]  # Remove 'Pda' prefix
                    potential_raw = parse_emstat_value(pda_value)
                    # Convert to V (assuming raw is in mV-like units)
                    data_point['potential'] = potential_raw / 1000.0
                
                # Parse current (ba prefix)  
                ba_section = parts[1].split(',')[0].strip()
                if ba_section.startswith('ba'):
                    ba_value = ba_section[2:]  # Remove 'ba' prefix
                    current_raw = parse_emstat_value(ba_value)
                    # Convert to µA (assuming raw is in nA-like units)
                    data_point['current'] = current_raw / 1000.0
                
                return data_point
    except Exception as e:
        print(f"Error parsing line: {e}")
        return None
    
    return None

# Test with your actual data
test_data = """Pda8000000 ;ba57F2238f,14,208
Pda7643EAEn;ba654C17Cf,14,208
Pda6C87D5Cn;ba654C17Cf,14,208
Pda62CBC0An;ba72A60BEf,14,207
Pda590FAB8n;ba72A60BEf,14,207
Pda4F53964n;ba72A60BEf,14,207
Pda4597814n;ba795305Ff,14,206
Pda3BDB6C0n;ba795305Ff,14,206
Pda321F570n;ba795305Ff,14,206
Pda2863418n;ba7CA9830f,14,205
Pda7FE7145u;ba7CA9830f,14,205
Pda7FE4965u;ba7CA9830f,14,205
Pda7FE2186u;ba7E54C18f,14,204"""

print("Testing EmStat Pico data parser...")
print("-" * 60)

data_points = []
for line in test_data.split('\n'):
    if line.strip():
        result = parse_data_line(line.strip())
        if result:
            data_points.append(result)
            print(f"Line: {line[:40]}")
            print(f"  Potential: {result['potential']:.6f} V")
            print(f"  Current: {result['current']:.6f} µA")
            print()

print(f"Total parsed data points: {len(data_points)}")

# Save to CSV for verification
if data_points:
    with open('test_parsed_data.csv', 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['potential', 'current'])
        writer.writerow({'potential': 'Potential (V)', 'current': 'Current (µA)'})
        writer.writerows(data_points)
    print("Data saved to test_parsed_data.csv")

    # Show data range
    potentials = [d['potential'] for d in data_points]
    currents = [d['current'] for d in data_points]
    print(f"\nPotential range: {min(potentials):.3f} to {max(potentials):.3f} V")
    print(f"Current range: {min(currents):.3f} to {max(currents):.3f} µA")