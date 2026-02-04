import pandas as pd
import io
import os
import sys
import re
from typing import List, Dict, Tuple, Optional
from datetime import datetime

# ===============================================================================
# FORMULA DATA AND STATE MAPPING
# ===============================================================================

FORMULA_DATA = [
    {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 30%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 40%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
    {"LOB": "BUS", "SEGMENT": "STAFF BUS", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "PO": "88% of Payin", "REMARKS": "NIL"}
]

STATE_MAPPING = {
    "DELHI": "DELHI", "Mumbai": "MAHARASHTRA", "Pune": "MAHARASHTRA", "Goa": "GOA",
    "Kolkata": "WEST BENGAL", "Hyderabad": "TELANGANA", "Ahmedabad": "GUJARAT",
    "Surat": "GUJARAT", "Jaipur": "RAJASTHAN", "Lucknow": "UTTAR PRADESH",
    "Patna": "BIHAR", "Ranchi": "JHARKHAND", "Bhuvaneshwar": "ODISHA",
    "Srinagar": "JAMMU AND KASHMIR", "Dehradun": "UTTARAKHAND", "Haridwar": "UTTARAKHAND",
    "Bangalore": "KARNATAKA", "Jharkhand": "JHARKHAND", "Bihar": "BIHAR",
    "Good GJ": "GUJARAT", "Bad GJ": "GUJARAT", "ROM1": "REST OF MAHARASHTRA",
    "Good TN": "TAMIL NADU", "Kerala": "KERALA", "Good MP": "MADHYA PRADESH",
    "Good RJ": "RAJASTHAN", "Good UP": "UTTAR PRADESH", "Punjab": "PUNJAB",
    "Jammu": "JAMMU AND KASHMIR", "Assam": "ASSAM", "HR Ref": "HARYANA",
    "ROM2": "REST OF MAHARASHTRA", "Andaman": "ANDAMAN AND NICOBAR ISLANDS"
}

# ===============================================================================
# CORE CALCULATION FUNCTIONS
# ===============================================================================

def safe_float(value) -> Optional[float]:
    """Safely convert value to float, handling various edge cases."""
    if pd.isna(value):
        return None
    s = str(value).strip().upper()
    if s in ["D", "NA", "", "NAN", "NONE"]:
        return None
    try:
        num = float(s.replace("%", ""))
        return num * 100 if 0 < num < 1 else num
    except:
        return None


def get_payin_category(payin: float) -> str:
    """Categorize payin percentage into predefined ranges."""
    if payin <= 20:
        return "Payin Below 20%"
    elif payin <= 30:
        return "Payin 21% to 30%"
    elif payin <= 50:
        return "Payin 31% to 50%"
    else:
        return "Payin Above 50%"


def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[str, float]:
    """Get formula and calculate payout based on LOB, segment, policy type, and payin."""
    segment_key = segment.upper()
    
    if lob == "TW":
        segment_key = "TW TP" if policy_type == "TP" else "TW SAOD + COMP"
    elif lob == "PVT CAR":
        segment_key = "PVT CAR TP" if policy_type == "TP" else "PVT CAR COMP + SAOD"
    elif lob in ["TAXI", "CV", "BUS", "MISD"]:
        segment_key = segment.upper()
    
    payin_cat = get_payin_category(payin)
    
    for rule in FORMULA_DATA:
        if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
            if rule["REMARKS"] == payin_cat or rule["REMARKS"] == "NIL":
                formula = rule["PO"]
                if "of Payin" in formula:
                    pct = float(formula.split("%")[0].replace("Less ", ""))
                    payout = round(payin * pct / 100, 2) if "Less" not in formula else round(payin - pct, 2)
                elif formula.startswith("-"):
                    ded = float(formula.replace("%", "").replace("-", ""))
                    payout = round(payin - ded, 2)
                else:
                    payout = round(payin - 2, 2)
                return formula, payout
    
    # Default fallback
    ded = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
    return f"-{ded}%", round(payin - ded, 2)


def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
    """Calculate payout with formula and rule explanation."""
    if payin == 0:
        return 0, "0% (No Payin)", "Payin is 0"
    formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
    return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"

# ===============================================================================
# PATTERN DETECTION
# ===============================================================================

class Pattern4WDetector:
    """Detects whether a 4W sheet is COMP/SAOD or SATP pattern."""
    
    @staticmethod
    def detect_pattern(df: pd.DataFrame) -> str:
        """
        Detect the pattern type based on sheet structure.
        Returns: 'comp_saod' or 'satp'
        """
        # Check if it's a dataframe with headers
        if isinstance(df.columns, pd.Index):
            columns = [str(col).upper().strip() for col in df.columns]
            
            # SATP pattern indicators
            if 'CLUSTER' in columns and 'CD2' in columns:
                # Check for additional SATP-specific columns
                if any(keyword in ' '.join(columns) for keyword in ['NEW SEGMENT', 'AGE BAND', 'MAPPING']):
                    return 'satp'
            
            # Check first few rows for pattern indicators
            first_rows = df.head(10).to_string().upper()
            if 'SATP' in first_rows or 'TP' in first_rows:
                return 'satp'
        
        # Check for COMP/SAOD pattern (header=None format)
        # Look for "Cluster" keyword and CD2 columns in raw data
        df_str = df.head(20).to_string().upper()
        
        if 'CLUSTER' in df_str and 'CD2' in df_str:
            # Check for COMP/SAOD indicators
            if any(keyword in df_str for keyword in ['COMP', 'SAOD', 'PETROL', 'HEV', 'RENEWAL']):
                return 'comp_saod'
        
        # Default to comp_saod
        return 'comp_saod'
    
    @staticmethod
    def detect_pattern_name(df: pd.DataFrame) -> str:
        """Get a descriptive name for the detected pattern."""
        pattern = Pattern4WDetector.detect_pattern(df)
        pattern_names = {
            'comp_saod': "4W COMP/SAOD Pattern (Private Car Comprehensive)",
            'satp': "4W SATP Pattern (Private Car Third Party)"
        }
        return pattern_names.get(pattern, "Unknown 4W Pattern")

# ===============================================================================
# PATTERN PROCESSORS
# ===============================================================================

class CompSaodProcessor:
    """Process COMP/SAOD pattern sheets for Private Car."""
    
    @staticmethod
    def process(content: bytes, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None,
                override_policy_type: str = None) -> List[Dict]:
        """Process COMP/SAOD pattern sheets."""
        records = []
        
        try:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            
            print(f"\n{'='*80}")
            print(f"Processing Sheet: {sheet_name} (COMP/SAOD Pattern)")
            print(f"{'='*80}")
            
            # Find Cluster column
            cluster_col = None
            for j in range(df.shape[1]):
                if df.iloc[:, j].astype(str).str.contains("Cluster", case=False, na=False).any():
                    cluster_col = j
                    break
            
            if cluster_col is None:
                print("❌ ERROR: Could not find 'Cluster' column.")
                return []
            
            print(f"✓ Found Cluster column at index {cluster_col}")
            
            # Find CD2 columns (any column to the right with "CD2" anywhere in the column)
            cd2_cols = []
            for j in range(cluster_col + 1, df.shape[1]):
                col_str = df.iloc[:, j].astype(str).str.cat(sep=' ')
                if "CD2" in col_str.upper():
                    cd2_cols.append(j)
            
            if not cd2_cols:
                print("⚠️  WARNING: No CD2 columns detected.")
                return []
            
            print(f"✓ Found {len(cd2_cols)} CD2 columns")
            
            # Build full header text for each CD2 column
            headers = {}
            cluster_header_row = None
            for i in range(df.shape[0]):
                if pd.notna(df.iloc[i, cluster_col]) and "cluster" in str(df.iloc[i, cluster_col]).lower():
                    cluster_header_row = i
                    break
            
            header_rows_range = range(0, cluster_header_row + 3 if cluster_header_row else 10)
            
            for j in cd2_cols:
                header_parts = []
                for i in header_rows_range:
                    val = df.iloc[i, j]
                    if pd.notna(val):
                        s = str(val).strip()
                        if "CD2" not in s.upper():  # Exclude the CD2 label itself
                            header_parts.append(s)
                headers[j] = " ".join(header_parts).strip()
            
            # Detect data start row
            if cluster_header_row is not None:
                data_start_row = cluster_header_row + 1
                # Skip any fully empty rows after header
                while data_start_row < df.shape[0] and pd.isna(df.iloc[data_start_row, cluster_col]):
                    data_start_row += 1
            else:
                data_start_row = 10  # safe fallback
            
            print(f"✓ Data starts at row {data_start_row + 1}")
            
            # Process data rows
            for i in range(data_start_row, df.shape[0]):
                cluster_cell = df.iloc[i, cluster_col]
                if pd.isna(cluster_cell):
                    continue
                cluster = str(cluster_cell).strip()
                if not cluster or "total" in cluster.lower() or cluster.lower() in ["grand total", "average"]:
                    continue
                
                state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
                
                for j in cd2_cols:
                    payin = safe_float(df.iloc[i, j])
                    if payin is None:
                        continue
                    
                    header_text = headers.get(j, "").upper()
                    
                    # Policy Type detection
                    if "SAOD" in header_text and "COMP" not in header_text:
                        policy_type = "SAOD"
                    elif "COMP" in header_text:
                        policy_type = "COMP"
                    else:
                        policy_type = "COMP"  # default
                    
                    # Fuel detection
                    fuel = "Petrol" if "PETROL" in header_text and "NON" not in header_text and "CNG" not in header_text else "Non-Petrol (incl. CNG)"
                    
                    # HEV vs Non-HEV
                    segment = "Non-HEV" if "NON HEV" in header_text or "NON-HEV" in header_text else "HEV"
                    
                    # Renewal?
                    renewal = " (Renewals)" if "RENEWAL" in header_text or "RENEW" in header_text else ""
                    
                    orig_seg = f"PVT CAR {segment} - {fuel}{renewal}".strip()
                    
                    lob_final = override_lob if override_enabled and override_lob else "PVT CAR"
                    segment_final = override_segment if override_enabled and override_segment else "PVT CAR COMP + SAOD"
                    policy_final = override_policy_type if override_policy_type else policy_type
                    
                    payout, formula, exp = calculate_payout_with_formula(lob_final, segment_final, policy_final, payin)
                    
                    records.append({
                        "State": state,
                        "Location/Cluster": cluster,
                        "Original Segment": orig_seg,
                        "Mapped Segment": segment_final,
                        "LOB": lob_final,
                        "Policy Type": policy_final,
                        "Payin (CD2)": f"{payin:.2f}%",
                        "Payin Category": get_payin_category(payin),
                        "Calculated Payout": f"{payout:.2f}%",
                        "Formula Used": formula,
                        "Rule Explanation": exp
                    })
            
            print(f"✓ Successfully extracted {len(records)} records")
            return records
            
        except Exception as e:
            print(f"❌ ERROR processing COMP/SAOD sheet '{sheet_name}': {e}")
            import traceback
            traceback.print_exc()
            return []


class SatpProcessor:
    """Process SATP (TP) pattern sheets for Private Car."""
    
    @staticmethod
    def process(content: bytes, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None,
                override_policy_type: str = None) -> List[Dict]:
        """Process SATP pattern sheets."""
        records = []
        
        try:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)
            
            print(f"\n{'='*80}")
            print(f"Processing Sheet: {sheet_name} (SATP Pattern)")
            print(f"{'='*80}")
            print(f"Columns: {df.columns.tolist()}")
            
            # Process rows
            for idx, row in df.iterrows():
                cluster = str(row.get('Cluster', '')).strip()
                if not cluster:
                    continue
                
                payin = safe_float(row.get('CD2'))
                if payin is None:
                    continue
                
                # State mapping
                state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
                
                # Get segment info if available
                new_segment = str(row.get('New Segment Mapping', '')).strip()
                age_band = str(row.get('New Age Band', '')).strip()
                
                # Build segment description
                segment_desc = "PVT CAR TP"
                if new_segment and new_segment != 'nan':
                    segment_desc += f" {new_segment}"
                if age_band and age_band != 'nan':
                    segment_desc += f" (Age: {age_band})"
                
                lob_final = override_lob if override_enabled and override_lob else "PVT CAR"
                segment_final = override_segment if override_enabled and override_segment else "PVT CAR TP"
                policy_final = override_policy_type if override_policy_type else "TP"
                
                payout, formula, exp = calculate_payout_with_formula(lob_final, segment_final, policy_final, payin)
                
                records.append({
                    "State": state.upper(),
                    "Location/Cluster": cluster,
                    "Original Segment": segment_desc.strip(),
                    "Mapped Segment": segment_final,
                    "LOB": lob_final,
                    "Policy Type": policy_final,
                    "Payin (CD2)": f"{payin:.2f}%",
                    "Payin Category": get_payin_category(payin),
                    "Calculated Payout": f"{payout:.2f}%",
                    "Formula Used": formula,
                    "Rule Explanation": exp
                })
            
            print(f"✓ Successfully extracted {len(records)} records")
            return records
            
        except Exception as e:
            print(f"❌ ERROR processing SATP sheet '{sheet_name}': {e}")
            import traceback
            traceback.print_exc()
            return []

# ===============================================================================
# PATTERN DISPATCHER
# ===============================================================================

class Pattern4WDispatcher:
    """Main dispatcher that routes to appropriate 4W pattern processor."""
    
    PATTERN_PROCESSORS = {
        'comp_saod': CompSaodProcessor,
        'satp': SatpProcessor
    }
    
    @staticmethod
    def process_sheet(content: bytes, sheet_name: str,
                     override_enabled: bool = False,
                     override_lob: str = None,
                     override_segment: str = None,
                     override_policy_type: str = None) -> List[Dict]:
        """
        Main entry point for processing any 4W sheet.
        Automatically detects pattern and routes to appropriate processor.
        """
        # Load sheet to detect pattern
        try:
            # Try loading with header first (for SATP)
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)
        except:
            # Fallback to no header (for COMP/SAOD)
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
        # Detect pattern
        pattern = Pattern4WDetector.detect_pattern(df)
        pattern_name = Pattern4WDetector.detect_pattern_name(df)
        
        print(f"\n🔍 Pattern Detection: {pattern_name}")
        
        # Get appropriate processor
        processor_class = Pattern4WDispatcher.PATTERN_PROCESSORS.get(pattern, CompSaodProcessor)
        
        # Process the sheet
        records = processor_class.process(
            content, sheet_name, override_enabled, 
            override_lob, override_segment, override_policy_type
        )
        
        return records

# ===============================================================================
# FILE HANDLING AND EXPORT
# ===============================================================================

def get_sheet_names(file_path: str) -> List[str]:
    """Get all sheet names from an Excel file."""
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    except Exception as e:
        print(f"❌ ERROR reading file: {e}")
        return []


def choose_sheet(sheets: List[str]) -> Optional[str]:
    """Display all sheets and let user choose which one to process."""
    print("\n" + "="*80)
    print("📋 Available Sheets in Excel File:")
    print("="*80)
    for i, sheet in enumerate(sheets, 1):
        print(f"{i}. {sheet}")
    print("="*80)
    
    while True:
        try:
            choice = input(f"\nEnter sheet number to process (1-{len(sheets)}) or 'q' to quit: ").strip()
            
            if choice.lower() == 'q':
                return None
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(sheets):
                selected_sheet = sheets[choice_num - 1]
                print(f"\n✅ Selected: {selected_sheet}")
                return selected_sheet
            else:
                print(f"❌ Please enter a number between 1 and {len(sheets)}")
        except ValueError:
            print("❌ Invalid input. Please enter a number or 'q' to quit.")


def export_to_excel(records: List[Dict], file_path: str, sheet_name: str) -> str:
    """Export records to Excel file."""
    if not records:
        print("⚠️  No records to export!")
        return None
    
    try:
        df = pd.DataFrame(records)
        
        # Generate output filename
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{base_name}_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
        df.to_excel(output_file, index=False, sheet_name='Processed')
        print(f"\n✅ Successfully exported {len(records)} records to: {output_file}")
        return output_file
    except Exception as e:
        print(f"❌ ERROR exporting to Excel: {e}")
        return None


def print_summary(records: List[Dict]):
    """Print a summary of processed records."""
    if not records:
        print("\n⚠️  No records to summarize!")
        return
    
    print("\n" + "="*80)
    print("📊 PROCESSING SUMMARY")
    print("="*80)
    print(f"Total Records: {len(records)}")
    
    # Group by state
    states = {}
    for record in records:
        state = record.get("State", "Unknown")
        states[state] = states.get(state, 0) + 1
    
    print(f"\nRecords by State:")
    for state, count in sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]:
        print(f"  {state}: {count}")
    
    # Group by policy type
    policies = {}
    for record in records:
        policy = record.get("Policy Type", "Unknown")
        policies[policy] = policies.get(policy, 0) + 1
    
    print(f"\nRecords by Policy Type:")
    for policy, count in sorted(policies.items()):
        print(f"  {policy}: {count}")
    
    # Calculate average payin
    payins = []
    for record in records:
        payin_str = record.get("Payin (CD2)", "0%")
        try:
            payin = float(payin_str.replace('%', ''))
            if payin > 0:
                payins.append(payin)
        except:
            pass
    
    if payins:
        avg_payin = sum(payins) / len(payins)
        print(f"\nAverage Payin: {avg_payin:.2f}%")
    
    print("="*80 + "\n")


def print_sample_records(records: List[Dict], count: int = 10):
    """Print sample records."""
    print(f"\n📋 Sample Records (first {min(count, len(records))}):")
    print("="*80)
    for i, record in enumerate(records[:count], 1):
        print(f"{i}. {record.get('Location/Cluster', 'N/A')} | "
              f"{record.get('Original Segment', 'N/A')} | "
              f"Policy: {record.get('Policy Type', 'N/A')} | "
              f"Payin: {record.get('Payin (CD2)', 'N/A')} → "
              f"Payout: {record.get('Calculated Payout', 'N/A')}")
    print("="*80 + "\n")

# ===============================================================================
# MAIN PROGRAM
# ===============================================================================

def main():
    """Main program entry point."""
    print("\n" + "="*80)
    print(" " * 15 + "DIGIT 4W PRIVATE CAR UNIFIED PROCESSOR")
    print("="*80)
    print("\nThis tool automatically detects and processes 4W Private Car sheets:")
    print("  • COMP/SAOD Pattern (Comprehensive & Stand Alone Own Damage)")
    print("  • SATP Pattern (Third Party)")
    print("="*80 + "\n")
    
    # Get file path from user
    while True:
        file_path = input("📁 Enter the Excel file path (or 'q' to quit): ").strip()
        
        if file_path.lower() == 'q':
            print("\n👋 Exiting...")
            return
        
        # Remove quotes if present
        file_path = file_path.strip('"').strip("'")
        
        if os.path.exists(file_path):
            break
        else:
            print(f"❌ File not found: {file_path}")
            print("Please check the path and try again.\n")
    
    # Display sheets and let user choose
    sheets = get_sheet_names(file_path)
    if not sheets:
        print("❌ No worksheets found or file is invalid.")
        return
    
    selected_sheet = choose_sheet(sheets)
    
    if not selected_sheet:
        print("\n👋 Exiting...")
        return
    
    # Ask for override options
    print("\n" + "="*80)
    print("Override Options (usually not needed)")
    print("="*80)
    override_choice = input("Enable override? (y/n): ").strip().lower()
    override_enabled = override_choice == 'y'
    
    override_lob = override_segment = override_policy_type = None
    if override_enabled:
        override_lob = input("Enter LOB (e.g. PVT CAR, TW) or press Enter to skip: ").strip().upper() or None
        override_segment = input("Enter Segment (e.g. PVT CAR COMP + SAOD) or press Enter to skip: ").strip().upper() or None
        override_policy_type = input("Enter Policy Type (COMP/SAOD/TP) or press Enter to skip: ").strip().upper() or None
    
    # Read file once
    print(f"\n📖 Loading file...")
    with open(file_path, "rb") as f:
        content = f.read()
    
    # Process the sheet using pattern detection
    print("\n🚀 Starting processing...")
    records = Pattern4WDispatcher.process_sheet(
        content, selected_sheet, override_enabled,
        override_lob, override_segment, override_policy_type
    )
    
    if not records:
        print("\n⚠️  No records extracted! Please check the sheet structure.")
        return
    
    # Print summary
    print_summary(records)
    
    # Print sample records
    print_sample_records(records)
    
    # Export to Excel
    output_file = export_to_excel(records, file_path, selected_sheet)
    
    if output_file:
        print(f"\n✨ Processing complete! Output file: {output_file}")
    else:
        print("\n⚠️  Export failed, but processing completed successfully.")


if __name__ == "__main__":
    main()
