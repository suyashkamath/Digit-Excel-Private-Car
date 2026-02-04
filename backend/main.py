from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import pandas as pd
import io
import os
from typing import List, Dict, Tuple, Optional
from datetime import datetime
import traceback
import tempfile
import shutil

app = FastAPI(title="DIGIT 4W Processor API")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://digit-excel-private-car.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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

# Store uploaded files temporarily
uploaded_files = {}

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
        """Detect the pattern type based on sheet structure."""
        if isinstance(df.columns, pd.Index):
            columns = [str(col).upper().strip() for col in df.columns]
            
            if 'CLUSTER' in columns and 'CD2' in columns:
                if any(keyword in ' '.join(columns) for keyword in ['NEW SEGMENT', 'AGE BAND', 'MAPPING']):
                    return 'satp'
            
            first_rows = df.head(10).to_string().upper()
            if 'SATP' in first_rows or 'TP' in first_rows:
                return 'satp'
        
        df_str = df.head(20).to_string().upper()
        
        if 'CLUSTER' in df_str and 'CD2' in df_str:
            if any(keyword in df_str for keyword in ['COMP', 'SAOD', 'PETROL', 'HEV', 'RENEWAL']):
                return 'comp_saod'
        
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
            
            # Find Cluster column
            cluster_col = None
            for j in range(df.shape[1]):
                if df.iloc[:, j].astype(str).str.contains("Cluster", case=False, na=False).any():
                    cluster_col = j
                    break
            
            if cluster_col is None:
                return []
            
            # Find CD2 columns
            cd2_cols = []
            for j in range(cluster_col + 1, df.shape[1]):
                col_str = df.iloc[:, j].astype(str).str.cat(sep=' ')
                if "CD2" in col_str.upper():
                    cd2_cols.append(j)
            
            if not cd2_cols:
                return []
            
            # Build headers
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
                        if "CD2" not in s.upper():
                            header_parts.append(s)
                headers[j] = " ".join(header_parts).strip()
            
            # Detect data start row
            if cluster_header_row is not None:
                data_start_row = cluster_header_row + 1
                while data_start_row < df.shape[0] and pd.isna(df.iloc[data_start_row, cluster_col]):
                    data_start_row += 1
            else:
                data_start_row = 10
            
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
                    
                    if "SAOD" in header_text and "COMP" not in header_text:
                        policy_type = "SAOD"
                    elif "COMP" in header_text:
                        policy_type = "COMP"
                    else:
                        policy_type = "COMP"
                    
                    fuel = "Petrol" if "PETROL" in header_text and "NON" not in header_text and "CNG" not in header_text else "Non-Petrol (incl. CNG)"
                    segment = "Non-HEV" if "NON HEV" in header_text or "NON-HEV" in header_text else "HEV"
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
            
            return records
            
        except Exception as e:
            print(f"Error processing COMP/SAOD sheet: {e}")
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
            
            for idx, row in df.iterrows():
                cluster = str(row.get('Cluster', '')).strip()
                if not cluster:
                    continue
                
                payin = safe_float(row.get('CD2'))
                if payin is None:
                    continue
                
                state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
                
                new_segment = str(row.get('New Segment Mapping', '')).strip()
                age_band = str(row.get('New Age Band', '')).strip()
                
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
            
            return records
            
        except Exception as e:
            print(f"Error processing SATP sheet: {e}")
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
        """Main entry point for processing any 4W sheet."""
        try:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)
        except:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
        pattern = Pattern4WDetector.detect_pattern(df)
        processor_class = Pattern4WDispatcher.PATTERN_PROCESSORS.get(pattern, CompSaodProcessor)
        
        records = processor_class.process(
            content, sheet_name, override_enabled, 
            override_lob, override_segment, override_policy_type
        )
        
        return records

# ===============================================================================
# API ENDPOINTS
# ===============================================================================

@app.get("/")
async def root():
    return {"message": "DIGIT 4W Processor API", "version": "1.0"}


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload an Excel file and return available worksheets."""
    try:
        # Validate file extension
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
        
        # Read file content
        content = await file.read()
        
        # Get sheet names
        xls = pd.ExcelFile(io.BytesIO(content))
        sheets = xls.sheet_names
        
        # Store file content with a unique ID
        file_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        uploaded_files[file_id] = {
            "content": content,
            "filename": file.filename,
            "sheets": sheets
        }
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "sheets": sheets,
            "message": f"File uploaded successfully. Found {len(sheets)} worksheet(s)."
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


@app.post("/process")
async def process_sheet(
    file_id: str,
    sheet_name: str,
    override_enabled: bool = False,
    override_lob: Optional[str] = None,
    override_segment: Optional[str] = None,
    override_policy_type: Optional[str] = None
):
    """Process a specific worksheet and return results."""
    try:
        # Check if file exists
        if file_id not in uploaded_files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file again.")
        
        file_data = uploaded_files[file_id]
        content = file_data["content"]
        
        # Validate sheet name
        if sheet_name not in file_data["sheets"]:
            raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in file")
        
        # Process the sheet
        records = Pattern4WDispatcher.process_sheet(
            content, sheet_name, override_enabled,
            override_lob, override_segment, override_policy_type
        )
        
        if not records:
            return {
                "success": False,
                "message": "No records extracted. Please check the sheet structure.",
                "records": [],
                "count": 0
            }
        
        # Calculate summary statistics
        states = {}
        policies = {}
        payins = []
        
        for record in records:
            state = record.get("State", "Unknown")
            states[state] = states.get(state, 0) + 1
            
            policy = record.get("Policy Type", "Unknown")
            policies[policy] = policies.get(policy, 0) + 1
            
            payin_str = record.get("Payin (CD2)", "0%")
            try:
                payin = float(payin_str.replace('%', ''))
                if payin > 0:
                    payins.append(payin)
            except:
                pass
        
        avg_payin = sum(payins) / len(payins) if payins else 0
        
        summary = {
            "total_records": len(records),
            "states": dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
            "policies": policies,
            "average_payin": round(avg_payin, 2)
        }
        
        return{
            "success": True,
            "message": f"Successfully processed {len(records)} records",
            "records": records,
            "count": len(records),
            "summary": summary
        }
        
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Error processing sheet: {str(e)}")


@app.post("/export")
async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
    """Export processed records to Excel file."""
    try:
        if not records:
            raise HTTPException(status_code=400, detail="No records to export")
        
        # Create DataFrame
        df = pd.DataFrame(records)
        
        # Generate output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
        # Create temporary file
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, filename)
        
        # Export to Excel
        df.to_excel(output_path, index=False, sheet_name='Processed')
        
        return FileResponse(
            path=output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting file: {str(e)}")


if __name__ == "__main__": 
    import uvicorn 
    uvicorn.run(app, host="0.0.0.0", port=8000)
