# from fastapi import FastAPI, File, UploadFile, HTTPException
# from fastapi.middleware.cors import CORSMiddleware
# from fastapi.responses import FileResponse, JSONResponse
# import pandas as pd
# import io
# import os
# from typing import List, Dict, Tuple, Optional
# from datetime import datetime
# import traceback
# import tempfile
# import shutil

# app = FastAPI(title="DIGIT 4W Processor API")

# # Enable CORS
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["https://digit-excel-private-car.vercel.app"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ===============================================================================
# # FORMULA DATA AND STATE MAPPING
# # ===============================================================================

# FORMULA_DATA = [
#     {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 30%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 40%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
#     {"LOB": "BUS", "SEGMENT": "STAFF BUS", "PO": "88% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "PO": "88% of Payin", "REMARKS": "NIL"}
# ]

# STATE_MAPPING = {
#     "DELHI": "DELHI", "Mumbai": "MAHARASHTRA", "Pune": "MAHARASHTRA", "Goa": "GOA",
#     "Kolkata": "WEST BENGAL", "Hyderabad": "TELANGANA", "Ahmedabad": "GUJARAT",
#     "Surat": "GUJARAT", "Jaipur": "RAJASTHAN", "Lucknow": "UTTAR PRADESH",
#     "Patna": "BIHAR", "Ranchi": "JHARKHAND", "Bhuvaneshwar": "ODISHA",
#     "Srinagar": "JAMMU AND KASHMIR", "Dehradun": "UTTARAKHAND", "Haridwar": "UTTARAKHAND",
#     "Bangalore": "KARNATAKA", "Jharkhand": "JHARKHAND", "Bihar": "BIHAR",
#     "Good GJ": "GUJARAT", "Bad GJ": "GUJARAT", "ROM1": "REST OF MAHARASHTRA",
#     "Good TN": "TAMIL NADU", "Kerala": "KERALA", "Good MP": "MADHYA PRADESH",
#     "Good RJ": "RAJASTHAN", "Good UP": "UTTAR PRADESH", "Punjab": "PUNJAB",
#     "Jammu": "JAMMU AND KASHMIR", "Assam": "ASSAM", "HR Ref": "HARYANA",
#     "ROM2": "REST OF MAHARASHTRA", "Andaman": "ANDAMAN AND NICOBAR ISLANDS"
# }

# # Store uploaded files temporarily
# uploaded_files = {}

# # ===============================================================================
# # CORE CALCULATION FUNCTIONS
# # ===============================================================================

# def safe_float(value) -> Optional[float]:
#     """Safely convert value to float, handling various edge cases."""
#     if pd.isna(value):
#         return None
#     s = str(value).strip().upper()
#     if s in ["D", "NA", "", "NAN", "NONE"]:
#         return None
#     try:
#         num = float(s.replace("%", ""))
#         return num * 100 if 0 < num < 1 else num
#     except:
#         return None


# def get_payin_category(payin: float) -> str:
#     """Categorize payin percentage into predefined ranges."""
#     if payin <= 20:
#         return "Payin Below 20%"
#     elif payin <= 30:
#         return "Payin 21% to 30%"
#     elif payin <= 50:
#         return "Payin 31% to 50%"
#     else:
#         return "Payin Above 50%"


# def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[str, float]:
#     """Get formula and calculate payout based on LOB, segment, policy type, and payin."""
#     segment_key = segment.upper()
    
#     if lob == "TW":
#         segment_key = "TW TP" if policy_type == "TP" else "TW SAOD + COMP"
#     elif lob == "PVT CAR":
#         segment_key = "PVT CAR TP" if policy_type == "TP" else "PVT CAR COMP + SAOD"
#     elif lob in ["TAXI", "CV", "BUS", "MISD"]:
#         segment_key = segment.upper()
    
#     payin_cat = get_payin_category(payin)
    
#     for rule in FORMULA_DATA:
#         if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
#             if rule["REMARKS"] == payin_cat or rule["REMARKS"] == "NIL":
#                 formula = rule["PO"]
#                 if "of Payin" in formula:
#                     pct = float(formula.split("%")[0].replace("Less ", ""))
#                     payout = round(payin * pct / 100, 2) if "Less" not in formula else round(payin - pct, 2)
#                 elif formula.startswith("-"):
#                     ded = float(formula.replace("%", "").replace("-", ""))
#                     payout = round(payin - ded, 2)
#                 else:
#                     payout = round(payin - 2, 2)
#                 return formula, payout
    
#     # Default fallback
#     ded = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
#     return f"-{ded}%", round(payin - ded, 2)


# def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
#     """Calculate payout with formula and rule explanation."""
#     if payin == 0:
#         return 0, "0% (No Payin)", "Payin is 0"
#     formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
#     return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"

# # ===============================================================================
# # PATTERN DETECTION
# # ===============================================================================

# class Pattern4WDetector:
#     """Detects whether a 4W sheet is COMP/SAOD or SATP pattern."""
    
#     @staticmethod
#     def detect_pattern(df: pd.DataFrame) -> str:
#         """Detect the pattern type based on sheet structure."""
#         if isinstance(df.columns, pd.Index):
#             columns = [str(col).upper().strip() for col in df.columns]
            
#             if 'CLUSTER' in columns and 'CD2' in columns:
#                 if any(keyword in ' '.join(columns) for keyword in ['NEW SEGMENT', 'AGE BAND', 'MAPPING']):
#                     return 'satp'
            
#             first_rows = df.head(10).to_string().upper()
#             if 'SATP' in first_rows or 'TP' in first_rows:
#                 return 'satp'
        
#         df_str = df.head(20).to_string().upper()
        
#         if 'CLUSTER' in df_str and 'CD2' in df_str:
#             if any(keyword in df_str for keyword in ['COMP', 'SAOD', 'PETROL', 'HEV', 'RENEWAL']):
#                 return 'comp_saod'
        
#         return 'comp_saod'
    
#     @staticmethod
#     def detect_pattern_name(df: pd.DataFrame) -> str:
#         """Get a descriptive name for the detected pattern."""
#         pattern = Pattern4WDetector.detect_pattern(df)
#         pattern_names = {
#             'comp_saod': "4W COMP/SAOD Pattern (Private Car Comprehensive)",
#             'satp': "4W SATP Pattern (Private Car Third Party)"
#         }
#         return pattern_names.get(pattern, "Unknown 4W Pattern")

# # ===============================================================================
# # PATTERN PROCESSORS
# # ===============================================================================

# class CompSaodProcessor:
#     """Process COMP/SAOD pattern sheets for Private Car."""
    
#     @staticmethod
#     def process(content: bytes, sheet_name: str,
#                 override_enabled: bool = False,
#                 override_lob: str = None,
#                 override_segment: str = None,
#                 override_policy_type: str = None) -> List[Dict]:
#         """Process COMP/SAOD pattern sheets."""
#         records = []
        
#         try:
#             df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
#             df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            
#             # Find Cluster column
#             cluster_col = None
#             for j in range(df.shape[1]):
#                 if df.iloc[:, j].astype(str).str.contains("Cluster", case=False, na=False).any():
#                     cluster_col = j
#                     break
            
#             if cluster_col is None:
#                 return []
            
#             # Find CD2 columns
#             cd2_cols = []
#             for j in range(cluster_col + 1, df.shape[1]):
#                 col_str = df.iloc[:, j].astype(str).str.cat(sep=' ')
#                 if "CD2" in col_str.upper():
#                     cd2_cols.append(j)
            
#             if not cd2_cols:
#                 return []
            
#             # Build headers
#             headers = {}
#             cluster_header_row = None
#             for i in range(df.shape[0]):
#                 if pd.notna(df.iloc[i, cluster_col]) and "cluster" in str(df.iloc[i, cluster_col]).lower():
#                     cluster_header_row = i
#                     break
            
#             header_rows_range = range(0, cluster_header_row + 3 if cluster_header_row else 10)
            
#             for j in cd2_cols:
#                 header_parts = []
#                 for i in header_rows_range:
#                     val = df.iloc[i, j]
#                     if pd.notna(val):
#                         s = str(val).strip()
#                         if "CD2" not in s.upper():
#                             header_parts.append(s)
#                 headers[j] = " ".join(header_parts).strip()
            
#             # Detect data start row
#             if cluster_header_row is not None:
#                 data_start_row = cluster_header_row + 1
#                 while data_start_row < df.shape[0] and pd.isna(df.iloc[data_start_row, cluster_col]):
#                     data_start_row += 1
#             else:
#                 data_start_row = 10
            
#             # Process data rows
#             for i in range(data_start_row, df.shape[0]):
#                 cluster_cell = df.iloc[i, cluster_col]
#                 if pd.isna(cluster_cell):
#                     continue
#                 cluster = str(cluster_cell).strip()
#                 if not cluster or "total" in cluster.lower() or cluster.lower() in ["grand total", "average"]:
#                     continue
                
#                 state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
                
#                 for j in cd2_cols:
#                     payin = safe_float(df.iloc[i, j])
#                     if payin is None:
#                         continue
                    
#                     header_text = headers.get(j, "").upper()
                    
#                     if "SAOD" in header_text and "COMP" not in header_text:
#                         policy_type = "SAOD"
#                     elif "COMP" in header_text:
#                         policy_type = "COMP"
#                     else:
#                         policy_type = "COMP"
                    
#                     fuel = "Petrol" if "PETROL" in header_text and "NON" not in header_text and "CNG" not in header_text else "Non-Petrol (incl. CNG)"
#                     segment = "Non-HEV" if "NON HEV" in header_text or "NON-HEV" in header_text else "HEV"
#                     renewal = " (Renewals)" if "RENEWAL" in header_text or "RENEW" in header_text else ""
                    
#                     orig_seg = f"PVT CAR {segment} - {fuel}{renewal}".strip()
                    
#                     lob_final = override_lob if override_enabled and override_lob else "PVT CAR"
#                     segment_final = override_segment if override_enabled and override_segment else "PVT CAR COMP + SAOD"
#                     policy_final = override_policy_type if override_policy_type else policy_type
                    
#                     payout, formula, exp = calculate_payout_with_formula(lob_final, segment_final, policy_final, payin)
                    
#                     records.append({
#                         "State": state,
#                         "Location/Cluster": cluster,
#                         "Original Segment": orig_seg,
#                         "Mapped Segment": segment_final,
#                         "LOB": lob_final,
#                         "Policy Type": policy_final,
#                         "Payin (CD2)": f"{payin:.2f}%",
#                         "Payin Category": get_payin_category(payin),
#                         "Calculated Payout": f"{payout:.2f}%",
#                         "Formula Used": formula,
#                         "Rule Explanation": exp
#                     })
            
#             return records
            
#         except Exception as e:
#             print(f"Error processing COMP/SAOD sheet: {e}")
#             traceback.print_exc()
#             return []


# class SatpProcessor:
#     """Process SATP (TP) pattern sheets for Private Car."""
    
#     @staticmethod
#     def process(content: bytes, sheet_name: str,
#                 override_enabled: bool = False,
#                 override_lob: str = None,
#                 override_segment: str = None,
#                 override_policy_type: str = None) -> List[Dict]:
#         """Process SATP pattern sheets."""
#         records = []
        
#         try:
#             df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)
            
#             for idx, row in df.iterrows():
#                 cluster = str(row.get('Cluster', '')).strip()
#                 if not cluster:
#                     continue
                
#                 payin = safe_float(row.get('CD2'))
#                 if payin is None:
#                     continue
                
#                 state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
                
#                 new_segment = str(row.get('New Segment Mapping', '')).strip()
#                 age_band = str(row.get('New Age Band', '')).strip()
                
#                 segment_desc = "PVT CAR TP"
#                 if new_segment and new_segment != 'nan':
#                     segment_desc += f" {new_segment}"
#                 if age_band and age_band != 'nan':
#                     segment_desc += f" (Age: {age_band})"
                
#                 lob_final = override_lob if override_enabled and override_lob else "PVT CAR"
#                 segment_final = override_segment if override_enabled and override_segment else "PVT CAR TP"
#                 policy_final = override_policy_type if override_policy_type else "TP"
                
#                 payout, formula, exp = calculate_payout_with_formula(lob_final, segment_final, policy_final, payin)
                
#                 records.append({
#                     "State": state.upper(),
#                     "Location/Cluster": cluster,
#                     "Original Segment": segment_desc.strip(),
#                     "Mapped Segment": segment_final,
#                     "LOB": lob_final,
#                     "Policy Type": policy_final,
#                     "Payin (CD2)": f"{payin:.2f}%",
#                     "Payin Category": get_payin_category(payin),
#                     "Calculated Payout": f"{payout:.2f}%",
#                     "Formula Used": formula,
#                     "Rule Explanation": exp
#                 })
            
#             return records
            
#         except Exception as e:
#             print(f"Error processing SATP sheet: {e}")
#             traceback.print_exc()
#             return []

# # ===============================================================================
# # PATTERN DISPATCHER
# # ===============================================================================

# class Pattern4WDispatcher:
#     """Main dispatcher that routes to appropriate 4W pattern processor."""
    
#     PATTERN_PROCESSORS = { 
#         'comp_saod': CompSaodProcessor,
#         'satp': SatpProcessor
#     }
    
#     @staticmethod
#     def process_sheet(content: bytes, sheet_name: str,
#                      override_enabled: bool = False,
#                      override_lob: str = None,
#                      override_segment: str = None,
#                      override_policy_type: str = None) -> List[Dict]:
#         """Main entry point for processing any 4W sheet."""
#         try:
#             df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)
#         except:
#             df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
#         pattern = Pattern4WDetector.detect_pattern(df)
#         processor_class = Pattern4WDispatcher.PATTERN_PROCESSORS.get(pattern, CompSaodProcessor)
        
#         records = processor_class.process(
#             content, sheet_name, override_enabled, 
#             override_lob, override_segment, override_policy_type
#         )
        
#         return records

# # ===============================================================================
# # API ENDPOINTS
# # ===============================================================================

# @app.get("/")
# async def root():
#     return {"message": "DIGIT 4W Processor API", "version": "1.0"}


# @app.post("/upload")
# async def upload_file(file: UploadFile = File(...)):
#     """Upload an Excel file and return available worksheets."""
#     try:
#         # Validate file extension
#         if not file.filename.endswith(('.xlsx', '.xls')):
#             raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
        
#         # Read file content
#         content = await file.read()
        
#         # Get sheet names
#         xls = pd.ExcelFile(io.BytesIO(content))
#         sheets = xls.sheet_names
        
#         # Store file content with a unique ID
#         file_id = datetime.now().strftime("%Y%m%d_%H%M%S")
#         uploaded_files[file_id] = {
#             "content": content,
#             "filename": file.filename,
#             "sheets": sheets
#         }
        
#         return {
#             "file_id": file_id,
#             "filename": file.filename,
#             "sheets": sheets,
#             "message": f"File uploaded successfully. Found {len(sheets)} worksheet(s)."
#         }
        
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


# @app.post("/process")
# async def process_sheet(
#     file_id: str,
#     sheet_name: str,
#     override_enabled: bool = False,
#     override_lob: Optional[str] = None,
#     override_segment: Optional[str] = None,
#     override_policy_type: Optional[str] = None
# ):
#     """Process a specific worksheet and return results."""
#     try:
#         # Check if file exists
#         if file_id not in uploaded_files:
#             raise HTTPException(status_code=404, detail="File not found. Please upload the file again.")
        
#         file_data = uploaded_files[file_id]
#         content = file_data["content"]
        
#         # Validate sheet name
#         if sheet_name not in file_data["sheets"]:
#             raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in file")
        
#         # Process the sheet
#         records = Pattern4WDispatcher.process_sheet(
#             content, sheet_name, override_enabled,
#             override_lob, override_segment, override_policy_type
#         )
        
#         if not records:
#             return {
#                 "success": False,
#                 "message": "No records extracted. Please check the sheet structure.",
#                 "records": [],
#                 "count": 0
#             }
        
#         # Calculate summary statistics
#         states = {}
#         policies = {}
#         payins = []
        
#         for record in records:
#             state = record.get("State", "Unknown")
#             states[state] = states.get(state, 0) + 1
            
#             policy = record.get("Policy Type", "Unknown")
#             policies[policy] = policies.get(policy, 0) + 1
            
#             payin_str = record.get("Payin (CD2)", "0%")
#             try:
#                 payin = float(payin_str.replace('%', ''))
#                 if payin > 0:
#                     payins.append(payin)
#             except:
#                 pass
        
#         avg_payin = sum(payins) / len(payins) if payins else 0
        
#         summary = {
#             "total_records": len(records),
#             "states": dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
#             "policies": policies,
#             "average_payin": round(avg_payin, 2)
#         }
        
#         return{
#             "success": True,
#             "message": f"Successfully processed {len(records)} records",
#             "records": records,
#             "count": len(records),
#             "summary": summary
#         }
        
#     except Exception as e:
#         traceback.print_exc()
#         raise HTTPException(status_code=500, detail=f"Error processing sheet: {str(e)}")


# @app.post("/export")
# async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
#     """Export processed records to Excel file."""
#     try:
#         if not records:
#             raise HTTPException(status_code=400, detail="No records to export")
        
#         # Create DataFrame
#         df = pd.DataFrame(records)
        
#         # Generate output filename
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         filename = f"Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
#         # Create temporary file
#         temp_dir = tempfile.gettempdir()
#         output_path = os.path.join(temp_dir, filename)
        
#         # Export to Excel
#         df.to_excel(output_path, index=False, sheet_name='Processed')
        
#         return FileResponse(
#             path=output_path,
#             filename=filename,
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
        
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Error exporting file: {str(e)}")


# if __name__ == "__main__": 
#     import uvicorn 
#     uvicorn.run(app, host="0.0.0.0", port=8000)




# from fastapi import FastAPI, File, UploadFile, HTTPException
# from fastapi.middleware.cors import CORSMiddleware
# from fastapi.responses import FileResponse, JSONResponse
# import pandas as pd
# import io
# import os
# from typing import List, Dict, Tuple, Optional
# from datetime import datetime
# import traceback
# import tempfile

# app = FastAPI(title="DIGIT 4W Processor API")

# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["https://digit-excel-private-car.vercel.app"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ===============================================================================
# # FORMULA DATA AND STATE MAPPING
# # ===============================================================================

# FORMULA_DATA = [
#     {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 30%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 40%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
#     {"LOB": "BUS", "SEGMENT": "STAFF BUS", "PO": "88% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "PO": "88% of Payin", "REMARKS": "NIL"}
# ]

# STATE_MAPPING = {
#     "DELHI": "DELHI", "DELHI NCR": "DELHI", "NCR": "DELHI",
#     "MUMBAI": "MAHARASHTRA", "PUNE": "MAHARASHTRA", "GOA": "GOA",
#     "KOLKATA": "WEST BENGAL", "HYDERABAD": "TELANGANA", "AHMEDABAD": "GUJARAT",
#     "BARODA": "GUJARAT", "SURAT": "GUJARAT", "JAIPUR": "RAJASTHAN",
#     "LUCKNOW": "UTTAR PRADESH", "PATNA": "BIHAR", "RANCHI": "JHARKHAND",
#     "BHUVANESHWAR": "ODISHA", "SRINAGAR": "JAMMU AND KASHMIR",
#     "DEHRADUN": "UTTARAKHAND", "HARIDWAR": "UTTARAKHAND",
#     "BANGALORE": "KARNATAKA", "JHARKHAND": "JHARKHAND", "BIHAR": "BIHAR",
#     "GOOD GJ": "GUJARAT", "BAD GJ": "GUJARAT", "GUJ": "GUJARAT",
#     "ROM1": "REST OF MAHARASHTRA", "ROM2": "REST OF MAHARASHTRA",
#     "REST OF MH": "REST OF MAHARASHTRA", "REST OF GUJARAT": "GUJARAT",
#     "GOOD TN": "TAMIL NADU", "KERALA": "KERALA",
#     "GOOD MP": "MADHYA PRADESH", "GOOD RJ": "RAJASTHAN",
#     "GOOD UP": "UTTAR PRADESH", "PUNJAB": "PUNJAB",
#     "JALANDHAR": "PUNJAB", "LUDHIANA": "PUNJAB",
#     "JAMMU": "JAMMU AND KASHMIR", "ASSAM": "ASSAM",
#     "HR REF": "HARYANA", "HR GOOD": "HARYANA", "HARYANA": "HARYANA",
#     "HP GOOD": "HIMACHAL PRADESH", "HP REF": "HIMACHAL PRADESH",
#     "HIMACHAL": "HIMACHAL PRADESH", "ANDAMAN": "ANDAMAN AND NICOBAR ISLANDS",
#     "JAIPUR": "RAJASTHAN", "JODHPUR": "RAJASTHAN",
#     "WEST BENGAL": "WEST BENGAL", "NORTH BENGAL": "WEST BENGAL",
#     "CHANDIGARH": "CHANDIGARH", "UTTARAKHAND": "UTTARAKHAND",
#     "KARNATAKA": "KARNATAKA", "TAMIL NADU": "TAMIL NADU",
#     "MAHARASHTRA": "MAHARASHTRA", "GUJARAT": "GUJARAT",
#     "RAJASTHAN": "RAJASTHAN", "UTTAR PRADESH": "UTTAR PRADESH",
# }

# uploaded_files = {}

# # ===============================================================================
# # HELPER FUNCTIONS
# # ===============================================================================

# def cell_to_str(val) -> str:
#     """Safely convert ANY cell value (float NaN, None, int, str) to string."""
#     if val is None:
#         return ""
#     try:
#         if pd.isna(val):
#             return ""
#     except (TypeError, ValueError):
#         pass
#     return str(val).strip()


# def safe_float(value) -> Optional[float]:
#     """Safely convert value to float, handling D, NA, blanks etc."""
#     if value is None:
#         return None
#     try:
#         if pd.isna(value):
#             return None
#     except (TypeError, ValueError):
#         pass
#     s = str(value).strip().upper().replace("%", "")
#     if s in ["D", "NA", "", "NAN", "NONE", "DECLINE"]:
#         return None
#     try:
#         num = float(s)
#         if num < 0:
#             return None
#         return num * 100 if 0 < num < 1 else num
#     except Exception:
#         return None


# def get_payin_category(payin: float) -> str:
#     if payin <= 20:
#         return "Payin Below 20%"
#     elif payin <= 30:
#         return "Payin 21% to 30%"
#     elif payin <= 50:
#         return "Payin 31% to 50%"
#     else:
#         return "Payin Above 50%"


# def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[str, float]:
#     segment_key = segment.upper()
#     if lob == "TW":
#         segment_key = "TW TP" if policy_type == "TP" else "TW SAOD + COMP"
#     elif lob == "PVT CAR":
#         segment_key = "PVT CAR TP" if policy_type == "TP" else "PVT CAR COMP + SAOD"
#     elif lob in ["TAXI", "CV", "BUS", "MISD"]:
#         segment_key = segment.upper()

#     payin_cat = get_payin_category(payin)

#     for rule in FORMULA_DATA:
#         if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
#             if rule["REMARKS"] == payin_cat or rule["REMARKS"] == "NIL":
#                 formula = rule["PO"]
#                 if "of Payin" in formula:
#                     pct = float(formula.split("%")[0].replace("Less ", "").strip())
#                     payout = round(payin - pct, 2) if "Less" in formula else round(payin * pct / 100, 2)
#                 elif formula.startswith("-"):
#                     ded = float(formula.replace("%", "").replace("-", ""))
#                     payout = round(payin - ded, 2)
#                 else:
#                     payout = round(payin - 2, 2)
#                 return formula, payout

#     # Fallback
#     ded = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
#     return f"-{ded}%", round(payin - ded, 2)


# def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
#     if payin == 0:
#         return 0, "0% (No Payin)", "Payin is 0"
#     formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
#     return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"


# def map_state(cluster: str) -> str:
#     """Map cluster name to state using STATE_MAPPING."""
#     cluster_upper = cluster.upper()
#     for key, val in STATE_MAPPING.items():
#         if key.upper() in cluster_upper:
#             return val
#     return "UNKNOWN"


# # ===============================================================================
# # PATTERN DETECTION
# # ===============================================================================

# class Pattern4WDetector:
#     """Detect whether a 4W sheet is COMP/SAOD or SATP pattern."""

#     @staticmethod
#     def detect_pattern(df: pd.DataFrame) -> str:
#         """
#         COMP/SAOD sheet → has multi-row headers with SAOD/COMP labels + CD2_OD rows
#         SATP sheet      → has simple column headers: Cluster, CD2, New Segment Mapping, etc.
#         """
#         # Build text from first 10 rows safely
#         sample = " ".join(
#             " ".join(cell_to_str(v) for v in df.iloc[i])
#             for i in range(min(10, df.shape[0]))
#         ).upper()

#         # SATP signal: has TP/SATP keyword in headers without SAOD/COMP
#         if "SATP" in sample and "SAOD" not in sample and "COMP" not in sample:
#             return "satp"

#         # COMP/SAOD signal: has SAOD or COMP as column headers
#         if "SAOD" in sample or ("COMP" in sample and "CD2" in sample):
#             return "comp_saod"

#         # Fallback: check if it looks like a simple Cluster+CD2 TP sheet
#         if "NEW SEGMENT" in sample or "AGE BAND" in sample:
#             return "satp"

#         return "comp_saod"

#     @staticmethod
#     def detect_pattern_name(df: pd.DataFrame) -> str:
#         pattern = Pattern4WDetector.detect_pattern(df)
#         return {
#             "comp_saod": "4W COMP/SAOD Pattern (Private Car Comprehensive)",
#             "satp":      "4W SATP Pattern (Private Car Third Party)"
#         }.get(pattern, "Unknown 4W Pattern")


# # ===============================================================================
# # COMP / SAOD PROCESSOR  — FIXED
# # ===============================================================================

# class CompSaodProcessor:
#     """
#     Process 4W COMP/SAOD sheets.

#     Sheet layout (from screenshot):
#       Row 0 : "Cluster" | "SAOD- Petrol" | "SAOD - Non-Petrol (incl. CNG)" | "COMP - Petrol" | ...
#       Row 1 :  (empty)  | "Non HEV"      | "Non HEV"                        | "Non HEV"       | ...
#       Row 2 :  (empty)  | "CD2_OD+Addon" | "CD2_OD+Addon"                   | "CD2_OD+Addon"  | ...
#       Row 3 :  filter row (col 0 empty — skip)
#       Row 4+: data rows  (Cluster name | values ...)

#     Processing rules:
#       - Only columns where row 2 contains "CD2" are processed (others ignored)
#       - Policy type from row 0:  SAOD → "SAOD";  COMP → "COMP"
#       - Row 1 value (e.g. "Non HEV") → goes into "Remarks" in output
#       - Row 0 label (e.g. "SAOD- Petrol") → goes into "Original Segment"
#     """

#     @staticmethod
#     def process(content: bytes, sheet_name: str,
#                 override_enabled: bool = False,
#                 override_lob: str = None,
#                 override_segment: str = None,
#                 override_policy_type: str = None) -> List[Dict]:
#         records = []
#         try:
#             # Always read raw (header=None) to keep full multi-row header structure
#             df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
    
#             # ── Step 1: find the "Cluster" header row ────────────────────────────
#             cluster_header_row = None
#             for i in range(min(15, df.shape[0])):
#                 if "CLUSTER" in cell_to_str(df.iloc[i, 0]).upper():
#                     cluster_header_row = i
#                     break

#             if cluster_header_row is None:
#                 print("   [COMP_SAOD] 'Cluster' header row not found")
#                 return records

#             # ── Step 2: sub-header rows ───────────────────────────────────────────
#             # Immediately after the Cluster row:
#             #   +1 → HEV-type labels    (Remarks)
#             #   +2 → CD2 labels         (confirms processable column)
#             hev_row = cluster_header_row + 1
#             cd2_row = cluster_header_row + 2

#             # ── Step 3: first actual data row (col-0 non-empty after cd2_row) ─────
#             data_start = cd2_row + 1
#             for i in range(cd2_row + 1, df.shape[0]):
#                 if cell_to_str(df.iloc[i, 0]):
#                     data_start = i
#                     break

#             print(f"   [COMP_SAOD] cluster_header={cluster_header_row}, "
#                   f"hev_row={hev_row}, cd2_row={cd2_row}, data_start={data_start}")

#             # ── Step 4: build column metadata ────────────────────────────────────
#             # For each column ≠ 0 whose cd2_row cell contains "CD2":
#             #   seg_header  = row cluster_header_row  e.g. "SAOD- Petrol"
#             #   hev_label   = row hev_row             e.g. "Non HEV"
#             #   policy_type = SAOD or COMP based on seg_header
#             col_meta = []
#             for col_idx in range(1, df.shape[1]):
#                 cd2_label = cell_to_str(df.iloc[cd2_row, col_idx]).upper()
#                 if "CD2" not in cd2_label:
#                     continue   # not a CD2 column — skip entirely

#                 seg_header = cell_to_str(df.iloc[cluster_header_row, col_idx])
#                 hev_label  = cell_to_str(df.iloc[hev_row, col_idx])

#                 seg_upper = seg_header.upper()
#                 if "SAOD" in seg_upper and "COMP" not in seg_upper:
#                     policy_type = "SAOD"
#                 elif "COMP" in seg_upper:
#                     policy_type = "COMP"
#                 else:
#                     policy_type = "COMP"

#                 col_meta.append({
#                     "col_idx":      col_idx,
#                     "policy_type":  policy_type,
#                     "orig_segment": seg_header,   # e.g. "SAOD- Petrol"
#                     "remarks":      hev_label,    # e.g. "Non HEV"
#                 })

#             if not col_meta:
#                 print("   [COMP_SAOD] No CD2 columns found")
#                 return records

#             print(f"   [COMP_SAOD] {len(col_meta)} CD2 column(s): "
#                   + str([f"col{m['col_idx']}={m['orig_segment']}" for m in col_meta]))

#             # ── Step 5: process data rows ─────────────────────────────────────────
#             lob_final     = override_lob     if override_enabled and override_lob     else "PVT CAR"
#             seg_default   = override_segment if override_enabled and override_segment else "PVT CAR COMP + SAOD"
#             skip_words    = {"total", "grand total", "average", "sum"}

#             for row_idx in range(data_start, df.shape[0]):
#                 cluster = cell_to_str(df.iloc[row_idx, 0])
#                 if not cluster or cluster.lower() in skip_words:
#                     continue

#                 state = map_state(cluster)

#                 for m in col_meta:
#                     payin = safe_float(df.iloc[row_idx, m["col_idx"]])
#                     if payin is None:
#                         continue

#                     policy_final  = override_policy_type if override_policy_type else m["policy_type"]
#                     segment_final = seg_default

#                     payout, formula, exp = calculate_payout_with_formula(
#                         lob_final, segment_final, policy_final, payin
#                     )

#                     records.append({
#                         "State":             state,
#                         "Location/Cluster":  cluster,
#                         "Original Segment":  m["orig_segment"],
#                         "Mapped Segment":    segment_final,
#                         "LOB":               lob_final,
#                         "Policy Type":       policy_final,
#                         "Status":            "STP",
#                         "Payin (CD2)":       f"{payin:.2f}%",
#                         "Payin Category":    get_payin_category(payin),
#                         "Calculated Payout": f"{payout:.2f}%",
#                         "Formula Used":      formula,
#                         "Rule Explanation":  exp,
#                         "Remarks":           m["remarks"],
#                     })

#             return records

#         except Exception as e:
#             print(f"   [COMP_SAOD] Error: {e}")
#             traceback.print_exc()
#             return []


# # ===============================================================================
# # SATP PROCESSOR
# # ===============================================================================

# # class SatpProcessor:
# #     """Process SATP (TP) pattern sheets for Private Car."""

# #     @staticmethod
# #     def process(content: bytes, sheet_name: str,
# #                 override_enabled: bool = False,
# #                 override_lob: str = None,
# #                 override_segment: str = None,
# #                 override_policy_type: str = None) -> List[Dict]:
# #         records = []
# #         try:
# #             df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)

# #             for idx, row in df.iterrows():
# #                 cluster = cell_to_str(row.get("Cluster", ""))
# #                 if not cluster:
# #                     continue

# #                 payin = safe_float(row.get("CD2"))
# #                 if payin is None:
# #                     continue

# #                 state = map_state(cluster)

# #                 new_segment = cell_to_str(row.get("New Segment Mapping", "") or row.get("Segment", ""))
# #                 age_band    = cell_to_str(row.get("New Age Band", "") or row.get("Age", ""))

# #                 segment_desc = "PVT CAR TP"
# #                 if new_segment and new_segment.lower() not in ("nan", ""):
# #                     segment_desc += f" {new_segment}"
# #                 if age_band and age_band.lower() not in ("nan", ""):
# #                     segment_desc += f" (Age: {age_band})"

# #                 lob_final     = override_lob     if override_enabled and override_lob     else "PVT CAR"
# #                 segment_final = override_segment if override_enabled and override_segment else "PVT CAR TP"
# #                 policy_final  = override_policy_type if override_policy_type else "TP"

# #                 payout, formula, exp = calculate_payout_with_formula(
# #                     lob_final, segment_final, policy_final, payin
# #                 )

# #                 records.append({
# #                     "State":             state.upper(),
# #                     "Location/Cluster":  cluster,
# #                     "Original Segment":  segment_desc.strip(),
# #                     "Mapped Segment":    segment_final,
# #                     "LOB":               lob_final,
# #                     "Policy Type":       policy_final,
# #                     "Status":            "STP",
# #                     "Payin (CD2)":       f"{payin:.2f}%",
# #                     "Payin Category":    get_payin_category(payin),
# #                     "Calculated Payout": f"{payout:.2f}%",
# #                     "Formula Used":      formula,
# #                     "Rule Explanation":  exp,
# #                 })

# #             return records

# #         except Exception as e:
# #             print(f"   [SATP] Error: {e}")
# #             traceback.print_exc()
# #             return []

# # IMPROVED PATTERN DETECTION AND SATP PROCESSOR
# # Replace these sections in your code

# # ===============================================================================
# # IMPROVED PATTERN DETECTION
# # ===============================================================================

# class Pattern4WDetector:
#     """Detect whether a 4W sheet is COMP/SAOD or SATP pattern."""

#     @staticmethod
#     def detect_pattern(df: pd.DataFrame) -> str:
#         """
#         COMP/SAOD sheet → has multi-row headers with SAOD/COMP labels + CD2_OD rows
#         SATP sheet      → has simple column headers: Cluster, CD2, Segment, Age
#         """
#         # Check if first row looks like simple headers
#         if df.shape[0] > 0:
#             first_row = " ".join(cell_to_str(v) for v in df.iloc[0]).upper()
            
#             # If row 0 contains standard SATP column names and NO multi-row structure
#             if "CLUSTER" in first_row and "CD2" in first_row and "SEGMENT" in first_row:
#                 # This is likely a simple header row - check if row 1 has data
#                 if df.shape[0] > 1:
#                     second_row = " ".join(cell_to_str(v) for v in df.iloc[1]).upper()
#                     # If row 1 doesn't have "CD2" or sub-headers, it's data = simple SATP
#                     if "CD2" not in second_row and "SAOD" not in second_row and "COMP" not in second_row:
#                         return "satp"
        
#         # Build text from first 10 rows safely for deeper analysis
#         sample = " ".join(
#             " ".join(cell_to_str(v) for v in df.iloc[i])
#             for i in range(min(10, df.shape[0]))
#         ).upper()

#         # COMP/SAOD signal: has SAOD or COMP as column headers with multi-row structure
#         if ("SAOD" in sample or "COMP" in sample) and "CD2_OD" in sample:
#             return "comp_saod"

#         # Default to SATP for simple structures
#         if "SEGMENT" in sample and "AGE" in sample and "CD2" in sample:
#             return "satp"

#         # Fallback
#         return "comp_saod"

#     @staticmethod
#     def detect_pattern_name(df: pd.DataFrame) -> str:
#         pattern = Pattern4WDetector.detect_pattern(df)
#         return {
#             "comp_saod": "4W COMP/SAOD Pattern (Private Car Comprehensive)",
#             "satp":      "4W SATP Pattern (Private Car Third Party)"
#         }.get(pattern, "Unknown 4W Pattern")


# # ===============================================================================
# # IMPROVED SATP PROCESSOR
# # ===============================================================================

# class SatpProcessor:
#     """Process SATP (TP) pattern sheets for Private Car."""

#     @staticmethod
#     def process(content: bytes, sheet_name: str,
#                 override_enabled: bool = False,
#                 override_lob: str = None,
#                 override_segment: str = None,
#                 override_policy_type: str = None) -> List[Dict]:
#         records = []
#         try:
#             # First, try reading with header=0 (assumes row 0 is header)
#             df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)
            
#             # Normalize column names
#             df.columns = df.columns.str.strip().str.upper()
            
#             print(f"   [SATP] Columns found: {list(df.columns)}")
            
#             # Check if we have the required columns
#             if "CLUSTER" not in df.columns or "CD2" not in df.columns:
#                 print(f"   [SATP] Missing required columns. Found: {list(df.columns)}")
#                 return records
            
#             skip_words = {"total", "grand total", "average", "sum", ""}
            
#             for idx, row in df.iterrows():
#                 cluster = cell_to_str(row.get("CLUSTER", ""))
#                 if not cluster or cluster.lower() in skip_words:
#                     continue

#                 payin = safe_float(row.get("CD2"))
#                 if payin is None:
#                     continue

#                 state = map_state(cluster)

#                 # Try multiple column name variations for segment (now all uppercase)
#                 new_segment = ""
#                 for col in ["NEW SEGMENT MAPPING", "SEGMENT", "NEW SEGMENT"]:
#                     if col in df.columns:
#                         new_segment = cell_to_str(row.get(col, ""))
#                         if new_segment and new_segment.lower() not in ("nan", "", "none"):
#                             break

#                 # Try multiple column name variations for age (now all uppercase)
#                 age_band = ""
#                 for col in ["NEW AGE BAND", "AGE", "AGE BAND"]:
#                     if col in df.columns:
#                         age_band = cell_to_str(row.get(col, ""))
#                         if age_band and age_band.lower() not in ("nan", "", "none"):
#                             break

#                 segment_desc = "PVT CAR TP"
#                 if new_segment:
#                     segment_desc += f" {new_segment}"
#                 if age_band:
#                     segment_desc += f" (Age: {age_band})"

#                 lob_final     = override_lob     if override_enabled and override_lob     else "PVT CAR"
#                 segment_final = override_segment if override_enabled and override_segment else "PVT CAR TP"
#                 policy_final  = override_policy_type if override_policy_type else "TP"

#                 payout, formula, exp = calculate_payout_with_formula(
#                     lob_final, segment_final, policy_final, payin
#                 )

#                 records.append({
#                     "State":             state.upper(),
#                     "Location/Cluster":  cluster,
#                     "Original Segment":  segment_desc.strip(),
#                     "Mapped Segment":    segment_final,
#                     "LOB":               lob_final,
#                     "Policy Type":       policy_final,
#                     "Status":            "STP",
#                     "Payin (CD2)":       f"{payin:.2f}%",
#                     "Payin Category":    get_payin_category(payin),
#                     "Calculated Payout": f"{payout:.2f}%",
#                     "Formula Used":      formula,
#                     "Rule Explanation":  exp,
#                 })

#             print(f"   [SATP] Processed {len(records)} records")
#             return records

#         except Exception as e:
#             print(f"   [SATP] Error: {e}")
#             traceback.print_exc()
#             return []


# # ===============================================================================
# # KEY IMPROVEMENTS:
# # ===============================================================================
# # 
# # 1. Pattern Detection:
# #    - Now checks if row 0 has simple column names (CLUSTER, CD2, SEGMENT)
# #    - Checks if row 1 contains data (not sub-headers)
# #    - Better distinguishes between simple SATP and multi-row COMP/SAOD
# #
# # 2. SATP Processor:
# #    - Converts ALL column names to uppercase for consistent matching
# #    - Added debug logging to show which columns are found
# #    - Better error handling if required columns are missing
# #    - More robust column name matching
# #
# # 3. The sheet in your image will now be correctly identified as SATP pattern
# #    and processed with the simple header structure

# # ===============================================================================
# # PATTERN DISPATCHER
# # ===============================================================================

# class Pattern4WDispatcher:
#     """Routes to the correct 4W processor based on detected pattern."""

#     PATTERN_PROCESSORS = {
#         "comp_saod": CompSaodProcessor,
#         "satp":      SatpProcessor,
#     }

#     @staticmethod
#     def process_sheet(content: bytes, sheet_name: str,
#                       override_enabled: bool = False,
#                       override_lob: str = None,
#                       override_segment: str = None,
#                       override_policy_type: str = None) -> List[Dict]:
#         # Detect pattern from raw (header=None) read
#         df_raw = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
#         pattern = Pattern4WDetector.detect_pattern(df_raw)
#         print(f"   [DISPATCHER] Sheet '{sheet_name}' → pattern: {pattern}")

#         processor_class = Pattern4WDispatcher.PATTERN_PROCESSORS.get(pattern, CompSaodProcessor)
#         return processor_class.process(
#             content, sheet_name,
#             override_enabled, override_lob, override_segment, override_policy_type
#         )


# # ===============================================================================
# # API ENDPOINTS
# # ===============================================================================

# @app.get("/")
# async def root():
#     return {
#         "message": "DIGIT 4W Processor API",
#         "version": "2.0.0 (Fixed CompSaod processor)",
#         "fix": "Multi-row header detection: Cluster / HEV-type / CD2 rows correctly parsed"
#     }


# @app.post("/upload")
# async def upload_file(file: UploadFile = File(...)):
#     """Upload Excel and return worksheet list."""
#     try:
#         if not file.filename.endswith((".xlsx", ".xls")):
#             raise HTTPException(status_code=400, detail="Only Excel files are supported")

#         content = await file.read()
#         xls     = pd.ExcelFile(io.BytesIO(content))
#         sheets  = xls.sheet_names

#         file_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
#         uploaded_files[file_id] = {
#             "content":  content,
#             "filename": file.filename,
#             "sheets":   sheets,
#         }

#         return {
#             "file_id":  file_id,
#             "filename": file.filename,
#             "sheets":   sheets,
#             "message":  f"Uploaded successfully. {len(sheets)} worksheet(s) found.",
#         }

#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Upload error: {str(e)}")


# @app.post("/process")
# async def process_sheet(
#     file_id: str,
#     sheet_name: str,
#     override_enabled: bool = False,
#     override_lob: Optional[str] = None,
#     override_segment: Optional[str] = None,
#     override_policy_type: Optional[str] = None,
# ):
#     """Process a specific worksheet."""
#     try:
#         if file_id not in uploaded_files:
#             raise HTTPException(status_code=404, detail="File not found. Please re-upload.")

#         file_data = uploaded_files[file_id]
#         if sheet_name not in file_data["sheets"]:
#             raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found")

#         records = Pattern4WDispatcher.process_sheet(
#             file_data["content"], sheet_name,
#             override_enabled, override_lob, override_segment, override_policy_type,
#         )

#         if not records:
#             return {
#                 "success": False,
#                 "message": "No records extracted. Check sheet structure.",
#                 "records": [],
#                 "count":   0,
#             }

#         # Summary stats
#         states   = {}
#         policies = {}
#         payins   = []
#         for r in records:
#             states[r.get("State", "?")] = states.get(r.get("State", "?"), 0) + 1
#             policies[r.get("Policy Type", "?")] = policies.get(r.get("Policy Type", "?"), 0) + 1
#             try:
#                 payins.append(float(r.get("Payin (CD2)", "0%").replace("%", "")))
#             except Exception:
#                 pass

#         avg_payin = round(sum(payins) / len(payins), 2) if payins else 0

#         return {
#             "success": True,
#             "message": f"Successfully processed {len(records)} records from '{sheet_name}'",
#             "records": records,
#             "count":   len(records),
#             "summary": {
#                 "total_records": len(records),
#                 "states":        dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
#                 "policies":      policies,
#                 "average_payin": avg_payin,
#             },
#         }

#     except Exception as e:
#         traceback.print_exc()
#         raise HTTPException(status_code=500, detail=f"Processing error: {str(e)}")


# @app.post("/export")
# async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
#     """Export records to Excel."""
#     try:
#         if not records:
#             raise HTTPException(status_code=400, detail="No records to export")

#         df        = pd.DataFrame(records)
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         filename  = f"4W_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
#         out_path  = os.path.join(tempfile.gettempdir(), filename)

#         df.to_excel(out_path, index=False, sheet_name="Processed")

#         return FileResponse(
#             path=out_path,
#             filename=filename,
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         )

#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Export error: {str(e)}")


# if __name__ == "__main__":
#     import uvicorn
#     print("\n" + "=" * 70)
#     print("DIGIT 4W Processor API - v2.0.0 (Fixed)")
#     print("Fix: CompSaod now correctly reads multi-row headers")
#     print("=" * 70 + "\n")
#     uvicorn.run(app, host="0.0.0.0", port=8000)



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

app = FastAPI(title="DIGIT 4W Processor API")

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
    "DELHI": "DELHI", "DELHI NCR": "DELHI", "NCR": "DELHI",
    "MUMBAI": "MAHARASHTRA", "PUNE": "MAHARASHTRA", "GOA": "GOA",
    "KOLKATA": "WEST BENGAL", "HYDERABAD": "TELANGANA", "AHMEDABAD": "GUJARAT",
    "BARODA": "GUJARAT", "SURAT": "GUJARAT", "JAIPUR": "RAJASTHAN",
    "LUCKNOW": "UTTAR PRADESH", "PATNA": "BIHAR", "RANCHI": "JHARKHAND",
    "BHUVANESHWAR": "ODISHA", "SRINAGAR": "JAMMU AND KASHMIR",
    "DEHRADUN": "UTTARAKHAND", "HARIDWAR": "UTTARAKHAND",
    "BANGALORE": "KARNATAKA", "JHARKHAND": "JHARKHAND", "BIHAR": "BIHAR",
    "GOOD GJ": "GUJARAT", "BAD GJ": "GUJARAT", "GUJ": "GUJARAT",
    "ROM1": "REST OF MAHARASHTRA", "ROM2": "REST OF MAHARASHTRA",
    "REST OF MH": "REST OF MAHARASHTRA", "REST OF GUJARAT": "GUJARAT",
    "GOOD TN": "TAMIL NADU", "KERALA": "KERALA",
    "GOOD MP": "MADHYA PRADESH", "GOOD RJ": "RAJASTHAN",
    "GOOD UP": "UTTAR PRADESH", "PUNJAB": "PUNJAB",
    "JALANDHAR": "PUNJAB", "LUDHIANA": "PUNJAB",
    "JAMMU": "JAMMU AND KASHMIR", "ASSAM": "ASSAM",
    "HR REF": "HARYANA", "HR GOOD": "HARYANA", "HARYANA": "HARYANA",
    "HP GOOD": "HIMACHAL PRADESH", "HP REF": "HIMACHAL PRADESH",
    "HIMACHAL": "HIMACHAL PRADESH", "ANDAMAN": "ANDAMAN AND NICOBAR ISLANDS",
    "JAIPUR": "RAJASTHAN", "JODHPUR": "RAJASTHAN",
    "WEST BENGAL": "WEST BENGAL", "NORTH BENGAL": "WEST BENGAL",
    "CHANDIGARH": "CHANDIGARH", "UTTARAKHAND": "UTTARAKHAND",
    "KARNATAKA": "KARNATAKA", "TAMIL NADU": "TAMIL NADU",
    "MAHARASHTRA": "MAHARASHTRA", "GUJARAT": "GUJARAT",
    "RAJASTHAN": "RAJASTHAN", "UTTAR PRADESH": "UTTAR PRADESH",
}

uploaded_files = {}

# ===============================================================================
# HELPER FUNCTIONS
# ===============================================================================

def cell_to_str(val) -> str:
    """Safely convert ANY cell value (float NaN, None, int, str) to string."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()


def safe_float(value) -> Optional[float]:
    """Safely convert value to float, handling D, NA, blanks etc."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    s = str(value).strip().upper().replace("%", "")
    if s in ["D", "NA", "", "NAN", "NONE", "DECLINE"]:
        return None
    try:
        num = float(s)
        if num < 0:
            return None
        return num * 100 if 0 < num < 1 else num
    except Exception:
        return None


def get_payin_category(payin: float) -> str:
    if payin <= 20:
        return "Payin Below 20%"
    elif payin <= 30:
        return "Payin 21% to 30%"
    elif payin <= 50:
        return "Payin 31% to 50%"
    else:
        return "Payin Above 50%"


def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[str, float]:
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
                    pct = float(formula.split("%")[0].replace("Less ", "").strip())
                    payout = round(payin - pct, 2) if "Less" in formula else round(payin * pct / 100, 2)
                elif formula.startswith("-"):
                    ded = float(formula.replace("%", "").replace("-", ""))
                    payout = round(payin - ded, 2)
                else:
                    payout = round(payin - 2, 2)
                return formula, payout

    # Fallback
    ded = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
    return f"-{ded}%", round(payin - ded, 2)


def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
    if payin == 0:
        return 0, "0% (No Payin)", "Payin is 0"
    formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
    return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"


def map_state(cluster: str) -> str:
    """Map cluster name to state using STATE_MAPPING."""
    cluster_upper = cluster.upper()
    for key, val in STATE_MAPPING.items():
        if key.upper() in cluster_upper:
            return val
    return "UNKNOWN"


# ===============================================================================
# PATTERN DETECTION
# ===============================================================================

class Pattern4WDetector:
    """Detect whether a 4W sheet is COMP/SAOD, SATP, or RenRoll/New pattern."""

    @staticmethod
    def detect_pattern(df: pd.DataFrame) -> str:
        """
        COMP/SAOD sheet     → multi-row headers with SAOD/COMP labels + CD2_OD rows
        SATP sheet          → simple column headers: Cluster, CD2, Segment, Age
        RenRoll/New sheet   → 4-row headers with Ren+Roll/New + SAOD(NCB)/Comp/1+3/3+3
                              + Non HEV/HEV + Net/OD+Add on
        """
        # ── Check for RenRoll/New pattern first ────────────────────────────────
        # Signals: "Ren+Roll" or "1+3" or "3+3" anywhere in first 10 rows
        sample_top = " ".join(
            " ".join(cell_to_str(v) for v in df.iloc[i])
            for i in range(min(10, df.shape[0]))
        ).upper()

        if "REN+ROLL" in sample_top or "REN + ROLL" in sample_top or "1+3" in sample_top or "3+3" in sample_top:
            return "renroll_new"

        # ── Check for simple SATP headers ──────────────────────────────────────
        if df.shape[0] > 0:
            first_row = " ".join(cell_to_str(v) for v in df.iloc[0]).upper()

            if "CLUSTER" in first_row and "CD2" in first_row and "SEGMENT" in first_row:
                if df.shape[0] > 1:
                    second_row = " ".join(cell_to_str(v) for v in df.iloc[1]).upper()
                    if "CD2" not in second_row and "SAOD" not in second_row and "COMP" not in second_row:
                        return "satp"

        # ── COMP/SAOD signal ────────────────────────────────────────────────────
        if ("SAOD" in sample_top or "COMP" in sample_top) and "CD2_OD" in sample_top:
            return "comp_saod"

        if "SEGMENT" in sample_top and "AGE" in sample_top and "CD2" in sample_top:
            return "satp"

        return "comp_saod"

    @staticmethod
    def detect_pattern_name(df: pd.DataFrame) -> str:
        pattern = Pattern4WDetector.detect_pattern(df)
        return {
            "comp_saod":    "4W COMP/SAOD Pattern (Private Car Comprehensive)",
            "satp":         "4W SATP Pattern (Private Car Third Party)",
            "renroll_new":  "4W Ren+Roll / New Pattern (Private Car SAOD/COMP with HEV)"
        }.get(pattern, "Unknown 4W Pattern")


# ===============================================================================
# COMP / SAOD PROCESSOR  — FIXED
# ===============================================================================

class CompSaodProcessor:
    """
    Process 4W COMP/SAOD sheets.

    Sheet layout (from screenshot):
      Row 0 : "Cluster" | "SAOD- Petrol" | "SAOD - Non-Petrol (incl. CNG)" | "COMP - Petrol" | ...
      Row 1 :  (empty)  | "Non HEV"      | "Non HEV"                        | "Non HEV"       | ...
      Row 2 :  (empty)  | "CD2_OD+Addon" | "CD2_OD+Addon"                   | "CD2_OD+Addon"  | ...
      Row 3 :  filter row (col 0 empty — skip)
      Row 4+: data rows  (Cluster name | values ...)

    Processing rules:
      - Only columns where row 2 contains "CD2" are processed (others ignored)
      - Policy type from row 0:  SAOD → "SAOD";  COMP → "COMP"
      - Row 1 value (e.g. "Non HEV") → goes into "Remarks" in output
      - Row 0 label (e.g. "SAOD- Petrol") → goes into "Original Segment"
    """

    @staticmethod
    def process(content: bytes, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None,
                override_policy_type: str = None) -> List[Dict]:
        records = []
        try:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)

            # ── Step 1: find the "Cluster" header row ────────────────────────────
            cluster_header_row = None
            for i in range(min(15, df.shape[0])):
                if "CLUSTER" in cell_to_str(df.iloc[i, 0]).upper():
                    cluster_header_row = i
                    break

            if cluster_header_row is None:
                print("   [COMP_SAOD] 'Cluster' header row not found")
                return records

            # ── Step 2: sub-header rows ───────────────────────────────────────────
            hev_row = cluster_header_row + 1
            cd2_row = cluster_header_row + 2

            # ── Step 3: first actual data row ─────────────────────────────────────
            data_start = cd2_row + 1
            for i in range(cd2_row + 1, df.shape[0]):
                if cell_to_str(df.iloc[i, 0]):
                    data_start = i
                    break

            print(f"   [COMP_SAOD] cluster_header={cluster_header_row}, "
                  f"hev_row={hev_row}, cd2_row={cd2_row}, data_start={data_start}")

            # ── Step 4: build column metadata ────────────────────────────────────
            col_meta = []
            for col_idx in range(1, df.shape[1]):
                cd2_label = cell_to_str(df.iloc[cd2_row, col_idx]).upper()
                if "CD2" not in cd2_label:
                    continue

                seg_header = cell_to_str(df.iloc[cluster_header_row, col_idx])
                hev_label  = cell_to_str(df.iloc[hev_row, col_idx])

                seg_upper = seg_header.upper()
                if "SAOD" in seg_upper and "COMP" not in seg_upper:
                    policy_type = "SAOD"
                elif "COMP" in seg_upper:
                    policy_type = "COMP"
                else:
                    policy_type = "COMP"

                col_meta.append({
                    "col_idx":      col_idx,
                    "policy_type":  policy_type,
                    "orig_segment": seg_header,
                    "remarks":      hev_label,
                })

            if not col_meta:
                print("   [COMP_SAOD] No CD2 columns found")
                return records

            print(f"   [COMP_SAOD] {len(col_meta)} CD2 column(s): "
                  + str([f"col{m['col_idx']}={m['orig_segment']}" for m in col_meta]))

            # ── Step 5: process data rows ─────────────────────────────────────────
            lob_final     = override_lob     if override_enabled and override_lob     else "PVT CAR"
            seg_default   = override_segment if override_enabled and override_segment else "PVT CAR COMP + SAOD"
            skip_words    = {"total", "grand total", "average", "sum"}

            for row_idx in range(data_start, df.shape[0]):
                cluster = cell_to_str(df.iloc[row_idx, 0])
                if not cluster or cluster.lower() in skip_words:
                    continue

                state = map_state(cluster)

                for m in col_meta:
                    payin = safe_float(df.iloc[row_idx, m["col_idx"]])
                    if payin is None:
                        continue

                    policy_final  = override_policy_type if override_policy_type else m["policy_type"]
                    segment_final = seg_default

                    payout, formula, exp = calculate_payout_with_formula(
                        lob_final, segment_final, policy_final, payin
                    )

                    records.append({
                        "State":             state,
                        "Location/Cluster":  cluster,
                        "Original Segment":  m["orig_segment"],
                        "Mapped Segment":    segment_final,
                        "LOB":               lob_final,
                        "Policy Type":       policy_final,
                        "Status":            "STP",
                        "Payin (CD2)":       f"{payin:.2f}%",
                        "Payin Category":    get_payin_category(payin),
                        "Calculated Payout": f"{payout:.2f}%",
                        "Formula Used":      formula,
                        "Rule Explanation":  exp,
                        "Remarks":           m["remarks"],
                    })

            return records

        except Exception as e:
            print(f"   [COMP_SAOD] Error: {e}")
            traceback.print_exc()
            return []


# ===============================================================================
# IMPROVED SATP PROCESSOR
# ===============================================================================

class SatpProcessor:
    """Process SATP (TP) pattern sheets for Private Car."""

    @staticmethod
    def process(content: bytes, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None,
                override_policy_type: str = None) -> List[Dict]:
        records = []
        try:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)

            # Normalize column names
            df.columns = df.columns.str.strip().str.upper()

            print(f"   [SATP] Columns found: {list(df.columns)}")

            if "CLUSTER" not in df.columns or "CD2" not in df.columns:
                print(f"   [SATP] Missing required columns. Found: {list(df.columns)}")
                return records

            skip_words = {"total", "grand total", "average", "sum", ""}

            for idx, row in df.iterrows():
                cluster = cell_to_str(row.get("CLUSTER", ""))
                if not cluster or cluster.lower() in skip_words:
                    continue

                payin = safe_float(row.get("CD2"))
                if payin is None:
                    continue

                state = map_state(cluster)

                new_segment = ""
                for col in ["NEW SEGMENT MAPPING", "SEGMENT", "NEW SEGMENT"]:
                    if col in df.columns:
                        new_segment = cell_to_str(row.get(col, ""))
                        if new_segment and new_segment.lower() not in ("nan", "", "none"):
                            break

                age_band = ""
                for col in ["NEW AGE BAND", "AGE", "AGE BAND"]:
                    if col in df.columns:
                        age_band = cell_to_str(row.get(col, ""))
                        if age_band and age_band.lower() not in ("nan", "", "none"):
                            break

                segment_desc = "PVT CAR TP"
                if new_segment:
                    segment_desc += f" {new_segment}"
                if age_band:
                    segment_desc += f" (Age: {age_band})"

                lob_final     = override_lob     if override_enabled and override_lob     else "PVT CAR"
                segment_final = override_segment if override_enabled and override_segment else "PVT CAR TP"
                policy_final  = override_policy_type if override_policy_type else "TP"

                payout, formula, exp = calculate_payout_with_formula(
                    lob_final, segment_final, policy_final, payin
                )

                records.append({
                    "State":             state.upper(),
                    "Location/Cluster":  cluster,
                    "Original Segment":  segment_desc.strip(),
                    "Mapped Segment":    segment_final,
                    "LOB":               lob_final,
                    "Policy Type":       policy_final,
                    "Status":            "STP",
                    "Payin (CD2)":       f"{payin:.2f}%",
                    "Payin Category":    get_payin_category(payin),
                    "Calculated Payout": f"{payout:.2f}%",
                    "Formula Used":      formula,
                    "Rule Explanation":  exp,
                })

            print(f"   [SATP] Processed {len(records)} records")
            return records

        except Exception as e:
            print(f"   [SATP] Error: {e}")
            traceback.print_exc()
            return []


# ===============================================================================
# REN+ROLL / NEW (4-ROW HEADER) PROCESSOR  ← NEW PATTERN
# ===============================================================================
#
# Sheet layout (from screenshot):
#
#   Row 0 : "Cluster" | "Ren+Roll" (merged) | "Ren+Roll" (merged) | "Ren+Roll" | "New" (merged)
#   Row 1 :  (empty)  | "SAOD (NCB)" | "SAOD (w/o NCB)" | "Comp (with Addon)" |
#             "Comp (without Addon)" | "All" | "1+3/ 3+3" | "1+3/ 3+3"
#   Row 2 :  (empty)  | "Non HEV" | "Non HEV" | "Non HEV" | "Non HEV" | "HEV" | "Non HEV" | "HEV"
#   Row 3 :  (empty)  | "Net" | "Net" | "Net" | "Net" | "OD +Add on" | "OD +Add on" | "OD +Add on"
#   Row 4 :  filter row (skip)
#   Row 5+:  data rows  (Cluster name | values ...)
#
# Remarks format: "<Business Type> | <HEV Label> | <Net/Addon Type>"
#   e.g. "Ren+Roll | Non HEV | Net"   or   "New | HEV | OD +Add on"
#
# Policy type:
#   SAOD (NCB), SAOD (w/o NCB)          → "SAOD"
#   Comp (with Addon), Comp (w/o Addon) → "COMP"
#   All, 1+3/ 3+3                       → "COMP"  (new car OD)

class RenRollNewProcessor:
    """
    Process 4W sheets with the 4-row merged-header structure:
      Row 0: Business type  → Ren+Roll / New
      Row 1: Sub-segment    → SAOD (NCB) / SAOD (w/o NCB) / Comp (with Addon) /
                               Comp (without Addon) / All / 1+3/ 3+3
      Row 2: HEV label      → Non HEV / HEV
      Row 3: Net/Addon type → Net / OD +Add on
    """

    @staticmethod
    def process(content: bytes, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None,
                override_policy_type: str = None) -> List[Dict]:
        records = []
        try:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)

            # ── Step 1: find the "Cluster" header row ────────────────────────────
            cluster_header_row = None
            for i in range(min(15, df.shape[0])):
                if "CLUSTER" in cell_to_str(df.iloc[i, 0]).upper():
                    cluster_header_row = i
                    break

            if cluster_header_row is None:
                print("   [RENROLL_NEW] 'Cluster' header row not found")
                return records

            # ── Step 2: map the four header sub-rows ─────────────────────────────
            biz_row     = cluster_header_row          # Row 0: Ren+Roll / New
            sub_seg_row = cluster_header_row + 1      # Row 1: SAOD (NCB) / Comp etc.
            hev_row     = cluster_header_row + 2      # Row 2: Non HEV / HEV
            net_row     = cluster_header_row + 3      # Row 3: Net / OD +Add on
            # Row +4 is typically a filter/blank row → skip
            data_start  = cluster_header_row + 5

            # Safety: scan forward to find first non-empty cluster cell
            for i in range(data_start, df.shape[0]):
                if cell_to_str(df.iloc[i, 0]):
                    data_start = i
                    break

            print(f"   [RENROLL_NEW] cluster_header={cluster_header_row}, "
                  f"biz_row={biz_row}, sub_seg_row={sub_seg_row}, "
                  f"hev_row={hev_row}, net_row={net_row}, data_start={data_start}")

            # ── Step 3: build column metadata ────────────────────────────────────
            # Excel merges cells: we forward-fill the business-type row (Row 0)
            # and sub-segment row (Row 1) so every column has the correct label.
            last_biz     = ""
            last_sub_seg = ""

            col_meta = []
            for col_idx in range(1, df.shape[1]):
                biz_val     = cell_to_str(df.iloc[biz_row,     col_idx])
                sub_seg_val = cell_to_str(df.iloc[sub_seg_row, col_idx])
                hev_val     = cell_to_str(df.iloc[hev_row,     col_idx])
                net_val     = cell_to_str(df.iloc[net_row,     col_idx])

                # Forward-fill merged cells
                if biz_val:
                    last_biz = biz_val
                if sub_seg_val:
                    last_sub_seg = sub_seg_val

                effective_biz     = last_biz
                effective_sub_seg = last_sub_seg

                # Skip columns with no meaningful sub-segment header
                if not effective_sub_seg:
                    continue

                # Determine policy type from sub-segment label
                sub_upper = effective_sub_seg.upper()
                if "SAOD" in sub_upper:
                    policy_type = "SAOD"
                elif "COMP" in sub_upper or "ADDON" in sub_upper or "ADD ON" in sub_upper:
                    policy_type = "COMP"
                elif "1+3" in sub_upper or "3+3" in sub_upper:
                    policy_type = "COMP"    # New car OD
                else:
                    policy_type = "COMP"    # Default

                # Build remarks: Business Type | HEV Label | Net/Addon Type
                remarks_parts = []
                if effective_biz:
                    remarks_parts.append(effective_biz)     # e.g. "Ren+Roll"
                if hev_val:
                    remarks_parts.append(hev_val)           # e.g. "Non HEV"
                if net_val:
                    remarks_parts.append(net_val)           # e.g. "Net"
                remarks = " | ".join(remarks_parts)

                col_meta.append({
                    "col_idx":      col_idx,
                    "policy_type":  policy_type,
                    "orig_segment": effective_sub_seg,  # e.g. "SAOD (NCB)"
                    "biz_type":     effective_biz,      # e.g. "Ren+Roll"
                    "hev_label":    hev_val,            # e.g. "Non HEV"
                    "net_label":    net_val,            # e.g. "Net"
                    "remarks":      remarks,
                })

            if not col_meta:
                print("   [RENROLL_NEW] No data columns found")
                return records

            print(f"   [RENROLL_NEW] {len(col_meta)} column(s): "
                  + str([f"col{m['col_idx']}={m['orig_segment']} ({m['biz_type']})" for m in col_meta]))

            # ── Step 4: process data rows ─────────────────────────────────────────
            lob_final   = override_lob     if override_enabled and override_lob     else "PVT CAR"
            seg_default = override_segment if override_enabled and override_segment else "PVT CAR COMP + SAOD"
            skip_words  = {"total", "grand total", "average", "sum"}

            for row_idx in range(data_start, df.shape[0]):
                cluster = cell_to_str(df.iloc[row_idx, 0])
                if not cluster or cluster.lower() in skip_words:
                    continue

                state = map_state(cluster)

                for m in col_meta:
                    payin = safe_float(df.iloc[row_idx, m["col_idx"]])
                    if payin is None:
                        continue

                    policy_final  = override_policy_type if override_policy_type else m["policy_type"]
                    segment_final = seg_default

                    payout, formula, exp = calculate_payout_with_formula(
                        lob_final, segment_final, policy_final, payin
                    )

                    records.append({
                        "State":             state,
                        "Location/Cluster":  cluster,
                        "Original Segment":  m["orig_segment"],
                        "Mapped Segment":    segment_final,
                        "LOB":               lob_final,
                        "Policy Type":       policy_final,
                        "Status":            "STP",
                        "Payin (CD2)":       f"{payin:.2f}%",
                        "Payin Category":    get_payin_category(payin),
                        "Calculated Payout": f"{payout:.2f}%",
                        "Formula Used":      formula,
                        "Rule Explanation":  exp,
                        "Remarks":           m["remarks"],
                    })

            print(f"   [RENROLL_NEW] Processed {len(records)} records")
            return records

        except Exception as e:
            print(f"   [RENROLL_NEW] Error: {e}")
            traceback.print_exc()
            return []


# ===============================================================================
# PATTERN DISPATCHER
# ===============================================================================

class Pattern4WDispatcher:
    """Routes to the correct 4W processor based on detected pattern."""

    PATTERN_PROCESSORS = {
        "comp_saod":   CompSaodProcessor,
        "satp":        SatpProcessor,
        "renroll_new": RenRollNewProcessor,   # ← NEW
    }

    @staticmethod
    def process_sheet(content: bytes, sheet_name: str,
                      override_enabled: bool = False,
                      override_lob: str = None,
                      override_segment: str = None,
                      override_policy_type: str = None) -> List[Dict]:
        # Detect pattern from raw (header=None) read
        df_raw = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        pattern = Pattern4WDetector.detect_pattern(df_raw)
        print(f"   [DISPATCHER] Sheet '{sheet_name}' → pattern: {pattern}")

        processor_class = Pattern4WDispatcher.PATTERN_PROCESSORS.get(pattern, CompSaodProcessor)
        return processor_class.process(
            content, sheet_name,
            override_enabled, override_lob, override_segment, override_policy_type
        )


# ===============================================================================
# API ENDPOINTS
# ===============================================================================

@app.get("/")
async def root():
    return {
        "message": "DIGIT 4W Processor API",
        "version": "3.0.0 (RenRoll/New pattern added)",
        "patterns_supported": [
            "comp_saod   → 4W COMP/SAOD (3-row header: Segment / HEV / CD2)",
            "satp        → 4W SATP (simple header: Cluster, CD2, Segment, Age)",
            "renroll_new → 4W Ren+Roll/New (4-row header: BizType / SubSeg / HEV / Net)"
        ]
    }


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload Excel and return worksheet list."""
    try:
        if not file.filename.endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Only Excel files are supported")

        content = await file.read()
        xls     = pd.ExcelFile(io.BytesIO(content))
        sheets  = xls.sheet_names

        file_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        uploaded_files[file_id] = {
            "content":  content,
            "filename": file.filename,
            "sheets":   sheets,
        }

        return {
            "file_id":  file_id,
            "filename": file.filename,
            "sheets":   sheets,
            "message":  f"Uploaded successfully. {len(sheets)} worksheet(s) found.",
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Upload error: {str(e)}")


@app.post("/process")
async def process_sheet(
    file_id: str,
    sheet_name: str,
    override_enabled: bool = False,
    override_lob: Optional[str] = None,
    override_segment: Optional[str] = None,
    override_policy_type: Optional[str] = None,
):
    """Process a specific worksheet."""
    try:
        if file_id not in uploaded_files:
            raise HTTPException(status_code=404, detail="File not found. Please re-upload.")

        file_data = uploaded_files[file_id]
        if sheet_name not in file_data["sheets"]:
            raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found")

        records = Pattern4WDispatcher.process_sheet(
            file_data["content"], sheet_name,
            override_enabled, override_lob, override_segment, override_policy_type,
        )

        if not records:
            return {
                "success": False,
                "message": "No records extracted. Check sheet structure.",
                "records": [],
                "count":   0,
            }

        # Summary stats
        states   = {}
        policies = {}
        payins   = []
        for r in records:
            states[r.get("State", "?")] = states.get(r.get("State", "?"), 0) + 1
            policies[r.get("Policy Type", "?")] = policies.get(r.get("Policy Type", "?"), 0) + 1
            try:
                payins.append(float(r.get("Payin (CD2)", "0%").replace("%", "")))
            except Exception:
                pass

        avg_payin = round(sum(payins) / len(payins), 2) if payins else 0

        return {
            "success": True,
            "message": f"Successfully processed {len(records)} records from '{sheet_name}'",
            "records": records,
            "count":   len(records),
            "summary": {
                "total_records": len(records),
                "states":        dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
                "policies":      policies,
                "average_payin": avg_payin,
            },
        }

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Processing error: {str(e)}")


@app.post("/export")
async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
    """Export records to Excel."""
    try:
        if not records:
            raise HTTPException(status_code=400, detail="No records to export")

        df        = pd.DataFrame(records)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename  = f"4W_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        out_path  = os.path.join(tempfile.gettempdir(), filename)

        df.to_excel(out_path, index=False, sheet_name="Processed")

        return FileResponse(
            path=out_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export error: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    print("\n" + "=" * 70)
    print("DIGIT 4W Processor API - v3.0.0")
    print("Added: RenRoll/New pattern (4-row header with HEV + Ren+Roll/New)")
    print("=" * 70 + "\n")
    uvicorn.run(app, host="0.0.0.0", port=8000)
