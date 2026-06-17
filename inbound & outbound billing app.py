import io
import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Border, Side
import streamlit as st

st.set_page_config(page_title="FedEx Universal File Merger", page_icon="📦", layout="wide")

st.markdown("""
    <style>
    html, body, [class*="css"] { font-size: 18px !important; }
    .fedex-title { color: #4D148C; font-weight: 900; font-size: 48px; letter-spacing: -1px; margin-bottom: 5px; }
    .fedex-divider { height: 6px; background: linear-gradient(to right, #4D148C 50%, #FF6600 50%); margin-bottom: 30px; border-radius: 3px; }
    .instruction-box { background-color: #F8F9FA; border-left: 5px solid #FF6600; padding: 20px; border-radius: 5px; margin-bottom: 30px; font-size: 18px; color: #333; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="fedex-title">📦 FedEx Billing Hub</div>', unsafe_allow_html=True)
st.markdown('<div class="fedex-divider"></div>', unsafe_allow_html=True)

workflow = st.radio("🔀 Select your workflow:", ["Inbound", "Outbound"], horizontal=True)

st.write("") 

if workflow == "Outbound":
    st.info("🚧 The Outbound processing module is currently under development. Please check back later!")

elif workflow == "Inbound":
    st.markdown("""
    <div class="instruction-box">
        <strong>Inbound File Portal</strong><br> 
        Upload ANY two Inbound Excel files below. Use the "Rows to skip" to bypass the junk titles. The system will auto-fill blanks, translate headers, automatically link the files using Shipment Number, and build your finalized report!
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**File 1 (Base Data)**")
        file1 = st.file_uploader("📂 Upload Primary File", type=["xlsx", "xls", "csv"])
        skip_rows_1 = st.number_input("⚙️ Rows to skip at top of File 1:", min_value=0, value=0)

    with col2:
        st.markdown("**File 2 (Data to Merge)**")
        file2 = st.file_uploader("📂 Upload Secondary File", type=["xlsx", "xls", "csv"])
        skip_rows_2 = st.number_input("⚙️ Rows to skip at top of File 2:", min_value=0, value=0)

    with st.expander("❓ What does 'Rows to skip' mean?"):
        st.write("""
        **"Junk rows"** are just the extra lines at the very top of your Excel file before your real data starts (like report titles, company logos, or blank spaces). 
        If we don't skip them, the app will get confused and try to read them as your column headers!
        """)

    st.divider()

    st.subheader("⚙️ Step 2: Output Settings")
    custom_filename = st.text_input("📝 Name your final Excel file:", value="Master_Merged_Report")
    
    final_filename = custom_filename.strip()
    if not final_filename.endswith(".xlsx"):
        final_filename += ".xlsx"
        
    st.write("") 

    if file1 is not None and file2 is not None:
        
        if st.button("🚀 Process & Combine Files", type="primary"):
            try:
                df1 = pd.read_excel(file1, skiprows=skip_rows_1)
                df2 = pd.read_excel(file2, skiprows=skip_rows_2)
                
                df1.columns = df1.columns.astype(str).str.strip()
                df2.columns = df2.columns.astype(str).str.strip()

                df1 = df1.replace(r'^\s*$', np.nan, regex=True).replace('', np.nan).ffill()
                df2 = df2.replace(r'^\s*$', np.nan, regex=True).replace('', np.nan).ffill()

                hardcoded_mapping = {
                    'DOCK DATE & TIME': 'Date Time IN',
                    'MANIFEST NBR': 'Airwaybill No.',
                    'LINE REFERENCE': 'PO Number',
                    'CUST REF NUM': 'Invoice No.',
                    'VENDOR INFO': 'Vendor Name',
                    'ITEM CODE': 'Supplier P/N',
                    'SECONDARY ITEM CODE': 'SAP No.',
                    'ITEM DESCRIPTION': 'Description',
                    'STORAGE TEMPERATURE': 'Temp',
                    'SHIPPED QTY': 'Qty',
                    'SHIPMENT NBR': 'SHIPMENT NBR' 
                }
                
                df1.columns = [hardcoded_mapping.get(str(col).strip().upper(), str(col).strip()) for col in df1.columns]
                df2.columns = [hardcoded_mapping.get(str(col).strip().upper(), str(col).strip()) for col in df2.columns]

                if 'SHIPMENT NBR' in df1.columns and 'SHIPMENT NBR' in df2.columns:
                    combined_df = pd.merge(df1, df2, on='SHIPMENT NBR', how='outer')
                    
                    for col in list(hardcoded_mapping.values()):
                        col_x = f"{col}_x"
                        col_y = f"{col}_y"
                        if col_x in combined_df.columns and col_y in combined_df.columns:
                            combined_df[col] = combined_df[col_x].combine_first(combined_df[col_y])
                            combined_df = combined_df.drop(columns=[col_x, col_y])
                else:
                    st.warning("⚠️ Could not find 'SHIPMENT NBR' in both files. The system stacked them instead of linking them.")
                    combined_df = pd.concat([df1, df2], ignore_index=True)

                combined_df = combined_df.loc[:, ~combined_df.columns.str.contains('^Unnamed')]

                final_column_order = [
                    'Date Time IN', 'Airwaybill No.', 'PO Number', 'Invoice No.',
                    'Vendor Name', 'Supplier P/N', 'SAP No.', 'Description',
                    'Type of Goods', 'Temp', 'Qty'
                ]
                
                for col in final_column_order:
                    if col not in combined_df.columns:
                        combined_df[col] = ""
                        
                combined_df = combined_df[final_column_order]

                if 'Temp' in combined_df.columns:
                    combined_df['Temp'] = combined_df['Temp'].replace(np.nan, '')
                    combined_df['Temp'] = combined_df['Temp'].astype(str).str.strip().str.title()
                    combined_df['Temp'] = combined_df['Temp'].replace(['Dg Ambient', 'Nan'], ['Ambient', ''])

                st.markdown('<div class="fedex-divider"></div>', unsafe_allow_html=True)
                
                st.subheader("✏️ Review and Edit Data")
                st.write("📝 Click directly into any cell to fix data. Click a column header to sort rows, or hover on the left to add/delete rows.")
                
                edited_df = st.data_editor(
                    combined_df, 
                    use_container_width=True,
                    num_rows="dynamic" 
                )
                st.write("") 

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    edited_df.to_excel(writer, index=False, sheet_name='Master Report')
                    worksheet = writer.sheets['Master Report']

                    header_fill = PatternFill(start_color="4D148C", end_color="4D148C", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF", size=14)
                    thick_border = Border(left=Side(style='thick'), right=Side(style='thick'), 
                                          top=Side(style='thick'), bottom=Side(style='thick'))
                    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                         top=Side(style='thin'), bottom=Side(style='thin'))

                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.border = thick_border

                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                        for cell in row:
                            cell.border = thin_border

                    for idx, col in enumerate(edited_df.columns):
                        max_data_length = edited_df[col].astype(str).str.len().max()
                        max_data_length = 0 if pd.isna(max_data_length) else int(max_data_length)
                        header_length = int(len(str(col)) * 1.4)
                        
                        perfect_width = max(max_data_length, header_length) + 4
                        col_letter = get_column_letter(idx + 1)
                        worksheet.column_dimensions[col_letter].width = perfect_width

                    worksheet.freeze_panes = 'A2'

                    summary_df = edited_df.copy()
                    summary_df['Qty_num'] = pd.to_numeric(summary_df['Qty'], errors='coerce').fillna(0)
                    summary_df['Vendor Name'] = summary_df['Vendor Name'].replace(r'^\s*$', '(Blank)', regex=True).fillna('(Blank)')
                    
                    vendor_sums = summary_df.groupby('Vendor Name')['Qty_num'].sum().reset_index()

                    worksheet.merge_cells('M4:N4')
                    title_cell = worksheet.cell(row=4, column=13, value="SUM OF QUANTITY")
                    title_cell.fill = header_fill
                    title_cell.font = header_font
                    title_cell.border = thick_border
                    worksheet.cell(row=4, column=14).border = thick_border 

                    head_v = worksheet.cell(row=5, column=13, value="Vendor Name")
                    head_t = worksheet.cell(row=5, column=14, value="Total")
                    head_v.font = Font(bold=True)
                    head_t.font = Font(bold=True)
                    head_v.border = thin_border
                    head_t.border = thin_border

                    current_row = 6
                    start_data_row = 6
                    for _, row_data in vendor_sums.iterrows():
                        worksheet.cell(row=current_row, column=13, value=row_data['Vendor Name']).border = thin_border
                        worksheet.cell(row=current_row, column=14, value=row_data['Qty_num']).border = thin_border
                        current_row += 1

                    gt_v = worksheet.cell(row=current_row, column=13, value="Grand Total")
                    
                    if current_row > start_data_row:
                        sum_formula = f"=SUM(N{start_data_row}:N{current_row - 1})"
                    else:
                        sum_formula = "=0" 

                    gt_t = worksheet.cell(row=current_row, column=14, value=sum_formula)
                    
                    gt_v.font = Font(bold=True)
                    gt_t.font = Font(bold=True)
                    gt_v.border = thin_border
                    gt_t.border = thin_border

                    worksheet.column_dimensions['M'].width = 30
                    worksheet.column_dimensions['N'].width = 15
                    
                    t2_start_row = current_row + 2
                    
                    worksheet.merge_cells(start_row=t2_start_row, start_column=13, end_row=t2_start_row, end_column=14)
                    t2_title = worksheet.cell(row=t2_start_row, column=13, value="Total IN")
                    t2_title.fill = header_fill
                    t2_title.font = header_font
                    t2_title.border = thick_border
                    worksheet.cell(row=t2_start_row, column=14).border = thick_border 
                    
                    t2_head_v = worksheet.cell(row=t2_start_row + 1, column=13, value="Temp")
                    t2_head_t = worksheet.cell(row=t2_start_row + 1, column=14, value="Type")
                    t2_head_v.font = Font(bold=True)
                    t2_head_t.font = Font(bold=True)
                    t2_head_v.border = thin_border
                    t2_head_t.border = thin_border

                    temp_categories = ["Ambient", "Chilled", "Frozen"]
                    
                    t2_current_row = t2_start_row + 2
                    t2_start_data = t2_current_row
                    
                    max_data_row = len(edited_df) + 1
                    if max_data_row < 2:
                        max_data_row = 2
                    
                    for temp_cat in temp_categories:
                        worksheet.cell(row=t2_current_row, column=13, value=temp_cat).border = thin_border
                        
                        sumifs_formula = f'=SUMIFS(K$2:K${max_data_row}, J$2:J${max_data_row}, M{t2_current_row})'
                        worksheet.cell(row=t2_current_row, column=14, value=sumifs_formula).border = thin_border
                        
                        t2_current_row += 1
                        
                    t2_gt_v = worksheet.cell(row=t2_current_row, column=13, value="Grand Total")
                    t2_sum_formula = f"=SUM(N{t2_start_data}:N{t2_current_row - 1})"
                    t2_gt_t = worksheet.cell(row=t2_current_row, column=14, value=t2_sum_formula)
                    
                    t2_gt_v.font = Font(bold=True)
                    t2_gt_t.font = Font(bold=True)
                    t2_gt_v.border = thin_border
                    t2_gt_t.border = thin_border

                buffer.seek(0)

                st.success(f"🎉 Your files were successfully processed, linked, and formatted! Ready for download.")
                
                st.download_button(
                    label=f"📥 Download {final_filename}",
                    data=buffer,
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error("🚨 An error occurred while parsing the files.")
                st.exception(e)

    else:
        st.info("Waiting for both files to be uploaded...")

# ENTER: python -m streamlit run "c:/Users/1699678/OneDrive - MyFedEx/Desktop/inbound & outbound billing app.py"