from flask import Flask, request, send_file, jsonify
import pdfplumber
import pandas as pd
import re
import io

app = Flask(__name__)

@app.route('/api/convert', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "업로드된 파일이 없습니다."}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "선택된 파일이 없습니다."}), 400

    try:
        all_items = []
        invoice_numbers = set()
        po_numbers = set()
        total_amount_sum = 0.0
        total_net_weight = 0.0

        with pdfplumber.open(file) as pdf:
            for i, page in enumerate(pdf.pages[1:]):
                text = page.extract_text()
                if not text:
                    continue
                
                invoice_match = re.search(r"INVOICE\s*#?\s*:\s*([A-Z0-9\-]+)", text)
                po_match = re.search(r"PO NO\.?\s*#?\s*:\s*([A-Z0-9\-]+)", text)
                nw_match = re.search(r"NET WEIGHT:\s*([\d\.]+)\s*KGS", text)
                
                if invoice_match:
                    invoice_numbers.add(invoice_match.group(1).strip())
                if po_match:
                    po_numbers.add(po_match.group(1).strip())
                if nw_match:
                    total_net_weight += float(nw_match.group(1).strip())

                is_ncv = "NO COMMERCIAL VALUE" in text
                term_val = "NO COMMERCIAL VALUE" if is_ncv else ""

                tables = page.extract_tables()
                for table in tables:
                    header_idx = -1
                    for idx, row in enumerate(table):
                        row_text = " ".join([str(c) for c in row if c])
                        if "ITEM" in row_text and "Q'TY" in row_text:
                            header_idx = idx
                            break
                    
                    if header_idx != -1:
                        for row in table[header_idx+1:]:
                            row_str = " ".join([str(c) for c in row if c])
                            if not row_str.strip() or "TOTAL" in row_str.upper():
                                continue
                            
                            clean_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                            
                            if len(clean_row) < 5 or not clean_row[0].isdigit():
                                continue
                                
                            item_no = clean_row[0]
                            desc = clean_row[1] + " " + clean_row[2] if len(clean_row) > 6 else clean_row[1]
                            desc = desc.strip()
                            
                            qty = clean_row[-4] if len(clean_row) >= 4 else "1"
                            um = clean_row[-3] if len(clean_row) >= 3 else "PC"
                            up = clean_row[-2] if len(clean_row) >= 2 else "0"
                            amount = clean_row[-1] if len(clean_row) >= 1 else "0"

                            up_val = re.sub(r'[^\d\.]', '', up)
                            amount_val = re.sub(r'[^\d\.]', '', amount)

                            all_items.append({
                                "ITEM": item_no,
                                "Item Code(Pre PR) / DESCRIPTION": f"PHOTOMASK {desc}",
                                "MASK NAME": "", 
                                "Q'TY": qty,
                                "U/M": um,
                                "U/P (USD)": up_val,
                                "AMOUNT (USD)": amount_val,
                                "Term": term_val
                            })
                            
                            if amount_val:
                                total_amount_sum += float(amount_val)

        df = pd.DataFrame(all_items)
        result_columns = ["col1", "ITEM", "Item Code(Pre PR) / DESCRIPTION", "MASK NAME", "col2", "Q'TY", "U/M", "U/P (USD)", "AMOUNT (USD)", "Term"]
        df_final = pd.DataFrame(columns=result_columns)
        
        for idx, row in df.iterrows():
            df_final.loc[idx] = ["", row["ITEM"], row["Item Code(Pre PR) / DESCRIPTION"], row["MASK NAME"], "", row["Q'TY"], row["U/M"], row["U/P (USD)"], row["AMOUNT (USD)"], row["Term"]]

        inv_list = sorted(list(invoice_numbers))
        inv_range = f"{inv_list[0]}-{inv_list[-1]}" if len(inv_list) > 1 else (inv_list[0] if inv_list else "")
        po_str = ", ".join(sorted(list(po_numbers)))

        footer_rows = [
            {"col1": "", "ITEM": "", "Item Code(Pre PR) / DESCRIPTION": "", "MASK NAME": "", "col2": "", "Q'TY": "", "U/M": "", "U/P (USD)": "", "AMOUNT (USD)": "", "Term": ""},
            {"col1": "", "ITEM": "", "Item Code(Pre PR) / DESCRIPTION": "", "MASK NAME": "", "col2": "", "Q'TY": "", "U/M": "", "U/P (USD)": "", "AMOUNT (USD)": "", "Term": ""},
            {"col1": "", "ITEM": "INVOICE NO.", "Item Code(Pre PR) / DESCRIPTION": inv_range, "MASK NAME": f"PO NO. {po_str}", "col2": "", "Q'TY": "", "U/M": "", "U/P (USD)": "TOTAL AMOUNT", "AMOUNT (USD)": str(int(total_amount_sum)), "Term": ""},
            {"col1": "", "ITEM": "", "Item Code(Pre PR) / DESCRIPTION": "", "MASK NAME": "", "col2": "", "Q'TY": "", "U/M": "NET WEIGHT", "U/P (USD)": str(total_net_weight), "AMOUNT (USD)": "", "Term": ""}
        ]
        
        df_footer = pd.DataFrame(footer_rows)
        df_combined = pd.concat([df_final, df_footer], ignore_index=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_combined.to_excel(writer, index=False, header=[col if "col" not in col else "" for col in df_combined.columns])
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name="Converted_Invoice.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500