import fitz  # PyMuPDF
import pandas as pd
import os

def highlight_pdf(pdf_path, excel_path, highlight_type="pf", output_folder="results"):
    """
    PDF highlighter with Excel + fixed keywords.
    """
    # Fixed phrases to always highlight
    FIXED_PHRASES = [
        "Employees' State Insurance Corporation",
        "EMPLOYEE'S PROVIDENT FUND ORGANISATION"
    ]

    # Load Excel
    df = pd.read_excel(excel_path, header=None)
    excel_values = df[0].astype(str).tolist()

    # Clear old outputs
    if os.path.exists(output_folder):
        for f in os.listdir(output_folder):
            os.remove(os.path.join(output_folder, f))
    else:
        os.makedirs(output_folder)

    pdf_doc = fitz.open(pdf_path)
    matched_pages = []
    not_found = []

    for page_num in range(len(pdf_doc)):
        page = pdf_doc[page_num]
        page_matched = False

        # Get words
        words = page.get_text("words")  # x0,y0,x1,y1,text,block,line,word

        # --- Excel values highlight ---
        for val in excel_values:
            matched_words = [w for w in words if val.lower() in w[4].lower()]
            if matched_words:
                page_matched = True
                for w in matched_words:
                    x0, y0, x1, y1, text, _, _, _ = w
                    if highlight_type.lower() == "pf":
                        rect = fitz.Rect(x0, y0, x1, y1)
                        page.add_highlight_annot(rect)
                    elif highlight_type.lower() == "esic":
                        row_words = [rw for rw in words if rw[1] >= y0-1 and rw[3] <= y1+1]
                        rect = fitz.Rect(min(rw[0] for rw in row_words),
                                         y0,
                                         max(rw[2] for rw in row_words),
                                         y1)
                        page.add_highlight_annot(rect)
            else:
                if val not in not_found:
                    not_found.append(val)

        # --- Fixed phrases highlight (line-level search) ---
        for phrase in FIXED_PHRASES:
            # Search page text for phrase positions
            text_instances = page.search_for(phrase)
            for rect in text_instances:
                page.add_highlight_annot(rect)
                page_matched = True

        if page_matched:
            matched_pages.append(page_num)

    # Save only matched pages
    if matched_pages:
        new_pdf = fitz.open()
        for num in matched_pages:
            new_pdf.insert_pdf(pdf_doc, from_page=num, to_page=num)
        output_pdf_path = os.path.join(output_folder, "highlighted_output.pdf")
        new_pdf.save(output_pdf_path)
    else:
        output_pdf_path = None

    # Save Not Found Excel
    if not_found:
        not_found_df = pd.DataFrame(not_found)
        not_found_path = os.path.join(output_folder, "Data_Not_Found.xlsx")
        not_found_df.to_excel(not_found_path, index=False, header=False)
    else:
        not_found_path = None

    return output_pdf_path, not_found_path
