import pandas as pd
from docx import Document
import os

def load_data(drive_file_url):
	response = requests.get(drive_file_url)
    excel_data = BytesIO(response.content)

    may_df = pd.read_excel(excel_file, sheet_name='May')
    june_df = pd.read_excel(excel_file, sheet_name='June')
    desc_df = pd.read_excel(excel_file, sheet_name='Pre Math')
    return may_df, june_df, desc_df

def generate_pre_math_narrative(level_may, level_june, desc_df):
    def get_details(level):
        row = desc_df[desc_df['Level'] == level].iloc[0]
        return row['What I Was Learning'], row['Why It Matters'], row['Real-World Application']

    learn_june, why, use = get_details(level_june)
    if level_june > level_may:
        return f"I moved from Level {level_may} to {level_june}, learning to {learn_june.lower()}. It matters because {why.lower()}, and I use it when I {use.lower()}."
    elif level_june == level_may:
        return f"I’m still working at Level {level_may}, learning to {learn_june.lower()}. It matters because {why.lower()}, and I’m getting better every day."
    else:
        return f"I'm revisiting Level {level_june} to strengthen how I {learn_june.lower()}. It helps because {why.lower()}, and I use it when I {use.lower()}."

def create_portfolio(child_row_may, child_row_june, desc_df, output_path):
    doc = Document()
    name = child_row_june['Child Name']
    school = child_row_june['School']
    grade = child_row_june['Grade']

    doc.add_heading(f"My Learning Portfolio – {name}", 0)
    doc.add_paragraph(f"School: {school}")
    doc.add_paragraph(f"Grade: {grade}")

    doc.add_heading(f"My Pre Math Journey", level=1)
    narrative = generate_pre_math_narrative(
        int(child_row_may['Pre Math']),
        int(child_row_june['Pre Math']),
        desc_df
    )
    doc.add_paragraph(narrative)

    filepath = os.path.join(output_path, f"{name.replace(' ', '_')}_Portfolio.docx")
    doc.save(filepath)
    return filepath
