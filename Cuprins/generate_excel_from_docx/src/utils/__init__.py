def read_docx(file_path):
    from docx import Document

    document = Document(file_path)
    songs = []

    for para in document.paragraphs:
        text = para.text.strip()
        if text:
            parts = text.split(',')
            if len(parts) >= 2:
                number = parts[0].strip()
                name = parts[1].strip()
                range_ = parts[2].strip() if len(parts) > 2 else ''
                theme = parts[3].strip() if len(parts) > 3 else ''
                songs.append({
                    'number': number,
                    'name': name,
                    'range': range_,
                    'theme': theme
                })

    return songs


def save_to_excel(songs, output_file):
    import pandas as pd

    df = pd.DataFrame(songs)
    df.to_excel(output_file, index=False, sheet_name='Songs')
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Songs')
    worksheet = writer.sheets['Songs']
    
    for col in df.columns:
        column = worksheet[col + '1']
        column.auto_filter = True

    writer.save()