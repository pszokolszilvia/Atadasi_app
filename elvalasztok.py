import os
import shutil

def generate_elvalasztok(borito_dir, oldalelvalaszto_sablon, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pythoncom.CoInitialize()
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False

    borito_files = sorted(f for f in os.listdir(borito_dir) if f.endswith(".docx"))

    for idx, borito_file in enumerate(borito_files, 1):
        borito_path = os.path.abspath(os.path.join(borito_dir, borito_file))

        doc_borito = word.Documents.Open(borito_path)
        try:
            if doc_borito.Tables.Count > 0:
                table = doc_borito.Tables(1)
                cell = table.Cell(table.Rows.Count, table.Columns.Count)
                cim = cell.Range.Text.strip().replace('\r', '').replace('\x07', '')
            else:
                cim = doc_borito.Content.Text.strip()
        except Exception as e:
            print(f"Hiba a borító ({borito_file}) szöveg kiolvasásakor: {e}")
            cim = "[Nincs szöveg]"
        doc_borito.Close()

        output_path = os.path.abspath(os.path.join(output_dir, f"elvalaszto_{idx:03d}.docx"))
        shutil.copy(oldalelvalaszto_sablon, output_path)
        doc_elv = word.Documents.Open(output_path)

        try:
            for section in doc_elv.Sections:
                footer = section.Footers(win32.constants.wdHeaderFooterPrimary)
                footer_text = footer.Range.Text.strip().replace('\r', '').replace('\x07', '')
                if "SZÖVEGHELY" in footer_text:
                    footer.Range.Text = footer_text.replace("SZÖVEGHELY", cim)
        except Exception as e:
            print(f"Hiba élőlábnál ({borito_file}): {e}")

        doc_elv.SaveAs(output_path)
        doc_elv.Close()

    word.Quit()
