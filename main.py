from io import BytesIO
from pptx import Presentation
from PyPDF2 import PdfMerger
import streamlit as st
from pathlib import Path

# Temporary directory for converted files
temp_dir = Path("temp_files")
temp_dir.mkdir(exist_ok=True)

merger = PdfMerger()
for sorted_filename in filenames_sorted:
    file = next(f for f in uploaded_files if f.name == sorted_filename)
    try:
        if file.type == "application/pdf":
            # PDF 파일은 그대로 추가
            merger.append(file)
        elif file.type in ["application/vnd.openxmlformats-officedocument.presentationml.presentation"]:
            # PPTX 파일만 처리
            presentation = Presentation(BytesIO(file.read()))
            pdf_path = temp_dir / f"{file.name}.pdf"
            presentation.save(pdf_path)
            merger.append(str(pdf_path))
        else:
            st.error(f"지원되지 않는 파일 형식: {file.name}")
    except Exception as e:
        st.error(f"파일 '{file.name}'을 처리하는 중 오류 발생: {e}")
        continue

# Save and download merged file
final_path = temp_dir / output_name
merger.write(final_path)
merger.close()

with open(final_path, "rb") as f:
    st.download_button(
        "📥 결합된 파일 다운로드",
        f,
        file_name=output_name,
        mime="application/pdf"
    )

# Clean up temporary files
for temp_file in temp_dir.iterdir():
    temp_file.unlink()
temp_dir.rmdir()
