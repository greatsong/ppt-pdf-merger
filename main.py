import streamlit as st
from PyPDF2 import PdfMerger
from pptx import Presentation
from pathlib import Path
import os

# App Title
st.title("📎 파일 결합 도구 (By 석리송)")

# Friendly Introduction
st.write(
    """
    이 앱은 여러 개의 PPT, PPTX, 또는 PDF 파일을 하나의 PDF 파일로 결합할 수 있도록 도와줍니다. 😊  
    **사용 방법:**  
    1. 파일을 업로드합니다 (PPT, PPTX, PDF만 가능합니다).  
    2. 파일 순서를 확인하거나 변경합니다.  
    3. 결합된 파일의 이름을 설정합니다.  
    4. '결합하기' 버튼을 클릭하여 하나의 PDF로 결합합니다!  
    """
)

# File Upload
uploaded_files = st.file_uploader(
    "📤 파일을 업로드하세요 (PPT, PPTX, PDF 지원)", 
    type=["ppt", "pptx", "pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    # Display uploaded files and allow reordering
    st.write("### 업로드된 파일 목록 및 순서 조정")
    filenames = [file.name for file in uploaded_files]
    filenames_sorted = st.experimental_data_editor(
        pd.DataFrame({"파일 이름": filenames}),
        use_container_width=True
    )["파일 이름"].tolist()

    # File name for the final merged file
    default_output_name = "결합된_파일.pdf"
    output_name = st.text_input(
        "📁 저장할 파일 이름을 입력하세요 (기본값: 결합된_파일.pdf)", 
        value=default_output_name
    )

    # Merge files
    if st.button("결합하기"):
        # Temporary directory to store converted files
        temp_dir = Path("temp_files")
        temp_dir.mkdir(exist_ok=True)

        merger = PdfMerger()
        for filename in filenames_sorted:
            file = next(f for f in uploaded_files if f.name == filename)
            if file.type in ["application/pdf"]:
                # Directly add PDF files
                merger.append(file)
            elif file.type in ["application/vnd.ms-powerpoint", "application/vnd.openxmlformats-officedocument.presentationml.presentation"]:
                # Convert PPT/PPTX to PDF
                presentation = Presentation(file)
                pdf_path = temp_dir / f"{file.name}.pdf"
                presentation.save(pdf_path)
                merger.append(str(pdf_path))

        # Save the merged file
        final_path = temp_dir / output_name
        merger.write(final_path)
        merger.close()

        # Provide download link
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

        st.success("파일 결합이 완료되었습니다! 🎉")
