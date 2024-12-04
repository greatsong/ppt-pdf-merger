import streamlit as st
from PyPDF2 import PdfMerger
from pptx import Presentation
from pathlib import Path
from streamlit_sortables import sort_items

# App Title
st.title("📎 파일 결합 도구 (By 석리송)")

# Friendly Introduction
st.write(
    """
    이 앱은 여러 개의 PPT, PPTX, 또는 PDF 파일을 하나의 PDF 파일로 결합할 수 있도록 도와줍니다. 😊  
    **사용 방법:**  
    1. 파일을 업로드합니다 (PPT, PPTX, PDF만 가능합니다).  
    2. 파일 순서를 Drag-and-Drop으로 조정합니다.  
    3. '결합하기' 버튼을 클릭하여 하나의 PDF로 결합합니다!  
    """
)

# File Upload
uploaded_files = st.file_uploader(
    "📤 파일을 업로드하세요 (PPT, PPTX, PDF 지원)", 
    type=["ppt", "pptx", "pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    # 기본적으로 파일 이름을 오름차순 정렬
    filenames = sorted([file.name for file in uploaded_files])
    
    # Drag-and-Drop으로 파일 순서 조정
    st.write("### 파일 순서를 Drag-and-Drop으로 조정하세요:")
    sorted_filenames = sort_items(filenames)

    # 결과 출력
    st.write("#### 선택된 순서:")
    st.write(sorted_filenames)

    # 결합된 파일의 이름 설정: 첫 번째 파일 이름 + "(결합)"
    default_output_name = f"{Path(sorted_filenames[0]).stem}(결합).pdf"
    output_name = st.text_input(
        "📁 저장할 파일 이름을 입력하세요", 
        value=default_output_name
    )

    # Merge files
    if st.button("결합하기"):
        # Temporary directory to store converted files
        temp_dir = Path("temp_files")
        temp_dir.mkdir(exist_ok=True)

        merger = PdfMerger()
        for sorted_filename in sorted_filenames:
            file = next(f for f in uploaded_files if f.name == sorted_filename)
            if file.type == "application/pdf":
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
