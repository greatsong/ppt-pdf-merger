import streamlit as st
from PyPDF2 import PdfMerger
from pptx import Presentation
from pathlib import Path
from io import BytesIO

# App Title
st.title("📎 파일 결합 도구 (By 석리송)")

# File Upload
uploaded_files = st.file_uploader(
    "📤 파일을 업로드하세요 (PPT, PPTX, PDF 지원)", 
    type=["ppt", "pptx", "pdf"], 
    accept_multiple_files=True
)

# 파일 업로드가 이루어진 경우만 처리
if uploaded_files:
    # 기본적으로 파일 이름을 오름차순으로 정렬
    filenames = sorted([file.name for file in uploaded_files])

    # Drag-and-Drop 또는 멀티셀렉트를 활용한 순서 설정
    st.write("### 파일 순서를 설정하세요:")
    filenames_sorted = st.multiselect(
        "파일 순서를 선택하세요:", filenames, default=filenames
    )

    if not filenames_sorted:
        st.warning("파일 순서를 설정해주세요!")
    else:
        st.write("선택된 파일 순서:", filenames_sorted)

        # 결합된 파일의 이름 설정
        default_output_name = f"{Path(filenames_sorted[0]).stem}(결합).pdf"
        output_name = st.text_input(
            "📁 저장할 파일 이름을 입력하세요", 
            value=default_output_name
        )

        # 결합 작업
        if st.button("결합하기"):
            temp_dir = Path("temp_files")
            temp_dir.mkdir(exist_ok=True)

            merger = PdfMerger()
            for sorted_filename in filenames_sorted:
                file = next(f for f in uploaded_files if f.name == sorted_filename)
                try:
                    if file.type == "application/pdf":
                        merger.append(file)
                    elif file.type in ["application/vnd.openxmlformats-officedocument.presentationml.presentation"]:
                        presentation = Presentation(BytesIO(file.read()))
                        pdf_path = temp_dir / f"{file.name}.pdf"
                        presentation.save(pdf_path)
                        merger.append(str(pdf_path))
                    else:
                        st.error(f"지원되지 않는 파일 형식: {file.name}")
                except Exception as e:
                    st.error(f"파일 '{file.name}' 처리 중 오류 발생: {e}")
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
else:
    st.warning("최소 하나의 파일을 업로드해주세요.")
