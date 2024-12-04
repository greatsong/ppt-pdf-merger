import streamlit as st
from PyPDF2 import PdfMerger
from pptx import Presentation
from pathlib import Path
from io import BytesIO

# App Title
st.title("📎 PPT와 PDF 파일 결합 도구 (By 석리송)")

# File Upload
uploaded_files = st.file_uploader(
    "📤 파일을 업로드하세요 (PPT, PPTX, PDF 지원)", 
    type=["ppt", "pptx", "pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    # 파일을 PPT와 PDF로 분류
    ppt_files = [file for file in uploaded_files if file.type in ["application/vnd.ms-powerpoint", "application/vnd.openxmlformats-officedocument.presentationml.presentation"]]
    pdf_files = [file for file in uploaded_files if file.type == "application/pdf"]

    # 정렬: 파일 이름 오름차순
    ppt_files = sorted(ppt_files, key=lambda f: f.name)
    pdf_files = sorted(pdf_files, key=lambda f: f.name)

    # 사용자에게 선택된 파일 확인
    if ppt_files:
        st.write("### PPT 파일 목록")
        st.write([file.name for file in ppt_files])

    if pdf_files:
        st.write("### PDF 파일 목록")
        st.write([file.name for file in pdf_files])

    # 출력 파일 이름 설정
    if ppt_files:
        ppt_output_name = st.text_input(
            "📁 저장할 PPT 결합 파일 이름을 입력하세요 (기본값: merged.pptx)",
            value="merged.pptx"
        )
    if pdf_files:
        pdf_output_name = st.text_input(
            "📁 저장할 PDF 결합 파일 이름을 입력하세요 (기본값: merged.pdf)",
            value="merged.pdf"
        )

    # 결합 버튼
    if st.button("결합하기"):
        # 결합 결과 저장할 임시 디렉토리
        temp_dir = Path("temp_files")
        temp_dir.mkdir(exist_ok=True)

        # PPT 파일 결합
        if ppt_files:
            try:
                merged_presentation = Presentation()
                for ppt_file in ppt_files:
                    presentation = Presentation(BytesIO(ppt_file.read()))
                    for slide in presentation.slides:
                        # 슬라이드 복사
                        blank_slide_layout = merged_presentation.slide_layouts[6]
                        slide_copy = merged_presentation.slides.add_slide(blank_slide_layout)
                        for shape in slide.shapes:
                            el = shape.element
                            slide_copy.shapes._spTree.insert_element_before(el, 'p:extLst')
                # 저장
                ppt_output_path = temp_dir / ppt_output_name
                merged_presentation.save(ppt_output_path)
                with open(ppt_output_path, "rb") as f:
                    st.download_button(
                        "📥 PPT 결합 파일 다운로드",
                        f,
                        file_name=ppt_output_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"PPT 결합 중 오류 발생: {e}")

        # PDF 파일 결합
        if pdf_files:
            try:
                merger = PdfMerger()
                for pdf_file in pdf_files:
                    merger.append(pdf_file)
                # 저장
                pdf_output_path = temp_dir / pdf_output_name
                merger.write(pdf_output_path)
                merger.close()
                with open(pdf_output_path, "rb") as f:
                    st.download_button(
                        "📥 PDF 결합 파일 다운로드",
                        f,
                        file_name=pdf_output_name,
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"PDF 결합 중 오류 발생: {e}")

        # 임시 디렉토리 정리
        for temp_file in temp_dir.iterdir():
            temp_file.unlink()
        temp_dir.rmdir()
else:
    st.warning("최소 하나의 파일을 업로드해주세요.")
