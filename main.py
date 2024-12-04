import streamlit as st
from PyPDF2 import PdfMerger
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pathlib import Path
from io import BytesIO
from streamlit_sortables import sort_items

# 슬라이드 복사 함수 (XML 기반)
def copy_slide(presentation, slide):
    slide_element = slide._element
    new_slide_element = parse_xml(slide_element.xml)
    presentation.slides._sldIdLst.append(new_slide_element)

# Streamlit 앱
st.title("📎 PDF & PPTX 병합 도구")

# 파일 업로드
uploaded_files = st.file_uploader(
    "📤 PDF와 PPTX 파일을 업로드하세요 (Drag-and-Drop 가능)", 
    type=["pdf", "pptx"], 
    accept_multiple_files=True
)

if uploaded_files:
    filenames = [file.name for file in uploaded_files]
    filenames = sorted(filenames)

    # 파일 순서 변경
    st.write("### 파일 순서를 Drag-and-Drop으로 변경하세요:")
    sorted_filenames = sort_items(filenames)
    st.write("#### 선택된 파일 순서:")
    st.write(sorted_filenames)

    # 출력 파일 이름 설정
    pptx_output_name = st.text_input("📁 PPTX 결합 파일 이름", value="merged.pptx")
    pdf_output_name = st.text_input("📁 PDF 결합 파일 이름", value="merged.pdf")

    if st.button("결합하기"):
        temp_dir = Path("temp_files")
        temp_dir.mkdir(exist_ok=True)

        # PPTX 병합 처리
        try:
            merged_presentation = Presentation()  # 빈 프레젠테이션 생성
            for filename in sorted_filenames:
                file = next(f for f in uploaded_files if f.name == filename)
                if file.name.endswith(".pptx"):
                    presentation = Presentation(BytesIO(file.read()))
                    for slide in presentation.slides:
                        # 원본 슬라이드 크기 동기화
                        merged_presentation.slide_width = presentation.slide_width
                        merged_presentation.slide_height = presentation.slide_height
                        # 슬라이드 복사
                        copy_slide(merged_presentation, slide)

            # 저장
            pptx_output_path = temp_dir / pptx_output_name
            merged_presentation.save(pptx_output_path)
            with open(pptx_output_path, "rb") as f:
                st.download_button(
                    "📥 PPTX 병합 파일 다운로드",
                    f,
                    file_name=pptx_output_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        except Exception as e:
            st.error(f"PPTX 병합 중 오류 발생: {e}")

        # PDF 병합 처리
        try:
            merger = PdfMerger()
            for filename in sorted_filenames:
                file = next(f for f in uploaded_files if f.name == filename)
                if file.name.endswith(".pdf"):
                    merger.append(BytesIO(file.read()))
            pdf_output_path = temp_dir / pdf_output_name
            merger.write(pdf_output_path)
            merger.close()
            with open(pdf_output_path, "rb") as f:
                st.download_button(
                    "📥 PDF 병합 파일 다운로드",
                    f,
                    file_name=pdf_output_name,
                    mime="application/pdf"
                )
        except Exception as e:
            st.error(f"PDF 병합 중 오류 발생: {e}")

        # 임시 파일 정리
        for temp_file in temp_dir.iterdir():
            temp_file.unlink()
        temp_dir.rmdir()
else:
    st.warning("최소 하나의 PDF 또는 PPTX 파일을 업로드해주세요.")
