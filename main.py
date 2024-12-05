import streamlit as st
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pdf2image import convert_from_bytes
from io import BytesIO
from pathlib import Path
from streamlit_sortables import sort_items

# 앱 제목
st.title("📎 PDF 병합 & 분할 도구 (By 석리송)")

# 업로드 파일 수집
uploaded_files = st.file_uploader(
    "📤 병합할 PDF 파일을 업로드하세요 (Drag-and-Drop 가능)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    filenames = [file.name for file in uploaded_files]

    # 기본 파일 순서 정렬
    filenames = sorted(filenames)

    # PDF 파일의 첫 페이지를 이미지로 변환하여 미리보기
    st.write("### 업로드된 PDF 미리보기")
    for file in uploaded_files:
        first_page_image = convert_from_bytes(file.read(), first_page=1, last_page=1)[0]
        st.write(f"**{file.name}** - 미리보기")
        st.image(first_page_image, caption=f"첫 페이지 미리보기 - {file.name}", width=400)

    # 파일 순서 변경
    st.write("### 파일 순서를 Drag-and-Drop 또는 입력으로 변경하세요:")
    sorted_filenames = sort_items(filenames)
    custom_order = st.text_input(
        "📋 파일 순서를 쉼표로 구분하여 입력하세요 (예: 2,1,3):",
        value=",".join(map(str, range(1, len(sorted_filenames) + 1))),
    )
    try:
        indices = list(map(int, custom_order.split(",")))
        sorted_filenames = [filenames[i - 1] for i in indices]
    except (ValueError, IndexError):
        st.error("순서 입력이 잘못되었습니다. 올바른 숫자를 쉼표로 구분해 입력해주세요.")

    st.write("#### 최종 파일 순서:")
    st.write(sorted_filenames)

    # 출력 파일 이름 설정
    pdf_output_name = st.text_input("📁 병합된 PDF 파일 이름", value="merged.pdf")

    # 병합 및 다운로드
    if st.button("📥 PDF 병합"):
        try:
            merger = PdfMerger()
            for filename in sorted_filenames:
                file = next(f for f in uploaded_files if f.name == filename)
                merger.append(BytesIO(file.read()))

            # 병합 결과 제공
            with BytesIO() as buffer:
                merger.write(buffer)
                st.download_button(
                    "📥 병합된 PDF 다운로드",
                    data=buffer.getvalue(),
                    file_name=pdf_output_name,
                    mime="application/pdf",
                )
            merger.close()
        except Exception as e:
            st.error(f"PDF 병합 중 오류 발생: {e}")

else:
    st.warning("최소 하나의 PDF 파일을 업로드해주세요.")
