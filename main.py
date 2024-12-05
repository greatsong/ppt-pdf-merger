import streamlit as st
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pathlib import Path
from io import BytesIO
from streamlit_sortables import sort_items
import base64

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

    # 파일 미리보기 제공
    st.write("### 업로드된 PDF 미리보기")
    for file in uploaded_files:
        reader = PdfReader(BytesIO(file.read()))
        first_page = reader.pages[0]
        st.write(f"**{file.name}** - {len(reader.pages)} 페이지")
        with BytesIO() as buffer:
            writer = PdfWriter()
            writer.add_page(first_page)
            writer.write(buffer)
            st.image(buffer.getvalue(), caption=f"첫 페이지 미리보기 - {file.name}", width=400)

    # 파일 순서 변경
    st.write("### 파일 순서를 Drag-and-Drop 또는 입력으로 변경하세요:")
    sorted_filenames = sort_items(filenames)

    # 숫자 입력을 통한 순서 변경 추가
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

    # PDF 분할 기능
    st.write("### 추가 기능: PDF 분할")
    split_file = st.selectbox("📂 분할할 PDF 파일을 선택하세요:", filenames)
    if split_file:
        page_range = st.text_input(
            "📄 분할할 페이지 범위를 입력하세요 (예: 1-3,5):",
            value="1-3",
        )
        try:
            file = next(f for f in uploaded_files if f.name == split_file)
            reader = PdfReader(BytesIO(file.read()))
            writer = PdfWriter()

            # 페이지 범위 파싱
            ranges = page_range.split(",")
            for r in ranges:
                if "-" in r:
                    start, end = map(int, r.split("-"))
                    for i in range(start - 1, end):
                        writer.add_page(reader.pages[i])
                else:
                    writer.add_page(reader.pages[int(r) - 1])

            # 분할 파일 다운로드 제공
            with BytesIO() as buffer:
                writer.write(buffer)
                st.download_button(
                    "📥 분할된 PDF 다운로드",
                    data=buffer.getvalue(),
                    file_name=f"split_{split_file}",
                    mime="application/pdf",
                )
        except Exception as e:
            st.error(f"PDF 분할 중 오류 발생: {e}")

else:
    st.warning("최소 하나의 PDF 파일을 업로드해주세요.")
