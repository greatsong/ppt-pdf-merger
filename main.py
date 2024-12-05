import streamlit as st
from PyPDF2 import PdfMerger
from pathlib import Path
from io import BytesIO
from streamlit_sortables import sort_items

# 앱 제목
st.title("📎 PDF 병합 도구 (By 석리송)")

# 파일 업로드
uploaded_files = st.file_uploader(
    "📤 병합할 PDF 파일을 업로드하세요", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    filenames = [file.name for file in uploaded_files]

    # 파일 이름 정렬
    filenames = sorted(filenames)

    # Drag-and-Drop으로 파일 순서 변경
    st.write("### 파일 순서를 Drag-and-Drop으로 변경하세요:")
    sorted_filenames = sort_items(filenames)
    st.write("#### 선택된 파일 순서:")
    st.write(sorted_filenames)

    # 출력 파일 이름 설정
    pdf_output_name = st.text_input(
        "📁 병합된 PDF 파일 이름", 
        value="merged.pdf"
    )

    # 결합 버튼
    if st.button("결합하기"):
        temp_dir = Path("temp_files")
        temp_dir.mkdir(exist_ok=True)

        try:
            # PDF 병합 처리
            merger = PdfMerger()
            for filename in sorted_filenames:
                file = next(f for f in uploaded_files if f.name == filename)
                merger.append(BytesIO(file.read()))

            # 병합 결과 저장
            pdf_output_path = temp_dir / pdf_output_name
            merger.write(pdf_output_path)
            merger.close()

            # 다운로드 버튼 제공
            with open(pdf_output_path, "rb") as f:
                st.download_button(
                    "📥 병합된 PDF 다운로드",
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
    st.warning("최소 하나의 PDF 파일을 업로드해주세요.")
