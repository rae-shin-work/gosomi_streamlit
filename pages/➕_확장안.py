import streamlit as st
import openai
import streamlit as st
from docx import Document
from docx import shared
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
from PIL import Image
import os
import io
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeResult
import numpy as np
from st_audiorec import st_audiorec
from urllib.error import URLError
import time

st.set_page_config(page_title="확장", page_icon="➕")

st.markdown("# 확장안")
st.sidebar.header("확장안")
st.write("고소미 서비스는 멀티 모달 AI 서비스로 확장 가능합니다.")
st.write("기존 이미지 뿐만 아니라 음성 및 영상으로 사용자로부터 입력받는 확장안을 시연하겠습니다.")

st.markdown("##### :cop: 소장 작성을 위해 정보를 제공해주세요.")
with st.form("acc_info"):
    st.write("사건 정보를 수집하겠습니다.")
    acc_date = st.date_input("언제 발생하셨나요?")
    st.markdown(
        """<style> div[class*="stWidgetLabel"] > label > div[data-testid="stMarkdownContainer"] > p {
        font-size: 20px;} </style>
        """,
        unsafe_allow_html=True,
    )

    st.write("육하원칙으로 고소하려는 사건에 대해 구체적으로 묘사해주세요.")
    wav_audio_data = st_audiorec()

    # if wav_audio_data is not None:
    #     st.audio(wav_audio_data, format='audio/wav')

    acc_info_submitted = st.form_submit_button("사건 정보 작성 완료")

if acc_info_submitted:
        with st.spinner('AI를 통해 음성 분석 중 입니다...'):
            time.sleep(10)

        st.success("""### :cop: 음성으로 사건을 파악했습니다.
    입력해주신 음성은 베트남어로 아래와 같이 인식되었습니다.
    "Ngày 20 tháng 4 năm 2024, tôi đã để quên máy tính xách tay trong phòng họp.
    Tôi đã hỏi cô ấy về máy tính xách tay, nhưng cô ấy đã chửi mắng thậm tệ thông qua Ka-ka-o-Tok.
    Tâm trạng tôi rất xấu và tức giận.
    Hy vọng cô ấy sẽ bị trừng phạt."
                   
    음성에서 파악된 사건에 대한 정보는 아래와 같습니다.
    *   2024년 4월 20일, 저는 노트북을 교실에 두고 왔습니다. 
    *   노트북에 대해 물었더니 카카오톡으로 욕을 많이 먹었습니다. 
    *   나는 기분이 매우 나빴고 화가 났습니다. 그녀가 벌을 받았으면 좋겠습니다.""")

try:
    st.write("")
    st.markdown("##### :cop: 증거가 있다면 업로드 해주세요.")

    video_file_buffer = st.file_uploader("Upload a video file", type="mp4")
    if video_file_buffer is not None:
        # 문서 분석 API 호출
        st.video(video_file_buffer, format="video/mp4", start_time=0, subtitles=None, end_time=None, loop=False)
        # with video_file_buffer as f:
        #     poller = document_intelligence_client.begin_analyze_document(
        #         "prebuilt-layout",
        #         analyze_request=f,
        #         content_type="application/octet-stream",
        #     )
        # result: AnalyzeResult = poller.result()

        # diaglog = []

        # def _in_span(word, spans):
        #     for span in spans:
        #         if word.span.offset >= span.offset and (
        #             word.span.offset + word.span.length
        #         ) <= (span.offset + span.length):
        #             return True
        #     return False

        # def get_words(page, line):
        #     result = []
        #     for word in page.words:
        #         if _in_span(word, line.spans):
        #             result.append(word)
        #     return result

        # for page in result.pages:
        #     if page.lines:
        #         for line_idx, line in enumerate(page.lines):
        #             words = get_words(page, line)
        #             # print(f"...Line # {line_idx} test: '{line.content}' ")
        #             if line.content != "1":
        #                 diaglog.append(line.content)

        #         # st.write(diaglog)

        # result1 = openai.chat.completions.create(
        #     model="gpt-35-turbo-001",
        #     temperature=1,  # 창의적으로 답변하도록 최대치인 1로 수정
        #     messages=[
        #         {"role": "system", "content": "법적인 고소장 양식에 맞게 요약해줘."},
        #         {
        #             "role": "user",
        #             "content": str(diaglog)
        #             + """구체적으로 어떤 욕설을 얼마나 했는지 정량적인 수치와 함께 한 문장으로 요약해줘.
        #         욕설의 예시로는 지랄, ㅂㅅ, 병신, 존나, 좃밥, ㅈ밥, 시발, ㅅㅂ 등이 있어.""",
        #         },
        #     ],
        # )
        # st.write(result1.choices[0].message.content)
        with st.spinner('AI를 통해 영상 분석 중 입니다...'):
            time.sleep(10)

        st.success("""## :cop: CCTV 영상 분석 결과
    # **시간:** 영상 시작 시점부터 약 9초간

    # **장소:** 회의실로 추정되는 장소

    # **대상:** 검은색 모자, 줄무늬 셔츠, 검은색 바지를 입은 사람

    # **행동:**

    *   대상은 문을 통해 회의실 안으로 들어옴.
    *   회의실 안쪽을 살핀 후 테이블 위에 있는 서류를 집어 듦.
    *   서류를 여러 장 넘겨보며 내용을 확인.
    *   서류를 다시 테이블 위에 놓고 노트북을 집어 듦.
    *   노트북을 가방에 넣음.
    *   회의실 안쪽으로 들어감.
    *   잠시 후 다시 회의실 안에서 나와 문을 통해 퇴장. 

    # **추가정보**

    *   대상은 마스크와 모자를 착용하고 있어 얼굴 식별이 어려움.
    *   영상에서 대상 이외의 다른 사람은 등장하지 않음. 
    *   대상이 가져간 것으로 확인되는 물품은 노트북 1대. 
    *   서류의 내용 및 노트북의 소유자는 알 수 없음. 
    *   대상의 행동이 절도인지 여부는 판단 불가.
        """)
        
except URLError as e:
    st.error(
        """
        **This demo requires internet access.**
        Connection error: %s
    """
        % e.reason
    )