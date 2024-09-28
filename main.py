import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(layout="wide")

st.title("스케줄 도구")

# 파일 업로드와 그룹 설정을 위한 칼럼 생성
col1, col2 = st.columns(2)

with col1:
    # 파일 업로드
    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx", "xls"])

# 전역 변수 초기화
all_names = set()
schedule_data = []
sheet_data = []

if uploaded_file:
    # 엑셀 파일 읽기
    xl = pd.ExcelFile(uploaded_file)
    sheets = xl.sheet_names
    
    # 디버깅: 시트 수와 이름 출력
    #st.write(f"Debug: 엑셀 파일에서 {len(sheets)}개의 시트를 발견했습니다:")
    #for sheet in sheets:
    #    st.write(f"- {sheet}")

    # 모든 시트의 데이터를 처리
    for sheet in sheets:
        df = xl.parse(sheet)
        #st.write(f"Debug: 시트 '{sheet}' 처리 중...")

        # 주의 시작 날짜 추출 (C2 셀)
        start_date = df.iloc[1, 2]
        if not isinstance(start_date, datetime):
            try:
                start_date = datetime.strptime(str(start_date), "%d/%m/%Y")
            except ValueError:
                st.error(f"시작 날짜 형식을 인식할 수 없습니다: '{start_date}' (시트: {sheet})")
                continue

        #st.write(f"Debug: 시트 '{sheet}'의 시작 날짜: {start_date}")

        # 날짜 리스트 생성 (C3 ~ I3)
        dates = [start_date + timedelta(days=i) for i in range(7)]

        # 시간 리스트 생성 (B5 ~ B38)
        times = df.iloc[4:38, 1].tolist()

        # 스케줄 매트릭스 (C5 ~ I38)
        schedule_matrix = df.iloc[4:38, 2:9]

        sheet_schedule_data = []

        # 데이터 구조화
        for i, time_value in enumerate(times):
            if pd.isna(time_value):
                continue

            # 시간 값 처리
            if isinstance(time_value, datetime):
                time_obj = time_value.time()
            else:
                time_str = str(time_value).strip()
                time_formats = [
                    '%H:%M',
                    '%H:%M:%S',
                    '%I:%M %p',
                    '%I:%M:%S %p',
                    '%Y-%m-%d %H:%M:%S'
                ]
                for fmt in time_formats:
                    try:
                        parsed_time = datetime.strptime(time_str, fmt)
                        time_obj = parsed_time.time()
                        break
                    except ValueError:
                        continue
                else:
                    st.error(f"시간 형식을 인식할 수 없습니다: '{time_str}' (시트: {sheet})")
                    continue

            for j, date in enumerate(dates):
                cell = schedule_matrix.iloc[i, j]
                if pd.isna(cell):
                    continue
                names = [name.strip() for name in str(cell).split(",")]

                for name in names:
                    name = name.strip()
                    # 이름이 빈 문자열이거나 '_x001e_'인 경우 제외
                    if not name or name == '_x001E_' or name == '_x001e_':
                        continue
                    all_names.add(name)
                    sheet_schedule_data.append({
                        "name": name,
                        "datetime": datetime.combine(date.date(), time_obj)
                    })

        schedule_data.extend(sheet_schedule_data)
        sheet_data.append({
            "sheet_name": sheet,
            "start_date": start_date,
            "dates": dates,
            "times": times,
            "schedule_data": sheet_schedule_data
        })

        #st.write(f"Debug: 시트 '{sheet}' 처리 완료. 현재까지 {len(all_names)}명의 인물과 {len(schedule_data)}개의 일정을 추출했습니다.")

    # 인물 리스트 정렬
    all_names = sorted(list(all_names))
    #st.write(f"Debug: 총 {len(all_names)}명의 인물과 {len(schedule_data)}개의 일정을 추출했습니다.")

    with col2:
        st.subheader("그룹 생성")

        # 그룹 수 선택
        group_count = st.number_input("생성할 그룹의 수를 입력하세요", min_value=1, step=1)

    # 그룹 설정을 위한 칼럼 생성
    group_cols = st.columns(group_count)

    groups = {}
    for i, col in enumerate(group_cols):
        with col:
            group_name = st.text_input(f"그룹 {i+1}의 이름", value=f"그룹 {i+1}", key=f"group_name_{i}")
            selected_names = st.multiselect(
                f"{group_name}에 포함될 인물",
                options=all_names,
                key=f"group_members_{i}"
            )
            groups[group_name] = selected_names

    if st.button("가능한 시간 확인"):
        for sheet in sheet_data:
            st.subheader(f"{sheet['sheet_name']} ({sheet['start_date'].strftime('%Y-%m-%d')} ~ {sheet['dates'][-1].strftime('%Y-%m-%d')})")

            # 각 그룹의 결과를 표시할 칼럼 생성
            result_cols = st.columns(len(groups))

            for (group_name, group_members), col in zip(groups.items(), result_cols):
                with col:
                    st.write(f"{group_name}의 가능한 시간")

                    # 그룹 멤버들의 바쁜 시간 추출
                    busy_times = [entry["datetime"] for entry in sheet['schedule_data'] if entry["name"] in group_members]

                    # 전체 시간 생성
                    total_times = []
                    time_slots = []
                    for time_value in sheet['times']:
                        if pd.isna(time_value):
                            continue
                        # 시간 값 처리 (앞에서 사용한 로직 재사용)
                        if isinstance(time_value, datetime):
                            time_obj = time_value.time()
                        else:
                            time_str = str(time_value).strip()
                            time_formats = [
                                '%H:%M',
                                '%H:%M:%S',
                                '%I:%M %p',
                                '%I:%M:%S %p',
                                '%Y-%m-%d %H:%M:%S'
                            ]
                            for fmt in time_formats:
                                try:
                                    parsed_time = datetime.strptime(time_str, fmt)
                                    time_obj = parsed_time.time()
                                    break
                                except ValueError:
                                    continue
                            else:
                                continue  # 시간 파싱 실패 시 해당 시간 건너뜀

                        time_slots.append(time_obj)

                        for date in sheet['dates']:
                            datetime_obj = datetime.combine(date.date(), time_obj)
                            total_times.append(datetime_obj)

                    # 가능한 시간 계산
                    available_times = set(total_times) - set(busy_times)
                    available_times = sorted(list(available_times))

                    # 주간 달력 데이터프레임 생성
                    date_strs = [date.strftime("%Y-%m-%d") for date in sheet['dates']]
                    calendar_df = pd.DataFrame(index=[t.strftime("%H:%M") for t in time_slots], columns=date_strs)

                    # 초기값 설정 (가능으로 설정)
                    calendar_df = calendar_df.fillna('가능')

                    # 바쁜 시간 표시
                    for busy_time in busy_times:
                        date_str = busy_time.strftime("%Y-%m-%d")
                        time_str = busy_time.strftime("%H:%M")
                        if time_str in calendar_df.index and date_str in calendar_df.columns:
                            calendar_df.at[time_str, date_str] = '불가능'

                    # 스타일 적용 (색상 강조)
                    def highlight_cell(cell):
                        if cell == '불가능':
                            return 'background-color: lightcoral; color: white;'
                        else:
                            return 'background-color: lightgreen;'

                    styled_df = calendar_df.style.map(highlight_cell)
                    st.write(styled_df.to_html(), unsafe_allow_html=True)

            st.write("---")  # 각 시트 결과 사이에 구분선 추가