from pyrfc import Connection
import pandas as pd

# SAP 접속 정보 (실제 정보로 교체 필요)
sap_conn_params = {
    'user': 'pop_inf',
    'passwd': 'interface',
    'ashost': 'vhkrclopcs.sap.korloy.com',
    'sysnr': '00',
    'client': '100',
    'lang': 'KO'
}

# SAP 연결
conn = Connection(**sap_conn_params)

# 조회 기간 설정
start_date = '20250721'
end_date = '20250725'

# RFC 함수 호출 (ZMMR012는 예시 이름)
result = conn.call('ZMMR012', I_BUDAT_LOW=start_date, I_BUDAT_HIGH=end_date)

# 결과 테이블 추출
pr_data = result.get('ET_PR_LIST', [])

# DataFrame으로 변환 후 엑셀 저장
df = pd.DataFrame(pr_data)
df.to_excel('구매요청_2025년7월21일_25일.xlsx', index=False)
