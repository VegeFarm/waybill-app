# Waybill Autofill (1+2 -> 3)

- 1번 스마트스토어 엑셀은 **비밀번호 0000**으로 고정하여 복호화 후 읽습니다.
- 2번 엑셀의 (주문자/수령자/주소)로 매칭하여 운송장번호를 가져옵니다.
- 3번(발송처리) 형식(상품주문번호/배송방법/택배사/송장번호)으로 xlsx를 생성합니다.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
