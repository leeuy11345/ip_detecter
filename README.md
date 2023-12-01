# 주제
- 파일 검수 웹 서비스

# 내용
> 이메일 서비스
-  회사에서 파일을 공유 및 전송 할 때 개인정보(이메일), 주석 등이 있으면 이메일을 전송하여 알려준다.

> 고객사 로그파일 분석 및 보고서 작성 자동화
- 고객사의 로그 파일을 보고 분 단위를 기준으로 몇번 접속을 시도 했는지 확인한다.
- 가장 많이 접속한 IP는 CIRIMAL API를 사용해서 IP정보를 가져온다.
- 이를 결합해서 보고서 템플릿을 활용하여 보고서를 자동화 작성한다.
- 작성된 보고서를 PDF로 사용자가 다운로드 할 수 있게 한다.

# 수정 할 부분
- Merger.py Criminal함수  API KEY 기입
``` 
class Criminal:
    def __init__(self, ip_address):
        
        self.url = f"https://api.criminalip.io/v1/asset/ip/summary?ip={ip_address}"

        self.payload={}
        self.headers = {
            "x-api-key": "<본인의 CRIMINAL API KEY를 기입>"
        }
```

- Merger.py 메인 함수 부분
```
if __name__ == "__main__":
    received_email = "<본인 네이버이메일 주소 기입>@naver.com"
    path = 'uploads'
    MT = Monitering(path,received_email)
    MT.inspect_annotation("#")
```

- .env 파일을 만들어 자신의 이메일 아이디 비밀 번호를 작성한다.
```
SECRET_ID = "이메일 아이디"
SECRET_PASS = "이메일 비밀번호"
```
