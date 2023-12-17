### [외교부 월별세출집행현황] ###

외교부 월별 세출집행현황(2023년1월~2023년6월) 데이터를 추출하여 결과 파일 작성



**[사전작업]**

Input 폴더 : 기본엑셀양식, VBA 코드 파일

Config.xlsx : 사이트URL, 결과파일, 템플릿파일경로, 검색월 1월~6월



### [작업과정] ###

1. **INITIALIZE PROCESS**
   
     input 폴더의 템플릿파일 복사해서 결과파일 생성

     시트들을 LIST로 만들어 변수에 저장
   
2. **GET TRANSACTION DATA**

    Transaction Item 타입 : string  시트명리스트를 transactionitem으로 활용


3. **PROCESS TRANSACTION**

  - config파일에서 1월~6월을 읽어와 array 변수에 저장
    
  - 데이터 추출(2023년 1월~ 2023년 6월까지의데이터)

  - 템플릿에 양식에 맞는 데이터만 남겨서 엑셀 작성

  - '종합' 시트의 각 월별 합계 기입
 
  
4. **END PROCESS**
   
  - 작성한 날짜를 기준으로 연도와 월 기입
    
  - 1월~6월까지의 총합계 기입
