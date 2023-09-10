# XLToJsonConverter

## 제작 기간 : 2023.08.27 ~ 2023.09.04

1. 개요
2. 파일 관련
3. 사용 방법

---

1. 개요

C# 리플렉션 정보를 사용하여 excel 데이터를 json 데이터로 변환해주는 프로그램입니다.

사용 방법을 참고하시어 excel 파일을 작성하면, 저장된 경로에 json 파일이 생성되며, 필요에 따라 XLToJsonConverter.cs에서 경로들을 수정하여 사용할 수 있습니다.

---

2. 파일 관련

dataOutlieFilePath : 어떤 데이터를 컨버팅할 것인지에 대한 파일 경로 / 기본적으로 

dataFilePath : excel 파일 데이터 파일 경로

jsonSavePath : 출력 결과로 나올 json 파일 데이터 경로

2.1 XLDataOutline.json
2.1.1 ObjectType : 컨버팅될 대상 클래스나 구조체 이름 / 기본 namespace로 Data를 달고 있어야하며, 그 이후로는 추가된 네임스페이스와 구조체 명으로 기술
2.1.2 XLFileName : 컨버팅될 대상 excel 파일
2.1.3 SheetName : 컨버팅될 대상 sheet의 이름
2.1.4 SaveFileName : 저장할 json 파일 이름
2.1.5 HeaderCount : 실제 데이터가 아닌 헤더들의 갯수
2.1.6 IsVerticalData : 수직 데이터 엑셀인지에 대한 값 / 예시로 작성된 Sheet2 참고

---

3. 사용 방법

### 경로가 수정되지 않았다는 가정하에 기술합니다.

