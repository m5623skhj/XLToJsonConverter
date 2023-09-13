# XLToJsonConverter

## 제작 기간 : 2023.08.27 ~ 2023.09.04

1. 개요
2. 파일 관련
3. 속성
4. 사용 방법

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

3. 속성

각 데이터 변수에 대해 속성을 부여할 수 있습니다. (단, 배열 타입에 대해서는 지원하지 않습니다.)

3.1 Required : 

해당 컬럼에 대해서 excel 파일에 기술된 모든 셀들이 반드시 채워져 있음을 보장합니다.
  
해당 컬럼에 대해서 비워져 있다면, 에러를 발생시킵니다.

3.2 Alias :

해당 컬럼 이름을 변경할 수 있습니다.
  
excel 파일에 Alias데 적용된 이름을 기술하면, 컨버팅시 Alias 이름을 컬럼을 검색합니다.
  
단, 원래 정의되었던 변수 이름은 사용하지 않게 됩니다.

3.3 Min / Max Value :

excel 파일에 해당 최소 / 최대값을 강제합니다.
  
숫자만을 대상으로 하며, MinValue 보다 작거나, MaxValue 보다 크면 에러를 발생시킵니다.

---

4. 사용 방법

### 경로가 수정되지 않았다는 가정하에 기술합니다.

빌드된 파일의 3폴더 위의 경로에 /Data, /Generated, OptionFile 폴더를 생성합니다.

/Data에는 컨버팅 대상이 되는 excel 파일을 위치시켜 놓습니다.

/OptionFile에는 컨버팅을 위한 정보들을 기술합니다. 기술 방법은 2에서 설명한 것과 같습니다.

필요한 데이터 구조체들을 작성하여 해당 프로젝트에 위치시키고, 빌드합니다.

이후 빌드된 프로그램을 실행시키면, /Generated에 json 파일들이 생성됩니다.

단, 에러가 하나라도 발생할 경우, 에러가 발생된 파일은 컨버팅되지 않습니다.

에러 확인은 컨버팅 작업이 완료된 후 발생된 모든 에러 리스트를 확인할 수 있습니다.

4.1 엑셀 기술 방법 :

excel 파일에는 하나의 ObjectType에 대한 기술을 합니다.

하나의 변수가 하나의 컬럽 이름이라고 생각하면 됩니다.

해당 변수가 구조체나 클래스 타입일 경우, 상위에 해당 변수의 이름을, 하위에는 구조체가 가지고 있는 변수들을 기술합니다. (단, 상위 변수는 병합해야 합니다.)

배열이 필요할 경우 List<T>로 기술하면 됩니다.

---

예시 1. 일반적인 형태의 시트

![image](https://github.com/m5623skhj/XLToJsonConverter/assets/42509418/da1012b8-7412-40b4-a7b9-3c533851cb03)

---

예시 2. 구조체를 포함하거나, 배열을 포함한 시트

![image](https://github.com/m5623skhj/XLToJsonConverter/assets/42509418/5f923a5a-2a58-41bc-a8d4-ed056f64f8a4)

---

예시 3. 수직으로 작성된 시트

![image](https://github.com/m5623skhj/XLToJsonConverter/assets/42509418/b17e58d1-a2dd-4b43-974f-150c3046a28d)
