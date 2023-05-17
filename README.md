![header](https://capsule-render.vercel.app/api?type=waving&color=gradient&height=200&section=header&text=Excel%20VBA&fontSize=50)

## _Trying to handle Visual Basic Analysis efficiently_ 📋

<img src="https://img.shields.io/badge/Excel&nbsp;VBA-217346?style=for-the-badge&logo=Microsoft&logoColor=white">
<br/>

_Ch0. Why VBA?_<br/>
<br/>
😀 데이터 분석이 왜 중요한가?
데이터 분석을 통해 인사이트 도출 -> 기업 가치창출, 의사결정
- 양질의 데이터 플랫폼 다수 등장
- 빅데이터 분석을 위한 소프트웨어 등장(ex 테블로나 엑셀 파워쿼리)
- 데이터 분석을 통해 의사결정, 설득, 조직변경을 함 

<br/>

😀 시각적분석 : 태블로 , Power Bi
1. 시각화?
2. 현재 현업의 많은 데이터들이 엑셀의 수식(함수)을 사용해서 설계 > 데이터베이스화 X > 새로운 보고서 작성X
3. 가장 어려운 점 : 데이터 구조화
<br/><br/>

😀 Data Digital Transformation -> RPA시대☘️☘️☘️<br/>
&nbsp; &nbsp; &nbsp; &nbsp;-> 자동화를 하려고 노력해야함, 가지고있는 Tool로도 자동화 할 줄 알아야함
<br/><br/>
<br/>
😀 데이터를 다룰 때의 태도는? 항상 데이터를 의심하고 검증하라(데이터가 분석하기에 바람직 한가?)<br/>
1. 데이터가 정형화데이터인지 비정형데이터인지<br/>
2. 일반 데이터 범위인지 표개체인지 확인하기<br/>
	일반 데이터 범위를 표개체로 바꿔야함 [삽입 -> 표(CTRL + T), 표로 바뀐 후 표도구 있는지 확인]<br/>
	표개체를 일반데이터 범위로 바꾸는 방법 [표도구 -> 일반데이터변환]<br/>
		1. 표개체<br/>
			- 동적 범위로 자동으로 잡아준다<br/>
			- 필드 단위 계산: 구조적 참조<br/>
3. 데이터가 구조화 되어있는지 확인하기(별표5개)<br/>
	헤더가 셀병합이 안되어있어야함<br/>
	세로방향으로 데이터 쌓여있기<br/>
	중복있어도 비워놓지말고 채우기<br/>
4. 필드별 데이터 타입이 정확하게 잡혀있는지<br/>
	숫자 문자 날짜 수식 中 하나<br/>
<br/>

😀 코딩 시 : 동료의 관점에서 디자인하고 코딩할 것
1. 안정성 고려
2. 처리속도 올리기
   - with 문
   - 변수 활용
   - 배열변수 활용
   - 엑셀 기능 활용
3. 로직 설계
   - 결과물을 그림으로 그려봐라
   - 반복되는 패턴 : 행 반복? 열 반복?
4. 동료의 엑셀 라이프 스타일까지 고려해서 코딩
<br/>

😀 양질의 데이터 수집
- KOSIS
- 서울열린데이터광장
- 부산광역시 빅데이터플랫폼
- 빅카인즈/썸트렌드
- Kaggle
- Gapminder
- KDX 한국데이터거래소
- DataWorld(가상데이터)



😀 통계 분석<br/>
예시: KOSIS 자료, 스트레스 인지율과 자살 시도율이 서로 상관관계가 있는지 알고 싶다.<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;만약에 이 두 변수가 상관관계가 있다면 스트레스 인지율 변화에 따른 자살 시도울을 예측하고 싶다.<br/>
      
1. 준비단계<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;개발도구 리본을 꺼내고 -> 엑셀 추가기능 -> 분석도구 팩 (그러면 데이터 리본에 데이터 분석으로 들어옴)<br/>
2. 산포도<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;범위잡고 삽입에 분산형 차트 내림버튼 그중에 산포도<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;추세선 추가로 수식을 차트에 표시(회귀식), R^2값 표시(결정계수)<br/>
3. 상관분석<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;데이터 리본 -> 데이터분석에 상관분석<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;1-<상관계수<1<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;-1에 가까울수록 음의 상관관계<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;1에 가까울수록 양<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;0은 아무 관련 없음<br/>
4. 회귀분석<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;스트레스 인지율 변화에 따른(독립변수) -> 왼쪽에 적는게 국룰<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;자살 시도율을 예측해보고 싶다(종속변수) -> 오른쪽 배치<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;결정계수는 0<<1<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;1에 가까울수록 좋음 (0.935 이면 93%의 설명력을 가지고 있다)<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;파이썬에서는 r 제곱값 R^2<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;-단순회귀분석: 독립변수1개, 종속1개<br/> 
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;예시: 1인당 국민 총소득 변화에 따른 운전면허소지자 수를 과학적으로 예측<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;회귀식(회귀모형): y= ax + b 꼴<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;회귀식으로 다음 값을 예측 가능<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;-다중회귀분석: 독립변수 여러개, 종속변수 1<br/>
&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;-로지스틱...<br/>



[Ch1_PowerQuery](https://github.com/Yeom-Suji/Excel-VBA/blob/main/Ch1_PowerQuery)
[Ch2_Union&Join](https://github.com/Yeom-Suji/Excel-VBA/blob/main/Ch2_Union%26Join)
[Ch3_DataVisualization](https://github.com/Yeom-Suji/Excel-VBA/blob/main/Ch3_DataVisualization)
[Ch4_VBAMacros](https://github.com/Yeom-Suji/Excel-VBA/blob/main/Ch4_VBAMacros)
