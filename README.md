# Hi-Club Help Desk Schedule Generater

 하이클럽 헬프데스크 시간표를 자동으로 생성해주는 프로그램입니다.

## How program works

매 학기 Google Docs를 활용하여 시간표를 조사합니다.
시간표의 경우 명륜, 율전으로 sheet를 구분하며, 형식 샘플은 schedule.xls를 참고바랍니다.

<img width="979" alt="time_sample" src="https://user-images.githubusercontent.com/41565118/53228957-3315c500-36c6-11e9-8355-b6fe62d15a29.png">
시간표 조사 시 하이클럽 멤버는 위와 같이 자신의 이름에 해당하는 열을 찾아 자신의 일정(수업, 알바 등)을 셀 배경색 변경을 통해 표시합니다.

**이번 학기부터는 우선순위를 통해 헬프데스크를 배정합니다.**
비고란에 **1~5 사이의 우선순위**를 적을 수 있으며, 1이 가장 높은 우선순위, 5가 가장 낮은 우선순위입니다.
즉, 동일한 시간대에 헬프데스크를 진행할 수 있는 인원이 5명이더라도 우선순위가 높은 사람부터
헬프데스크에 배정되므로 자신이 가장 헬프데스크를 서고 싶은 요일에 높은 우선순위를 배정해 주시면 됩니다.

**만약 우선순위를 작성해주지 않으실 경우 무조건 우선순위를 1이라고 가정하며, 자신이 원하지 않는 요일에 배정될 확률이 높아질 수 있습니다.**

이전 학기처럼 율전 수업이 있는 명륜 학생이 시간표 조사 시 비고란에 **율전**이라고 적을 경우, 프로그램 상으로는 이를 헬데를 설 수 없는 날이라고 인지하지 못합니다.
이런 상황의 경우 우선순위를 5로 배정하거나, 일정이 하루 종일 있는 것으로 간주하여 해당 요일의 셀 배경색을 모두 바꾸어 주시면 됩니다.

헬프데스크는 명륜의 경우 한타임에 3명, 율전의 경우 2명이 배정되며, 명륜과 율전 모두 한 사람이 서게 되는 총 헬데 횟수는 비슷하도록 배정할 예정입니다.
헬프데스크 배정 시간은 다음과 같습니다.

1. 1타임 : 10:15~11:45
2. 2타임 : 11:45~13:15
3. 3타임 : 13:15~14:45
4. 4타임 : 14:45~16:30

프로그램은 완벽하지 않습니다! 회장단이 시간표를 짤 때 편하게 짤 수 있도록 만든 것이기 때문에 시간표 생성 이후 회장단이 추가 조정할 수 있습니다.

## How code works

시간표 생성은 **헬데를 가장 많이 서는 사람의 횟수 - 헬데를 가장 적게 서는 사람의 횟수 > 2**일 때까지 실행됩니다.
즉, 두 수 간의 편차가 2보다 클 경우 시간표를 다시 생성하게 됩니다.
*헬데를 설 수 있는 시간이 하나도 없다고 표시한 사람이 있을 경우 프로그램이 무한 반복되는 오류가 날 수 있으나 그런 일은 없다고 가정하고 무시하였습니다.*

예시를 들어 설명하겠습니다.

### Case 1
***
월요일 1타임에 율전 헬프데스크를 설 수 있는 사람이 다음과 같다고 가정합니다.
: 앞의 숫자는 각자 배정한 우선순위이며 괄호 안의 숫자는 지금까지 그 사람이 헬프데스크에 배정받은 횟수입니다.

```
1 : A(0)  
2 : B(0), C(0), D(0)  
3 :   
4 :  
5 : E(0)  
```

현재 모든 사람이 헬프데스크에 배정받은 사람이 없기 때문에 이부분은 고려하지 않습니다.

율전 헬프데스크는 한타임에 총 2명을 배정하므로, 우선순위가 높은 A가 먼저 자동으로 배정됩니다.
이후 B,C,D는 동일한 우선순위를 배정하였으므로 랜덤으로 B,C,D 중 한명이 선택됩니다.
***

### Case 2
***
월요일 1타임에 명륜 헬프데스크를 설 수 있는 사람이 다음과 같다고 가정합니다.
: 앞의 숫자는 각자 배정한 우선순위이며 괄호 안의 숫자는 지금까지 그 사람이 헬프데스크에 배정받은 횟수입니다.

```
1 :  
2 : A(4)  
3 : B(4), C(0)  
4 : D(2)  
5 : E(0)  
```

명륜 헬프데스크는 한타임에 총 3명을 배정하며, 1인당 최소 3번, 최대 4번 헬프데스크를 서게 된다고 하였을 때
A는 이미 최대 횟수를 초과하였기 때문에 우선순위에서 삭제하지만,
그 시간대에 헬프데스크가 가능한 사람이 없을 경우 부득이하게 그 시간에 배정을 해야 하므로 임의로 우선순위 6을 부여합니다.
즉, 다음과 같은 상태가 됩니다.

```
1 :  
2 :   
3 : B(4), C(0)  
4 : D(2)  
5 : E(0)  
6 : A(4)  
```

B도 마찬가지로 이미 헬프데스크를 최대 횟수만큼 섰기 때문에 우선순위 6을 부여합니다.

```
1 :  
2 :   
3 : C(0)  
4 : D(2)  
5 : E(0)  
6 : A(4), B(4)  
```

C는 이전에 헬데에 배정받지 않았으며 현재 상태에서 우선순위가 가장 높기 때문에 C가 배정받게 되며,
차례대로 D와 E가 헬프데스크에 배정받게 됩니다.
***

## How to use
Google Docs에서 작성된 시간표 파일을 **xls**형식으로 다운받아 입력합니다.
시간표 생성 이후, xlsx 형식의 **new_schedule.xls** 파일이 생성되며, 총 6개의 Sheet가 생성됩니다.

### 첫번째 시트

<img width="769" alt="2019-02-22 11 46 39" src="https://user-images.githubusercontent.com/41565118/53249724-3166f400-36fc-11e9-9c64-ef3125a633c0.png">

각 타임별로 헬데에 배정받은 사람의 이름이 표시되며, A7부터 해당하는 셀에는 각각 인원의 헬데 총 배정 횟수가 표시됩니다.
적정 횟수의 경우 노란색, 부족할 경우 빨간색, 초과했을 경우 초록색으로 구분하여 표시합니다.

### 두번째 이후의 시트

<img width="766" alt="2019-02-22 11 54 52" src="https://user-images.githubusercontent.com/41565118/53250218-47c17f80-36fd-11e9-9017-b4581b26f94e.png">

요일이 총 5개이기 때문에 추가로 5개의 Sheet가 더 생성됩니다. 
임원진의 경우 추가로 시간표를 수정해야 일이 있을 수 있으므로, 참고용으로 생성된 Sheet이며 Sheet의 이름은 월요일, 화요일, 수요일, 목요일, 금요일입니다.

A열에 쓰인 10:15-11:45 등은 Sheet의 이름에 해당하는 요일의 헬프데스크 타임이며, 1-5라고 쓰인 열 아래에 적힌 이름은
그 시간에 헬프데스크 활동을 할 수 있는 인원들이 그 요일에 배정한 우선순위입니다.
즉 'C2'열을 통해, B라는 학생이 10:15~11:45 타임에 헬프데스크를 우선순위 2만큼 희망한다는 것을 알 수 있습니다.


###### 추후 추가할 기능 및 해결해야 하는 부분
1. 요일별로 우선순위를 적기 때문에 헬프데스크를 하루에 연달하 하는 상황 발생 가능
2. 에브리타임 URL 입력 시 자동으로 엑셀 형태로 변경
3. 프로그램이 현재 python 파일인데, 웹으로 접속 가능하게 변경
