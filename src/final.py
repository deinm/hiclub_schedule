from tkinter import *
import xlrd
import random
import xlsxwriter
from tkinter import messagebox

def call():
    create(var2.get(),var.get(),int(playField.get()))

def sel1():
    print(var2.get())
    selection = str(var2.get())
    label1.config(text = selection)

def sel2():
    selection2 = "한 타임 " + str(var.get())+ " 명"
    label2.config(text = selection2)

def CallBack():
   messagebox.showinfo("Success", "생성이 완료되었습니다.")

def getBGColor(book, sheet, row, col):
    xfx = sheet.cell_xf_index(row, col)
    xf = book.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    pattern_colour = book.colour_map[bgx]

    return pattern_colour

def create(campus, min_people, campus_people):
    # Create an new Excel file and add a worksheet.
    new_workbook = xlsxwriter.Workbook('new_schedule.xlsx')
    worksheet = new_workbook.add_worksheet('최종 헬데 시간표')
    worksheet1 = new_workbook.add_worksheet('월요일')
    worksheet2 = new_workbook.add_worksheet('화요일')
    worksheet3 = new_workbook.add_worksheet('수요일')
    worksheet4 = new_workbook.add_worksheet('목요일')
    worksheet5 = new_workbook.add_worksheet('금요일')

    daily_sheet = [worksheet1, worksheet2, worksheet3, worksheet4, worksheet5]

    # 가로 폭 넓히기
    worksheet.set_column(0, 5, 20)
    worksheet1.set_column(0, 5, 20)
    worksheet2.set_column(0, 5, 20)
    worksheet3.set_column(0, 5, 20)
    worksheet4.set_column(0, 5, 20)
    worksheet5.set_column(0, 5, 20)

    center = new_workbook.add_format({'align': 'center'})

    cell_format = new_workbook.add_format({
        'border':1,
        'align': 'center',
        'valign' : 'vcenter'})

    warn_format = new_workbook.add_format({
        'bg_color':   '#FFC7CE',
        'font_color': '#9C0006'})

    basic_format = new_workbook.add_format({
        'bg_color':   '#FFEB9C',
        'font_color': '#9C6500'})

    more_format = new_workbook.add_format({
        'bg_color':   '#C6EFCE',
        'font_color': '##006100'})

    date = ['월','화','수','목','금']
    point = [1,2,3,4,5]
    help_time = ['10:15~11:45', '11:45~13:15', '13:15~14:45', '14:45~16:30','']
    col_alpha = ['B','C','D','E','F']
    row_num = ['2','3','4','5']
    cnt_value = ['총횟수']

    worksheet1.write_column('A2',help_time, center)
    worksheet2.write_column('A2',help_time, center)
    worksheet3.write_column('A2',help_time, center)
    worksheet4.write_column('A2',help_time, center)
    worksheet5.write_column('A2',help_time, center)
    help_time.append('비고')

    worksheet1.write_row('B1', point, center)
    worksheet2.write_row('B1', point, center)
    worksheet3.write_row('B1', point, center)
    worksheet4.write_row('B1', point, center)
    worksheet5.write_row('B1', point, center)

    while True:
        schedule = xlrd.open_workbook("schedule.xls", formatting_info=True)
        # Schedule = book

        sheets = schedule.sheet_names()
        print("sheets are:", sheets)

        campus_num=0

        if campus=="율전":
            campus_num = 1
        elif campus=="명륜":
            campus_num = 0

        sheet1 = schedule.sheet_by_name(sheets[campus_num])

        # nrow = sheet1.nrows
        # ncol = sheet1.ncols
        # npeople = ncol - 1 # 19
        # print(nrow,ncol)

        # nrow : 196 ncol : 9

        nrow = 196
        ncol = campus_people + 1
        npeople = campus_people

        cnt = 0
        people = []
        single_person = []

        # key : 이름, value : 헬데 횟수
        people_cnt = {}

        # key : 이름, value : 우선순위 포인트
        people_point = {}

        # 세로로 체크
        for y in range(0, ncol):
            for cell in sheet1.col(y):
                print(cell.value, end=" ")
                background_color = getBGColor(schedule, sheet1, cnt, y)

                # 빈칸이 아닐 경우 값 자체 저장
                # 일정 있는 시간 1, 아니면 0
                if cell.value != "":
                    print("TYPE",type(cell.value))
                    single_person.append(cell.value)
                elif background_color != None:
                    single_person.append(1)
                else:
                    single_person.append(0)

                print(background_color)
                cnt += 1

            people.append(single_person)
            single_person = []
            cnt = 0
            print()

        single_point = []

        # 헬데 횟수 count dictionary 생성
        # point dictionary 별도 생성
        for person in people[1:]:
            people_cnt[person[0]] = 0
            for i in range(0, 5):
                single_point.append(person[39 * i + 39])

            people_point[person[0]] = single_point
            single_point = []

        print(people_cnt)

        # n번째 타임에 설 수 있는 사람 = time[n]
        time = []
        single_time = []

        # 월 화 수 목 금(5일)
        for day in range(0, 5):
            # i번째 타임 슬사람
            for i in range(0, 4):
                # n명 결정
                # range(i, j) : i~j-1
                for person in people[1:]:
                    # 한 타임은 무조건 1시간 30분이라고 가정
                    # 39*day = 요일(하루당 셀 39개)
                    # 6*i+7 = 한 타임 1시간 30분
                    print(person[39 * day + 6 * i + 7:39 * day + 6 * i + 13], end=" ")
                    if 1 in person[39 * day + 6 * i + 7:39 * day + 6 * i + 13]:
                        print("False")
                    else:
                        single_time.append(person[0])
                        print("True")

                time.append(single_time)
                single_time = []

        # 가능한 사람 리스트
        print(time)

        day = 0

        find_min = {}
        total_points = []
        one = []
        two = []
        three = []
        four = []
        five = []

        for part_time in time:
            # part_time : 그 시간에 가능한 모든 사람 이름 list

            # 포인트는 요일별로 받으므로 그 요일에 대한 포인트 가져오기
            for person in part_time:
                print("TYPE",type(people_point[person][int(day / 4)]))

                if type(people_point[person][int(day / 4)]) != float:
                    people_point[person][int(day / 4)] = 1.0

                if people_point[person][int(day / 4)] == 1.0:
                    one.append(person)
                elif people_point[person][int(day / 4)] == 2.0:
                    two.append(person)
                elif people_point[person][int(day / 4)] == 3.0:
                    three.append(person)
                elif people_point[person][int(day / 4)] == 4.0:
                    four.append(person)
                elif people_point[person][int(day / 4)] == 5.0:
                    five.append(person)

            # find_min : 특정 타임에 헬데 설 수 있는 사람들 요일별 베팅포인트
            find_min[1] = one
            find_min[2] = two
            find_min[3] = three
            find_min[4] = four
            find_min[5] = five

            # dictionary 모은 list
            total_points.append(find_min)

            # 초기화
            find_min = {}
            one = []
            two = []
            three = []
            four = []
            five = []

            day += 1

        # 총 헬데 횟수 = 요일(5) * 타임(4) = 20
        # 총 인원수 = 15
        # 1인당 최소 1, 누군가는 2번

        # 최소 서야 하는 헬데 횟수
        # 4*5(일주일) * 한 타임당 헬데 인원 / 총 헬데 가능 인원
        min_help = int(20*min_people / npeople)
        print("MIN HELP",min_help)
        max_help = min_help+1

        print("MIN",min_help)
        print("PEOPLE",npeople)

        help_desk = []
        single_desk = []

        temp_col = ['B','C','D','E','F','G']
        temp_row = ['2','3','4','5']
        page = 0
        pagecount = 1

        for single_day in total_points:
            # 하나의 타임에 key 점수 value 이름 list
            print("DAY_Point", single_day)

            print("PAGE", page)
            print("PAGECOUNT",pagecount)

            for key, value in single_day.items():
                print("VALUE",value)
                print("KEY",key)
                print("PLACE,",temp_col[key-1]+str(pagecount+1))
                daily_sheet[page].write_string(temp_col[key-1]+str(pagecount+1), ','.join(value), center)

            ##print("DAILY_SHEET", daily_sheet[page])
            ##daily_sheet[page].write_string('A1',"!", center)

            # 1~6 key로 설정해 loop (6은 임의 부여한 우선순위)
            for i in range(1, 7):
                if i != 6:
                    # 목록에 있는 사람이 최소 횟수만큼 헬데를 이미 섰다면 우선순위 6을 부여
                    for member in single_day[i]:
                        # 이름 기록해야 함
                        # 황경원(5), 김수아(3), ...
                        # worksheet1.write_column('A2', single_day[i], center)
                        if people_cnt[member] >= min_help:
                            # 최대 횟수보다 많이 섰다면 아예 삭제
                            # 이 경우 마지막 날짜로 가면 헬데에 설 수 있는 사람이 없을 수 있음
                            # 우선순위 7 부여?
                            if people_cnt[member] > 2:
                                print(single_day[i])
                                print(people_cnt[member])
                                single_day[i].remove(member)
                                print(single_day[i])
                            elif 6 in single_day.keys():
                                single_day[6].append(member)
                            else:
                                single_day[6] = [member]

                    length = len(single_day[i])

                    # 원래 우선순위에서 삭제
                    for k in range(0, length):
                        member = single_day[i][length - 1 - k]
                        if 6 in single_day.keys():
                            if member in single_day[6]:
                                single_day[i].remove(member)

                print("The 6-point", single_day)

                # 헬데 한타임에 서야 하는 사람이 2라고 가정했을 때
                # 현재 뽑은 헬데 인원이 2보다 안되면 뽑기 진행
                if len(single_desk) < min_people:
                    try:
                        day_len = len(single_day[i])
                        # 전체 인원 수 - 현재 인원 수 보다 지원자가 적으면 다 넣어버림
                        if day_len <= min_people - len(single_desk):
                            for member in single_day[i]:
                                single_desk.append(member)
                                people_cnt[member] += 1

                        # 지원자가 2명 이상이라 랜덤픽 해야함
                        else:
                            # 지원자가 n명 이상이 될 때까지
                            while len(single_desk) < min_people:
                                # single_desk.append(random.choice(single_day[i]))
                                randnum = random.randint(0, len(single_day[i]) - 1)
                                item = single_day[i].pop(randnum)
                                single_desk.append(item)
                                people_cnt[item] += 1
                    except:
                        pass

            print("FIN", single_desk)
            print("Current count", people_cnt)

            help_desk.append(single_desk)
            single_desk = []

            if pagecount%4==0:
                page+=1
                pagecount=0

            pagecount+=1

        # 최대 헬데 - 최소 헬데 = 2 초과일 경우 다시 생성
        if max(list(people_cnt.values())) - min(list(people_cnt.values())) > 2 :
            print("TRYING",max(list(people_cnt.values())) - min(list(people_cnt.values())))
            continue
        else:
            break

        # 수동 거르기
        # if 0 in people_cnt.values() or max_help+1 in people_cnt.values() or max_help+2 in people_cnt.values() or max_help+3 in people_cnt.values():
        #     print("!")
        #     continue
        # else:
        #     break

    print(help_desk)
    print(people_cnt)

    for key,value in people_cnt.items():
        help_time.append(key)
        cnt_value.append(value)

    worksheet.write_column('A2', help_time, center)
    worksheet.write_row('B1', date, center)

    worksheet.write_column('B7', cnt_value, center)

    #19명 : B8 ~
    check_area = 8 + npeople
    area_string = 'B8:B'+str(check_area-1)
    worksheet.conditional_format(area_string, {'type':'cell',
                                        'criteria': '<',
                                        'value':    min_help,
                                        'format':   warn_format})

    worksheet.conditional_format(area_string, {'type':'cell',
                                        'criteria': '==',
                                        'value':    min_help,
                                        'format':   basic_format})

    worksheet.conditional_format(area_string, {'type':'cell',
                                        'criteria': '>',
                                        'value':    min_help,
                                        'format':   more_format})

    cnt = 0
    write_member = ""

    print(help_desk)

    for col in col_alpha:
        for row in row_num:
            new = col+row

            worksheet.write_string(new, ",".join(help_desk[cnt]), center)
            cnt+=1

    new_workbook.close()
    CallBack()
    print(campus)

display = Tk()
display.title('Hi-Club Help Desk generater')
display.geometry('500x450')

TitleLabel = Label(display, text="하이클럽 헬프데스크 생성기", font=("Helvetica", 22, "bold"))
TitleLabel.pack()

TitleLabel = Label(display, text="같은 폴더에 조사한 시간표 파일을 'schedule.xls'로 저장 후 실행해 주세요.", font=("Helvetica", 13, "bold"))
TitleLabel.pack()

insLabel = Label(display, text="1. 캠퍼스", font=("Helvetica", 17))
insLabel.pack(anchor = W)

var2 = StringVar()
R1 = Radiobutton(display, text="자연과학캠퍼스(율전)", variable=var2, value="율전",
                  command=sel1)
R1.pack( anchor = W )

R2 = Radiobutton(display, text="인문사회과학캠퍼스(명륜)", variable=var2, value="명륜",
                  command=sel1)
R2.pack( anchor = W )

imsiLabel = Label(display, text="")
imsiLabel.pack(anchor = W)

ins2Label = Label(display, text="2. 캠퍼스 총 인원 수", font=("Helvetica", 17))
ins2Label.pack(anchor = W)

playField = Entry(display, width=5)
# ~ , 함수
playField.bind("<Return>")
playField.pack(anchor = W)

imsi2Label = Label(display, text="")
imsi2Label.pack(anchor = W)

ins1Label = Label(display, text="3. 헬데스크 한 타임 인원 수", font=("Helvetica", 17))
ins1Label.pack(anchor = W)

var = IntVar()
R4 = Radiobutton(display, text="2", variable=var, value=2,
                  command=sel2)
R4.pack( anchor = W )

R5 = Radiobutton(display, text="3", variable=var, value=3,
                  command=sel2)
R5.pack( anchor = W )

R6 = Radiobutton(display, text="4", variable=var, value=4,
                  command=sel2)
R6.pack(anchor = W)

label1 = Label(display)
label1.pack()

label2 = Label(display)
label2.pack()

imsi3Label = Label(display, text="")
imsi3Label.pack(anchor = W)

button = Button(display, text ="Submit", width = 20, font=("Helvetica", 17), command = call)
button.pack()

display.mainloop()
