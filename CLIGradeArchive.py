import sqlite3
import os
import re
import time
import openpyxl
import time

# DB의 path를 연결해 DB 객체 생성
db=sqlite3.connect('./students.db')

# ----- 성적 입력 -----
def Insert(student):
    # Window 환경에서 구동 시 clear -> cls로 변경 요함. 
    # pycharm terminal에서는 정상작동 하지 않을 수 있음
    os.system('clear')

    # 정규표현식 사용 : 10자리의 숫자 필터링
    id_regex = re.compile("[0-9]{10}")

    print('-성적 데이터 입력-\n')
    print(' * 입력시 주의 사항')
    print(' * 1. 학번형식 지켜서 입력')
    print(' *    학번은 입학년도 4자리 + 학과코드 3자리 + 학생코드 3자리 총 10자리로 구성됨 e.g. 2018036054')
    print(' * 2. 동명이인의 경우 이름 뒤 구분자 입력  e.g. 홍길동A, 홍길동B')
    print(" * 3. 성적은 국어, 영어, 수학 순으로 입력 / 구분자 ','\n")

    student_id = input('학번 입력 : ')
    validation = id_regex.search(student_id.replace(" ",""))

    # 학번 형식 체크
    if validation:
        # 학번 중복 체크
        for key, value in student.items():
            if student_id == value[0]:
                os.system('clear')
                print('중복된 학번입니다.\n')
                temp = input('기존 정보에 덮어쓰기 하시겠습니까? [Y/N] : ')

                # 대소문자 구분 X
                if temp == 'Y' or temp == 'y' :
                    os.system('clear')
                    print('\n * 입력시 주의 사항')
                    print(' * 1. 학생 이름은 관리자 확인 후 DB에서 수정바랍니다.')
                    print(" * 2. 성적은 국어, 영어, 수학 순으로 입력 / 구분자 ','")

                    name = key
                    kor, eng, math = map(int, input("성적 입력 : ").split(','))
                    score = kor + eng + math
                    avg = round(score/3, 3)

                    qry="update student set Score_Kor=?,Score_Eng=?,Score_Math=?,Score_Avg=? where Student_ID=?;"
                    try:
                        cur=db.cursor()
                        cur.execute(qry, (kor, eng, math, avg, student_id))
                        db.commit()
                        os.system('clear')
                        print("1행이 추가되었습니다.\n")
                        backToMain = input("아무키나 누르면 초기 메뉴로 돌아갑니다.")
                        return 
                    except sqlite3.ProgrammingError as e:
                        print("error in operation\n")
                        print(e)
                        time.sleep(5)
                        db.rollback()
                        db.close()
                    return

                elif temp == 'N' or temp == 'n':
                    backToMain = input("\n아무키나 누르면 초기 메뉴로 돌아갑니다.")
                    return

                else:
                    print('\n[Y/N] 중 1가지를 선택해주세요.\n')
                    print('3초 후 입력 메뉴로 돌아갑니다.')
                    time.sleep(3)
                    return Insert(student)

            else:
                continue

        name = input('이름 입력 : ')        
        kor, eng, math = map(int, input("성적 입력 : ").split(','))
        score = kor + eng + math
        avg = round(score/3, 3)

        qry="insert into student (Student_Name, Student_ID, Score_Kor, Score_Eng, Score_Math, Score_Avg) values(?,?,?,?,?,?);"
        try:
            cur=db.cursor()
            cur.execute(qry, (name, student_id, kor, eng, math, avg))
            db.commit()
            os.system('clear')
            print("1행이 추가되었습니다.\n")
            backToMain = input('아무키나 누르면 초기 메뉴로 돌아갑니다.')
            return 
        # 디버깅 편의를 위한 에러메시지 출력
        except sqlite3.ProgrammingError as e:
            print("error in operation\n")
            print(e)
            time.sleep(5)
            db.rollback()
            db.close()
    else:
        os.system('clear')
        print('잘못된 학번 형식입니다.\n')
        print('3초 후 입력 메뉴로 돌아갑니다.')
        time.sleep(3)
        return Insert(student)

    student[name] = [student_id, kor, eng, math, avg]
    return 


# ----- 학생 검색 -----
def Search(student): 
    RenderList(student)
    print('[ 1. 이름 ] [ 2. 학번 ] [3. 초기메뉴]')
    fieldSelect = int(input('검색할 필드를 선택해주세요. : '))
    
    if fieldSelect == 1:
        RenderList(student)
        print('이름으로 검색합니다.')
        keyForSearch = input('이름 입력 : ')

        # 검색한 이름이 딕셔너리내의 Key값과 일치하는 경우
        if(keyForSearch in student) == True:
            qry="select rowid from student where Student_Name=?;"
            try:
                cur=db.cursor()
                cur.execute(qry, (keyForSearch,))
                searchResult=cur.fetchall()

            except sqlite3.ProgrammingError as e:
                print("error in operation\n")
                print(e)
                time.sleep(5)
                db.rollback()
                db.close()

            RenderList(student)
            print(keyForSearch, ':' , student.get(keyForSearch),' ')        # 딕셔너리에서 키의 값을 가져옴 => dic.get(key)
            print('\n{} 행에서 검색되었습니다.'.format(searchResult[0][0]))
            print("=======================================================")
            backToMain = input('\n아무키나 누르면 초기 메뉴로 돌아갑니다.')
            return 

        else:
            print('해당 이름이 없습니다.\n')
            backToMain = input('아무키나 누르면 초기 메뉴로 돌아갑니다.')
            return 

    elif fieldSelect == 2:
        id_regex = re.compile("[0-9]{10}")
        RenderList(student)
        print('학번으로 검색합니다.')
        keyForSearch = input('학번 입력 : ')
        validation = id_regex.search(keyForSearch.replace(" ",""))
        
        #학번 형식 체크
        if validation:
            #학번 유무 체크
            for key, value in student.items():
                if keyForSearch == value[0]:
                    qry="select rowid, Student_Name from student where Student_ID=?;"
                    try:
                        cur=db.cursor()
                        cur.execute(qry, (keyForSearch,))
                        searchResult=cur.fetchall()

                    except sqlite3.ProgrammingError as e:
                        print("error in operation\n")
                        print(e)
                        time.sleep(5)
                        db.rollback()
                        db.close()

                    RenderList(student)

                    name = searchResult[0][1]
                    print(name, ':' , student.get(name),' ')        # 딕셔너리에서 키의 값을 가져옴 => dic.get(key)
                    print('\n{} 행에서 검색되었습니다.'.format(searchResult[0][0]))
                    print("=======================================================")
                    backToMain = input('\n아무키나 누르면 초기 메뉴로 돌아갑니다.')
                    return
                else:
                    continue
        
            print('해당 학번이 없습니다.\n')
            backToMain = input('아무키나 누르면 메뉴로 돌아갑니다.')
            return 
        else:
            print('잘못된 학번 형식입니다.\n')
            print('3초 후 검색메뉴로 돌아갑니다.')
            time.sleep(3)
            return Search(student)

    elif fieldSelect == 3:
        return 
        
    else:
        print('제시된 필드 중 선택해주세요.\n')
        print('3초 후 검색 메뉴로 돌아갑니다.')
        time.sleep(3)
        return Search(student)



    
# ----- 성적 수정 -----
def Update(student):
    id_regex = re.compile("[0-9]{10}")
    RenderList(student)
    editSelect = input('수정할 행의 학번을 입력해주세요. : ')
    validation = id_regex.search(editSelect.replace(" ",""))
    
    if validation:
        for key, value in student.items():
            if editSelect == value[0]:
                RenderList(student)
                print(" * 성적은 국어, 영어, 수학 순으로 입력 / 구분자 ','\n")

                kor, eng, math = map(int, input("성적 입력 : ").split(','))
                score = kor + eng + math
                avg = round(score/3, 3)

                qry="update student set Score_Kor=?,Score_Eng=?,Score_Math=?,Score_avg=? where Student_ID=?;"
                try:
                    cur=db.cursor()
                    cur.execute(qry, (kor, eng, math, avg, editSelect))
                    db.commit()
                except sqlite3.ProgrammingError as e:
                    print("error in operation\n")
                    print(e)
                    time.sleep(5)
                    db.rollback()
                    db.close()

                RenderList(student)
                print("1행이 수정되었습니다.\n")
                backToMain = input("아무키나 누르면 초기 메뉴로 돌아갑니다.")
                return

            else:
                continue
        else:
            print("해당 학번이 없습니다.\n")
            backToMain = input('아무키나 누르면 초기 메뉴로 돌아갑니다.')
            return 
    else:
            print('잘못된 학번 형식입니다.\n')
            print('3초 후 수정 메뉴로 돌아갑니다.')
            return Update(student)


# ----- 학생 삭제 -----
def Delete(student):
    id_regex = re.compile("[0-9]{10}")
    RenderList(student)
    deleteSelect = input('삭제할 행의 학번을 입력해주세요. : ')
    validation = id_regex.search(deleteSelect.replace(" ",""))

    if validation:
        for key, value in student.items():
            if deleteSelect == value[0]:
                qry="DELETE from student where Student_ID=?;"
                try:
                    cur=db.cursor()
                    cur.execute(qry, (deleteSelect,))
                    db.commit() 
                except sqlite3.ProgrammingError as e:
                    print("error in operation\n")
                    print(e)
                    time.sleep(5)
                    db.rollback()
                    db.close()
                RenderList(student)
                print("1행이 삭제되었습니다.\n")
                backToMain = input("아무키나 누르면 초기 메뉴로 돌아갑니다.")
                return

            else:
                continue

        print("해당 학번이 없습니다.\n")
        backToMain = input('아무키나 누르면 초기 메뉴로 돌아갑니다.')
        return 
    else:
            print('잘못된 학번 형식입니다.\n')
            print('아무키나 누르면 삭제 메뉴로 돌아갑니다.')
            return Delete(student)


def RenderList(student):
    os.system('clear')
    numberOfstudents = 0

    print("┌--------------------------------------┐")
    print("│                                      │")
    print("│       성적 관리 프로그램 v0.2        │")               
    print("│                                      │")
    print("└--------------------------------------┘\n")

    print('이름 / 학번 / 국어 / 영어 / 수학 / 평균')
    print("=====================================")
    for key, value in student.items():
        print(key,':', value)
        numberOfstudents += 1
    print('\n{} results found'.format(numberOfstudents))
    print("=====================================")

def Export(student):
    # 현 시각 정보 객체 생성
    now = time.localtime()

    # 새 엑셀 생성
    wb = openpyxl.Workbook()

    # 활성화된 시트 지정
    ws = wb.active

    # row 데이터 헤더 
    ws.append(["No.", "이름", "학번", "국어점수", "영어점수", "수학점수", "평균"])
    index = 1
    for key, value in student.items():
        ws.append([index,key,value[0],value[1],value[2],value[3],value[4]])
        index += 1
    wb.save("%04d%02d%02d_%02d%02d%02d.xlsx" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec))
    RenderList(student)
    print("\n저장 완료\n")
    backToMain = input("아무키나 누르면 초기화면으로 돌아갑니다.")
    return
    
def main():
    #student 딕셔너리 생성
    
    
    while True:   

        student = dict()

        sql="SELECT * from student;"
        cur=db.cursor()
        cur.execute(sql)    

        students=cur.fetchall()
        for rec in students:
            student[rec[0]] = [rec[1], rec[2], rec[3], rec[4], rec[5]]
            # person = student[rec[0]] 

        RenderList(student)
        
        select = int(input("1.입력 2.검색 3.수정 4.삭제 5.엑셀저장 6.종료 \n"))

        # ----- 성적 입력 -----
        if select == 1:
            student = Insert(student)

        # ----- 학생 검색 -----
        elif select == 2:
            student = Search(student)

        # ----- 학생 수정 -----
        elif select == 3:
            student = Update(student)

        # ----- 학생 삭제 -----
        elif select == 4:
            student = Delete(student)
            
        elif select == 5:
            Export(student)

        else:
            print("종료되었습니다.")
            # db.close()
            exit()

main()
