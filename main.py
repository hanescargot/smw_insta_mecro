# https://colab.research.google.com/drive/1J6JZji_z-5vpMQplmueado4s1iFmrS1t?usp=sharing
import pandas as pd
#  todo :  pip install openpyxl

def mainRun():

    #  3번째 열부터
    dataframe = pd.read_excel("/content/drive/MyDrive/SmallWave/data_table.xlsx", "그 외 아티스트", 3)



    #  데이터 컬럼만 뽑아내기 (복잡)
    userDataFrame = dataframe.iloc[::2, [2, 9]]  # 열 짝수로만 추출  (행선택),(열 선택)
    print(userDataFrame)

    # 그룹 수 구하기
    wholeDataCount = int(userDataFrame.count())
    print("값을 가진 컬럼 수 : ", wholeDataCount)

    # 각 파일에 하나씩 들어갈 문장 List 반환
    emailList = getHdrEmailList(userDataFrame, wholeDataCount)


    # 파일로 저장 끝
    for i in range(len(emailList)):
        exportStringFile(i, emailList[i])
    print("작업 끝")



def isNaN(string):
    return string != string

def getHdrEmailList(df, wholeDataCount):
    emailStrArr = [] # join(",") 할 100개
    groupStrArr = [] # 각 파일에 하나씩 들어갈 문장 목록
    count = 0
    dataCount = 0
    nullDataCount = 0
    while count - nullDataCount != wholeDataCount:
        isCoEmail = False
        if (str(df.iloc[count, 1]) == "o"): isCoEmail = True
        if (str(df.iloc[count, 1]) == "O"): isCoEmail = True
        if (str(df.iloc[count, 1]) == "ㅇ"): isCoEmail = True
        if (not isNaN(df.iloc[count, 0])):
            if (isCoEmail):
                # 소속사 있는것은 제외
                count += 1
                continue
            # row, column
            emileStr = getEmailFromString(str(df.iloc[count, 0]))

            if emileStr != "":
                dataCount += 1
                emailStrArr.append(emileStr)

            if dataCount != 0 and dataCount % 100 == 0:  # 100개 묶음 완성
                groupStrItem = (", ".join(emailStrArr))
                groupStrArr.append(groupStrItem)
                # 여기서 바로 프린트 해도 결곽 값은 잘 나옴
                emailStrArr = []  # 초기화
                dataCount = 0
        else:
            nullDataCount += 1
        count += 1


    # while 끝
    if dataCount != 0:
        # 100으로 나눠지지 않은 나머지 값
        groupStrArr.append(emailStrArr)
    print("얻어낸 이메일 수 : ", (len(groupStrArr)-1)*100 + len(groupStrArr[-1]))
    print("생성 가능한 파일 수 : ", len(groupStrArr))
    return groupStrArr

def getEmailFromString(dataStr):
    # null 이면 ""반환 : 이미 한번 해서 없어도 됨
    if(not dataStr):
        # print("isempty")
        return ""

    if("인스타" not in dataStr):
        return ""

    #공백, 띄어쓰기
    retrunStr = ""
    retrunStr = dataStr.replace(" ", "")
    retrunStr = retrunStr.replace("\n", "")
    retrunStr = retrunStr.replace("/", "")
    retrunStr = retrunStr.replace("\\", "") # \오타 존재함
    retrunStr = retrunStr.replace(":", "")
    retrunStr = retrunStr.replace("(소속사)", "") # 이메일 : (소속사) 이런식으로 되어 있음
    retrunStr = retrunStr.replace("(개인)", "")

    if("이메일" in dataStr ):
        retrunStr = retrunStr[retrunStr.index("인스타") + 3:retrunStr.index("이메일")]
    else:
        retrunStr = retrunStr[retrunStr.index("인스타")+3:]

    # print("data: ", dataStr)
    # print("return: ",retrunStr)
    return retrunStr

def exportStringFile(num, text):
    #      todo : 이메일들을 100개씩 콤마로 구분한 메모파일 생성 또는 프린트
    file = open(f"/content/drive/MyDrive/SmallWave/{num}.txt", "w")
    # print(len(text))
    file.write(str(text))
    file.close()
    return True

if __name__ == '__main__':
    mainRun()