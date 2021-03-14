import os , sys , getopt
import pandas as pd , numpy as np


def classTime (time):
    timeList = { '1' : '오후 09:00' ,
                 '2' : '오전 10:00' ,
                 '3': '오전 11:00',
                 '4' : '오후 12:00',
                 '5' : '오후 13:00',
                 '6' : '오후 14:00',
                 '7' : '오후 15:00',
                 '8' : '오후 16:00',
                 '9' : '오후 17:30',
                 '10': '오후 18:25',
                 'A' : '오전 09:30',
                 'B' : '오전 11:00',
                 'C' : '오후 13:00',
                 'D' : '오후 14:30',
                 'E' : '오후 16:00'}
    clsTime = list(timeList.keys())
    return timeList[clsTime[clsTime.index(time)]] , timeList[clsTime[clsTime.index(time)+1]]

def fileSplit(file ,time) :
    first , second  = classTime(time)
    firstHour = int(first[3:5])
    secondHour = int(second[3:5])
    checkTime = []
    checkPoint = []
    point = None
    point2 = None

    for i in file :
        if len(i)<40:
            pass
        elif i.find('2021년') != -1 :
            if i.find('오전') != -1 :
                index = i.find('오전')
            else :
                index = i.find('오후')
            if i.find(':')+2 - index <7 :
                chatTime = i[index:index + 7]
                chatTime = chatTime.replace('\t', '')
                chatHour = int(chatTime[3:4])
            else :
                chatTime = i[index:index+8]
                chatTime = chatTime.replace('\t', '')
                chatHour = int(chatTime[3:5])
            if chatHour <= firstHour :
                """checkTime.append(i[20:28])
                checkPoint.append(file.index(i)+20)"""
                point = file.index(i)
            elif chatHour == secondHour :
                point2 = file.index(i)

    checkPoint.append(point)
    checkPoint.append(point2)
    checkTime.append(first)
    second = second.replace('00','50')
    checkTime.append(second)
    print(checkPoint)
    print(checkTime)
    return checkTime , checkPoint
def acceptRange(hour,minute):
    if minute - 5  < 0 :
        min_minute = 60-5+minute
        min_hour = hour -1
    else :
        min_minute = minute - 5
        min_hour = hour
    if minute + 5 > 60 :
        max_minute = minute - 60 + 5
        max_hour = hour+1
    else :
        max_minute = minute + 5
        max_hour = hour
    return min_hour,max_hour,min_minute , max_minute


def findChat(file , checkTime):

    name = []
    chat = []
    time = []
    studentList = []
    result = []
    acceptMinute = int(checkTime[6:8])
    acceptHour = int(checkTime[3:5])
    minHour,maxHour,minRange, maxRange = acceptRange(acceptHour, acceptMinute)
    #find chat time , chatter , chat contents
    for i in file :
        if i == '\n':
            pass
        elif i.find('2021년') != -1 :
            if i.find('오전') != -1:
                index = i.find('오전')
            else:
                index = i.find('오후')
            if i.find(':')+2 - index <7 :
                chatTime = i[index:index + 7]
                chatTime = chatTime.replace('\t', '')
                chatHour = int(chatTime[3:4])
                chatMinute = int(chatTime[5:7])
            else :
                chatTime = i[index:index+8]
                chatTime = chatTime.replace('\t', '')
                chatHour = int(chatTime[3:5])
                chatMinute = int(chatTime[6:8])


            if chatMinute-5<=0 :
                chatMinute = chatMinute+60

            if minHour<=chatHour<=maxHour and minRange<= chatMinute <= maxRange+60 and'201' in i[56:] :
                name.append(i[i.index('201'):i.index('201')+13])
                chat.append(i[i.rfind('201'):].replace('\n',''))
                time.append(chatTime)
    #find student name
    for i in chat :
        loc = i.find('201')
        studentList.append(i[loc:loc+13])
    #extract student name and time
    for i in range(0,len(studentList)) :
        temp = []
        temp.append(time[i])
        temp.append(name[i])
        temp.append(studentList[i])
        result.append(temp)
    return result
def attendenceCheck(first , second ):
    attendance = []
    absent = []
    for i in first:
        ok = False
        temp = []
        for j in second:
            if i[1] == j[1]:
                temp.append(i[1])
                temp.append('O')
                temp.append('O')
                attendance.append(temp)
                ok = True
        if ok is False:
            temp.append(i[1])
            temp.append('O')
            temp.append('X')
            attendance.append(temp)
            absent.append(temp)
    return attendance , absent

def exportAsExecel(date,attd) :
    for i in attd :
        if len(i)>3 :
            del i[3:]
    atd = pd.DataFrame(attd,columns=['Student','first class','Second class'])
    print(atd.head())
    atd.to_excel(date+'-Attendence.xlsx' , sheet_name = 'Sheet1')



def controller(source , classTime) :
    f = open(source, 'r', encoding='utf-16')
    date = source[source.index('2021-'):source.index('2021-')+10]
    file = f.readlines()
    checkTime, checkPoint = fileSplit(file,classTime)

    # find first attendance chat
    first_chat = findChat(file[:checkPoint[0]], checkTime[0])
    first_chat = sorted(first_chat, key=lambda student: student[1])

    second_chat = findChat(file[checkPoint[0] + 1:checkPoint[1]], checkTime[1])
    second_chat = sorted(second_chat, key=lambda student: student[1])

    attendense, absent = attendenceCheck(first_chat, second_chat)
    print('attend : ' + str(len(attendense) - len(absent)))
    print('absent : ' + str(len(absent)))
    print(absent)

    exportAsExecel(date,attendense)
    f.close()


def main(argv):
    SourcePath = None
    classStartTime = None
    try :
        opts , etc_args = getopt.getopt(argv[1:],"h:i:c:",['help','InputFile = ','ClassTime ='])
    except:
        print(argv[0],'-i <InputFile> -c <CLassTime>')
        sys.exit(2)
    for opt, arg in opts :
        if opt in("-h", '--help'):
            print(argv[0], '-i <InputFile> -c <CLassTime>')
        elif opt in ('-i','--input'):
            SourcePath = arg
        elif opt in ('-c','--classtime'):
            classStartTime = arg
    if len(SourcePath)<1 :
        print(argv[0], "-i option is mandatory")
        sys.exit(2)
    elif len(classStartTime)<1 :
        print(argv[0], "-c option is mandatory")
        sys.exit(2)
    else :
        controller(SourcePath , classStartTime)

if __name__ == "__main__":
    main(sys.argv)
