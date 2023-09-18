import pandas as pd

names=[]
timday=[[],[],[],[],[]]
def readtable(name,day,week):
    table= pd.read_excel("课表/"+name+".xls")
    values=table.values[5,day]
#    print(values)
    for i in range(1,len(values)-2):
        if (values[i]+values[i+1]+values[i+2]=="[周]"):
            j=i-1
            fbegin=0
            fend=0
            if(values[j].isdigit()):
                    fend+=int(values[j])
                    j=j-1
                    if(values[j].isdigit()):
                        fend+=int(values[j])*10
                        j=j-1
                        if(values[j]=='-'):
                            j=j-1
                            if(values[j].isdigit()):
                                fbegin+=int(values[j])
                                j=j-1
                                if(values[j].isdigit()):
                                    fbegin+=int(values[j])*10     
                    elif(values[j]=='-'):
                        j=j-1
                        fbegin+=int(values[j])
            if(fbegin !=0):
#                print(fbegin,fend)
                if(week>=fbegin and week<=fend):
#                    print("没空")
                    return 1
        if(values[i]+values[i+1]+values[i+2]=="[单周" or values[i]+values[i+1]+values[i+2]=="[双周" or values[i]+values[i+1]+values[i+2]=="[周]"):
            j=i-1
#            print("e ")
            timetable=[]
            while(j):
                if(values[j].isdigit()):
                    if(values[j-1].isdigit()):
                        timetable.append(int(values[j-1]+values[j]))
                        j-=3;
                    elif(values[j-1]=="," or values[j-1]=="\n"):
                        timetable.append(int(values[j]))
                        j-=2;
                    else :
                        break
                else :
                    break
#            print(timetable)
            if (week in timetable):
#                print("没空")
                return 1
    return 0
            
                        

def main():
    allday=[]
    fin=open("名单.txt",mode="r",encoding='UTF-8')
    for line in fin.readlines():
        names.append(line.strip())
    print("益科技术组排班工具\n本工具可以找出值班时间内没有课有同学")
    week=input("请输入周次：")
    iname=0
    for name in names:
        tiday=["","","","","","","","","","","","",]
        tiday[0]=name
        for day in range(1,6):
            if( readtable(name,day,int(week))):
                tiday[day*2-1]="忙"
        for day in range(1,6):
            if( not  readtable(name,day,int(week))):
                timday[day-1].append(name)
        allday.append(tiday)
        iname+=1
    file=open("导出文件/"+"第"+str(week)+"周空闲时间表.txt",mode="w")
    file.writelines("以下同学下午4-6点没有课\n")
    for i in range(0,5):
        file.writelines('周'+str(i+1)+"："+str(timday[i])+"\n")
#         print('周',i+1,"：",str(timday[i]))
#    print(allday)
    file.close()
    table2=  pd.DataFrame(allday)
    table2.to_excel("导出文件/"+"第"+str(week)+"周空闲时间表.xlsx",header=None,index=False)
    print("数据己生成至以下文件：\n","第"+str(week)+"周空闲时间表.txt","第"+str(week)+"周空闲时间表.xlsx")
    input("按回车结束............")
            



if __name__ == '__main__':
    main()
 