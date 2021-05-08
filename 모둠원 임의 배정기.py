import random as rn
import openpyxl as xl
import os, time

def xl_w(cla, n, g):
    sheet["A"+str(n+1)]=cla
    sheet["B"+str(n+1)]=n
    sheet["C"+str(n+1)]=g
    if str(n) in bose:
        sheet["D"+str(n+1)]=1

def arrange(i,limit): 
    while 1: 
        tmp=rn.randint(1,g_num)
        if group[tmp]<limit:
            stu[i]=tmp 
            group[tmp]+=1
            break
print("\n##모둠원 임의 배정기(v2.0)##")
print("programed by 득쌤❤")
time.sleep(1)
print("\n( ﾉ ﾟｰﾟ)ﾉ ~~ 안녕하세요.")
print("\n본 프로그램은 모둠원들을 임의로 배정하여 엑셀파일로 출력해주는 기능을 수행합니다.")
print("\n먼저 모둠의 수와 모둠장이 결정되어 있어야 합니다.")
print("\n그럼 지금부터 모둠원들을 배정해보겠습니다. 뿅~~ (∩^o^)⊃━☆\n\n")

for i in range(8,-1,-1):
    print(i)
    time.sleep(1)
    
os.system("cls")
print("\n아래의 5가지 질문에 숫자로만 답해주세요.")
print("\n여러 명을 입력할 경우, 공백1개로 구분하여 입력하세요. (예시: 3 6 9 12)")
time.sleep(0.5)
cla=int(input("\n1) 몇 반인가요? "))
tot=int(input("\n2) 학급의 마지막 번호를 입력하세요. "))
g_num=int(input("\n3) 모둠의 수는? "))

while 1:
    bose=input("\n4) 모둠장의 번호를 모둠번호 순으로 입력하세요. (예: 3 6 19 21) ").split()
    if len(bose)!=g_num: print("\n(경고!!) 모둠장의 수와 모둠의 수가 다릅니다. 다시 입력하세요.")
    else: break
    
exc=input("\n5) 랜덤배정에서 제외시켜야하는 번호는?(골프부, 전학생 등 / 없으면 그냥 엔터~!) ").split()

g_s_num=(tot-len(exc))//g_num     # 모둠별 기본 학생수
g_s_num_add=(tot-len(exc))%g_num  # 모둠별 추가해야 할 학생수
 
stu=[0]*(tot+1)  # 학생별 모둠 번호
group=[0]*(g_num+1)  # 모둠별 배정 인원

# 모둠장 모둠 배정
b_i=1
for i in bose:
    i=int(i)
    stu[i]=b_i
    group[b_i]+=1
    b_i+=1

# 제외학생 반영
for i in exc: 
    i=int(i)
    stu[i]=-1

#(기본 학생)학생별 모둠 배정
i=1; cnt=g_num
while cnt<g_s_num*g_num: 
    if stu[i]==0:
        arrange(i,g_s_num)
        cnt+=1
    i+=1

#(추가 학생)학생별 모둠 배정
cnt=0
while cnt<g_s_num_add: 
    if stu[i]==0:
        arrange(i,g_s_num+1)
        cnt+=1
    i+=1

book=xl.Workbook()
sheet=book.active
sheet=book.create_sheet(str(cla)+"반",0)
sheet["A1"]="반"
sheet["B1"]="번호"
sheet["C1"]="모둠"
sheet["D1"]="모둠장"

print()
for i in range(1,tot+1):
    print(stu[i],end=' ')
    xl_w(cla,i,stu[i])

book.save("Group.xlsx")

print("\n\n#################")
print("\n\n배정이 끝났습니다. 엑셀파일을 확인하세요.")
print("\n엑셀파일은 본 프로그램과 동일한 폴더에 있습니다.")
print("\n즐거운 모둠활동 되시길~~ ヾ(⌐■_■)ノ♪\n\n")

os.system("Pause")
