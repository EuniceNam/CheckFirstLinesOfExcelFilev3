#xls용, 파일명, 기업명, 사업자번호, 기업형태, 휴폐업, 주소 비교
#목록과 비교하지 않음
#파일 리스트 가져오기
from os import listdir, getcwd
path_dir = getcwd() 
file_list = listdir(path_dir)
fnames = []
emptyfiles = []
otherxlsfiles = []
for i in file_list:
    if (i[-4:] == '.xls') and i[:2] != '~$':
      if len(i.split('_')[0])>5:
        fnames.append(i)
      else:
        otherxlsfiles.append(i)
    elif i[-5:] == '.xlsx' and i[:2] != '~$':
      emptyfiles.append(i)
 
fo = open('report.txt', "w+") 
#CITY = "성남시"
print(".xls 수: "+str(len(fnames)))
CITY = input("도시 입력(예 - 성남시): ")
fo.write("도시: "+CITY+'\n')

def isEmpty(string):
  for i in string:
    if i != '\t' and i != '\n' and i != ' ': return False
  return True       
def parseAddr(string):
  a=[]
  a=string[:-1].split('\n')
  b=''
  if len(a) == 1: 
    return delTab(a[0])
  elif len(a) > 1: 
    b = delTab(a[1])
    a = b.split(' ')
    return a[0]+' '+a[1]
  else: 
    return delTab(string)
def delTab(string):
  ret=''
  for i in string:
    if i != '\t': ret+=i
  return ret

errfilecnt=0
othercnt=0
errbcnt=[0,0,0,0]
filecnt=0
errfileflag = False

cnumber =''

for fname in fnames:
  #print("file: "+fname)
  errfileflag = False
  cnumber = ''
  cnumberF=True
  try:
    fi = open(fname, 'r', encoding='UTF8')
  except: 
    fo.write(fname+' 을 열 수 없음\n')
    othercnt+=1
  firstIn=[True, True, True, True]
  try:
    line = fi.readline()
    trcnt = 0
    while line:
      if '<tr>' in line:
        trcnt+=1
        #print("trcnt: "+str(trcnt))
      if trcnt == 3 and firstIn[0]: #3째줄 기업명
        firstIn[0] = False
        line = fi.readline() #<th>기업명</th>
        line = fi.readline()
        cname = line.split('>')[1].split('<')[0]
        cname = cname.replace('  ', ' ')
        fcname = fname.split('_')[0]
        if fname[4]!='.' and (fname.find('.')+1<=len(fcname)): fcname = fcname[fname.find('.')+1:]
        else: fcname = fcname[5:]
        if fcname != cname and fcname != cname[3:] and fcname != cname[:-3]:
          fo.write(fname+"\t- 기업명 '"+cname+"'\n")
          errfileflag = True
          errbcnt[0]+=1
        if fcname[:3] == '(주)' or fcname[-3:] == '(주)':
          fo.write(fname+"\t- 파일명: (주) 포함\n")
          errfileflag = True
          errbcnt[0]+=1
      elif trcnt == 4 and cnumberF: #4째줄 사업자번호 tr
        line = fi.readline()
        line = fi.readline()
        cnumber = line.split('>')[1].split('<')[0]
        #print("should be 사업자번호 숫자: "+cnumber+'\n') #사업자번호 숫자
        cnumberF = False
      elif trcnt == 5 and firstIn[1]: #5째줄 기업형태
        firstIn[1] = False
        while line:
          line = fi.readline()
          if '기업형태' in line: break
        line = fi.readline()
        if '개인사업자' in line: 
          fo.write(fname+"\t- 기업형태: 개인사업자\n 사업자번호: "+cnumber+'\n')
          errfileflag = True
          errbcnt[1]+=1
      elif trcnt == 7 and firstIn[2]: #7째줄 주소
        #print(7)
        firstIn[2] = False
        line = fi.readline() #<th ...주소(도로명)</th>
        line = fi.readline() #<td ..>
        addr=''
        line = fi.readline() #(style)
        while line:
          if '</td>' in line: break
          if not isEmpty(line):  
            addr += line
          line = fi.readline()
        if CITY not in addr:
          addr = parseAddr(addr)
          fo.write(fname+"\t- 주소: "+addr+'\n 사업자번호: '+cnumber+'\n')
          errfileflag = True
          errbcnt[2]+=1
      elif trcnt>15 and firstIn[3]:
        #print(15)
        firstIn[3] = False
        while line:
          if '주요 신용정보' in line: break
          line = fi.readline()
          if '자산총계' in line: break #test
        tmpcnt=0
        while line:
          if '<tr>' in line:
            tmpcnt+=1
            if tmpcnt == 2: break
          line = fi.readline()
        corpstate = ''
        while line:
          if '</td>' in line: break
          if not isEmpty(line):
            corpstate += line
          line = fi.readline()
        if '휴업' in corpstate:
          fo.write(fname+"\t- 휴폐업: 휴업자\n"+' 사업자번호: '+cnumber+'\n')
          errfileflag = True
          errbcnt[3]+=1
        elif '폐업' in corpstate: 
          fo.write(fname+"\t- 휴폐업: 폐업자\n"+' 사업자번호: '+cnumber+'\n')
          errfileflag = True  
          errbcnt[3]+=1
      
      line = fi.readline()

    if errfileflag:
      errfilecnt+=1
    filecnt+=1
  except:
    #print('err')
    fo.write(fname+" 파일이 너무 짧습니다.\n")
    errfilecnt+=1
    filecnt+=1
  fi.close()
fo.write("\n검사완료 .xls 파일 수: "+str(filecnt)+'\n')
if len(otherxlsfiles)>0:
  fo.write("파일명 형식이 다른 .xls 파일("+str(len(otherxlsfiles))+"개): "+str(otherxlsfiles)+'\n')
if len(emptyfiles)>0: 
  fo.write("비어있어야하는 파일("+str(len(emptyfiles))+"개): "+str(emptyfiles)+'\n')
if othercnt>0:
  fo.write("열 수 없는 파일 "+str(othercnt)+'개\n')

if errbcnt[0]>0: fo.write("파일명 오류: "+str(errbcnt[0])+"개\n")
if errbcnt[1]>0: fo.write("개인사업자 : "+str(errbcnt[1])+"개\n")
if errbcnt[2]>0: fo.write("주소 오류 : "+str(errbcnt[2])+"개\n")
if errbcnt[3]>0: fo.write("휴폐업 오류: "+str(errbcnt[3])+"개\n")

if errfilecnt == 0:
  fo.write("\n모든 파일 OK")  
else: 
  fo.write("\n의심 파일 수: "+str(errfilecnt))
fo.close()