#요구사항 바뀜

from openpyxl import load_workbook
#파일 리스트 가져오기
from os import listdir, getcwd
path_dir = getcwd() 
file_list = listdir(path_dir)
fnames = []
for i in file_list:
#    print(i)
    if i[-4:] == 'xlsx' and i[:2] != '~$':
        fnames.append(i)

#첫 줄 비교
def isEqui(t1, t2):
  for i in range(len(t2)):
    if t1[i] != str(t2[i].value):
      return False 
  return True

#연도 줄 비교
def isEquiY(t1, t2):
  for i in range(len(t2)):
    if t1[i] != str(t2[i].value):
      return False 
  return True

#비교 양식 하드코딩 #양식 리스트는 다른 프로그램으로 뽑음
'''양식입력'''
'''cform = []
cform.append(['2020-01-01 00:00:00', '2020-01-02 00:00:00', '2020-01-03 00:00:00', '2020-01-04 00:00:00', '2020-01-05 00:00:00', '2020-01-06 00:00:00', '2020-01-07 00:00:00', 'None'])
cform.append(['2020-02-01 00:00:00', '2020-02-02 00:00:00', '2020-02-03 00:00:00', '2020-02-04 00:00:00', '2020-02-05 00:00:00', '2020-02-06 00:00:00', '2020-02-07 00:00:00', 'None'])
cform.append(['2020-03-01 00:00:00', '2020-03-02 00:00:00', '2020-03-03 00:00:00', '2020-03-04 00:00:00', '2020-03-05 00:00:00', '2020-03-06 00:00:00', '2020-03-07 00:00:00', 'None'])
'''
cform = [
  ['기업체명', 'None', '삼성에스디아이(주) ', '영문기업명', 'SAMSUNG SDI CO., LTD.', 'None', 'None', 'None'],
  ['결산기준일자', '총자산', '납입자본금', '자본총계', '매출액', '영업이익', '순이익', 'None'],
  ['연혁일자 ▲▼', '내용', 'None', 'None', 'None', 'None', 'None', 'None'],
  ['내용', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['구분', '상시종업원수', 'None', 'None', 'None', 'None', '평균근속년수', '연급여총액'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['구분', '상시종업원수', 'None', 'None', 'None', 'None', '평균근속년수', '연급여총액'], ['구분', '상시종업원수', 'None', 'None', 'None', 'None', '평균근속년수', '연급여총액'], ['구분', '상시종업원수', 'None', 'None', 'None', 'None', '평균근속년수', '연급여총액'], ['구분', '상시종업원수', 'None', 'None', 'None', 'None', '평균근속년수', '연급여총액'], ['번호', '구분', '사업자번호', '사업장명', '주생산품', '주소', '전화번호', 'None'], ['구분', 'None', '본사', '사업장명', '본점', 'None', 'None', 'None'], ['구분', 'None', '국내지사', '사업장명', '부산공장', 'None', 'None', 'None'], ['구분', 'None', '국내지사', '사업장명', '천안지점', 'None', 'None', 'None'], ['설비명', '2013-12-31 00:00:00', 'None', 'None', '2014-12-31 00:00:00', 'None', 'None', '2015-12-31 00:00:00'], ['설비명', '2014-12-31 00:00:00', 'None', 'None', '2015-12-31 00:00:00', 'None', 'None', '2016-12-31 00:00:00'], ['설비명', '2015-12-31 00:00:00', 'None', 'None', '2016-12-31 00:00:00', 'None', 'None', '2017-12-31 00:00:00'], ['설비명', '2016-12-31 00:00:00', 'None', 'None', '2017-12-31 00:00:00', 'None', 'None', '2018-12-31 00:00:00'], ['설비명', '2017-12-31 00:00:00', 'None', 'None', '2018-12-31 00:00:00', 'None', 'None', '2019-12-31 00:00:00'], ['사업부문', '주요제품(품목)', '매출액구성비(%)', 'None', 'None', 'None', 'None', 'None'], ['사업부문', '주요제품(품목)', '매출액구성비(%)', 'None', 'None', 'None', 'None', 'None'], ['사업부문', '주요제품(품목)', '매출액구성비(%)', 'None', 'None', 'None', 'None', 'None'], ['사업부문', '주요제품(품목)', '매출액구성비(%)', 'None', 'None', 'None', 'None', 'None'], ['사업부문', '주요제품(품목)', '매출액구성비(%)', 'None', 'None', 'None', 'None', 'None'], ['사업부문', '주요제품(품목)', '매출액구성비(%)', 'None', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['기업명', '사업자번호', '대표자명', '거래비중', '결산년도', '자본금', '자산총계', '매출액'], ['기업명', '사업자번호', '대표자명', '거래비중', '결산년도', '자본금', '자산총계', '매출액'], ['기업명', '사업자번호', '대표자명', '거래비중', '결산년도', '자본금', '자산총계', '매출액'], ['기업명', '사업자번호', '대표자명', '거래비중', '결산년도', '자본금', '자산총계', '매출액'], ['기업명', '사업자번호', '대표자명', '거래비중', '결산년도', '자본금', '자산총계', '매출액'], ['기업명', '사업자번호', '대표자명', '거래비중', '결산년도', '자본금', '자산총계', '매출액'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2017-12-31 00:00:00', '2018-12-31 00:00:00', '2019-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['계정명', '2014-12-31 00:00:00', '2015-12-31 00:00:00', '2016-12-31 00:00:00', 'None', 'None', 'None', 'None'], ['등급', '평가(산출)일자', '재무기준일자', '등급구분', 'None', 'None', 'None', 'None']]
#from openpyxl import load_workbook

SHEETNO = 54
infoStr = ''
sinfoStr = '' #상세
rn = '\n'

#개선점: 나중에 읽어오도록 고치기, 
#개선점: 앞에 두 줄만 읽어오기
#개선점: tkinter로 표시
#개선점: 어디가 의심스러운지 표시

for fname in fnames:
  infoStr = infoStr + '파일 '+fname+': '+rn
  wb = load_workbook(fname)

  emptySheetNames = [] #빈 시트 이름 리스트
  milSheetNames = [] #밀린 시트 이름 리스트
  susSheetNames = [] #첫 줄/형식 의심스러운 시트 이름 리스트
  errStr = ''

  #시트 이름 리스트
  sheets = wb.sheetnames
  smax = len(sheets)

  #시트별 루프
  for n in range(smax):
    sname = sheets[n] #시트명
    ws = wb[sname]
    #개선점: 가장 긴 첫 줄 기준으로 길이 변경
    fline = ws['A1:H1'] #tuple
    sline = ws['A2:H2'] #second line(empty check)
    
    #시트 수 체크
    if smax != SHEETNO: errStr = errStr + "시트 수: "+smax+rn

    #빈 시트 체크
    #len으로 체크하도록 바꾸기

    if fline[0][0].value == None: 
      if fline[0][1].value != None or sline[0][0].value != None or sline[0][1].value != None: 
        #밀린 시트
        milSheetNames.append(sname)
      else:
        #빈 시트
        emptySheetNames.append(sname)
      continue
    else: #문제없음 
      pass

    #첫 페이지와 사업자번호 체크
    if n == 0: #첫 시트
      '''체크 필요 HERE'''
      if ws['A2'] != '사업자번호': susSheetNames.append(sname) 
      cname = str(ws['C2'].value) #위치 모름, 확인필요
      if smax > SHEETNO: cname = '★'+cname
      if fname[:-5] != cname: 
        errStr = errStr + '파일명이 (★)사업자번호와 다릅니다: 파일명-'+fname[:-5]+', 사업자번호-'+cname+rn

    #첫 줄 체크
    #일반적인 범위
    elif n < SHEETNO:
      if not isEqui(cform[n], fline[0]):
        #err msg
        susSheetNames.append(sname)
      #연도 체크 넣는다면 여기
    #넘어가는 범위
    else: 
      #시트이름 체크
      if sname[:len(str(n+1))] != str(n+1):
        errStr = errStr + '시트 위치가 올바르지 않습니다: 시트 순번-'+str(n+1)+', 시트 이름-'+sname +rn
      #시트 첫 줄 체크
      '''체크 필요 HERE'''
      if not isEqui(cform[2], fline[0]):
        #err msg
        susSheetNames.append(sname)
  wb.close()
  #output
  if len(susSheetNames) >0: 
    errStr = errStr + '첫 줄/사업자번호가 의심스러운 시트 리스트: '+str(susSheetNames)+rn
  if len(milSheetNames) >0: 
    errStr = errStr + '붙여넣기 위치가 잘못된 시트 리스트: '+str(milSheetNames)+rn
  if len(emptySheetNames) >0: 
    errStr = errStr + '빈 시트 리스트: '+str(emptySheetNames)+rn

  if errStr == '': infoStr+='OK'+rn
  else: infoStr= infoStr + errStr+rn

fo = open('report.txt', "w+") 
fo.write(infoStr)
fo.close()