#xls용
#입출력: 31.용인시.csv, 31.용인시_브리핑.csv -> 31.용인시_비교결과.txt, 31.용인시_비교결과_빈파일.txt
#예시 입력 기업 리스트 파일명: 1.성남시.csv, 30.동두천.csv
#예시 입력 브리핑 리스트 파일명: 성남시_브리핑.csv, 동두천_브리핑.csv

CITY = input("도시 입력 숫자+.+3글자(예 - '1.성남시'): ")

text = ''
cnt, errcnt, emptycnt = 0, 0, 0
cnum, cname, cbn, bnum, bname, bbn, btp = '','','','','','', ''

def splitL(line):
	num, tmp = line[:-1].split(',', 1) #마지막 엔터 제외
	if tmp[0] == '"':
		if tmp.rfind('"') != -1:
			x = tmp.rfind('"')
			name = tmp[1:x]
			if tmp[x+1] == ',': bntp = tmp[x+2:]
			else: bntp = tmp[x+1:]
		else: name, bntp = tmp.split(',', 1)
	else: name, bntp = tmp.split(',', 1)
	if name[-1] == ' ': name=name[:-1] #리스트 맨 뒤의 공백문자 하나 처리
	if bntp != '':
		#용인시 사업자번호 중 '-' 없는 경우 때문
		if bntp[0] in '1234567890':
			if bntp[3] != '-' and bntp[6] != '-':
				bntp = bntp[:3]+'-'+bntp[3:5]+'-'+bntp[5:]
		if bntp[-1]==' ': bntp = bntp[:-1]
	name = name.replace('  ', ' ')
	return num, name, bntp

try:
	fic = open(CITY+'.csv', "r", encoding='UTF8') 
	fib = open(CITY+'_브리핑.csv', "r", encoding='UTF8') 
	fo = open(CITY+"_비교결과.txt", "w+", encoding='UTF8')
	fo2 = open(CITY+"_비교결과_빈파일.txt", "w+", encoding='UTF8')
	fo.write("번호 업체명 사업자번호 확장자\n")
	fo2.write("번호,업체명,파일 사업자번호,확장자,기타에러,목록 사업자번호\n")
	#첫 라인
	fic.readline()
	fib.readline()
	cline = fic.readline()
	bline = fib.readline()
	
	while(True):
		# 맨 뒤의 \n 외에 ,로 나누기
		# 업체명 안에 ',' 포함인 경우 ""로 감싸여있음

		cnum, cname, cbn = splitL(cline)
		bnum, bname, btp = splitL(bline)
		bbn, btp = btp.split(',', 1)

		#print("cbn: '"+str(cbn)+"', bbn: '"+str(bbn)+"', btp: '"+str(btp)+"'\n")

		#cnum 앞에 '0' 패딩 #짧으니까 하드코딩이 나을 듯
		if len(cnum)<4: # 4나 5인 경우 있음
			if len(cnum) == 1:
				cnum = '000'+cnum
			elif len(cnum) == 2:
				cnum = '00'+cnum
			elif len(cnum) == 3:
				cnum = '0'+cnum
		
		#리스트 업체명 앞뒤 '(주)', '㈜', 그리고 그 앞뒤의 띄어쓰기 처리
		if cname[:3] == '(주)': cname = cname[3:]
		elif cname[0] == '㈜': cname = cname[1:]
		if cname[-3:] == '(주)': cname = cname[:-3]
		elif cname[-1] == '㈜': cname = cname[:-1]
		if cname[0] == ' ': cname = cname[1:]
		if cname[-1] == ' ': cname = cname[:-1]

		#비교
		text = ''
		text2=''
		if cnum != bnum: 
			text+= '\n - 목록 숫자: '+cnum
			text2+='숫자'
		if cname != bname: 
			text+= '\n - 목록 업체명: '+cname
			if text2!='': text2+='와 '
			text2+='업체명'
		if cbn != bbn: 
			text+= '\n - 목록 사업자번호: '+cbn
		if text != '': 
			if bbn[0] not in '1234567890':
				if text2=='': text2='없음'
				text = bnum+','+bname+','+bbn+','+btp + ','+text2+','+cbn + '\n'
				fo2.write(text)
				emptycnt+=1
			else: 
				text = bnum+' '+bname+' '+bbn+' '+btp + text + '\n'
				fo.write(text)
				errcnt +=1
		cnt+=1
		cline = fic.readline()
		bline = fib.readline()
		#이스케이프
		if not cline:
			if not bline: break
			else: 
				fo.write("목록 수("+str(cnt)+") < 브리핑 파일 수\n")
				break
		elif not bline:
			fo.write("목록 수 > 브리핑 파일 수("+str(cnt)+")\n")
			break

	fo.write("검사 수: "+str(cnt)+", 문제 파일 수: "+str(errcnt)+", 판독불가 파일 수: "+str(emptycnt)+"\n")
	fo2.write("검사 수: "+str(cnt)+", 문제 파일 수:"+str(errcnt)+", 판독불가 파일 수: "+str(emptycnt)+'\n')
except:
	print("err, 검사 수: "+str(cnt)+", 문제 파일 수:"+str(errcnt)+", 판독불가 파일 수: "+str(emptycnt)+'\n')
	fo.write("err, 검사 수: "+str(cnt)+", 문제 파일 수:"+str(errcnt)+", 판독불가 파일 수: "+str(emptycnt)+'\n')
	fo2.write("err, 검사 수: "+str(cnt)+", 문제 파일 수:"+str(errcnt)+", 판독불가 파일 수: "+str(emptycnt)+'\n')
finally:
	fic.close()
	fib.close()
	fo.close()
	fo2.close()
