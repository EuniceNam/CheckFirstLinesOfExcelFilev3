FNAME = input(".xlsx 파일명 입력: ")
FNAME = FNAME+".xlsx"
import pandas as pd
sheetnames = pd.read_excel(FNAME, None).keys()
for i in sheetnames:
	fi=pd.read_excel(FNAME, i).to_csv(i+'.csv', index=False)
