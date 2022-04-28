from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import tqdm
import gspread
import datetime

scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("hah-project-8b41c84e8c4a.json", scope)

spreadsheet_name = "부자재 입출고 재고 시트"
client = gspread.authorize(creds)
#spreadsheet = client.open(spreadsheet_name)

# 접속 가능한 시트 파일 조회  ->  시트별 공유 설정된 파일 조회 가능함
#print(client.list_spreadsheet_files())

sheet_file = client.open('부자재 입출고 재고 시트')
worksheet_list = sheet_file.worksheets()

#현재 시간
#currentTime = datetime.datetime.now()
#DATE = str(currentTime.date()).split('-')
#TIME = str(currentTime.time().replace(microsecond=0)).split(':')
#print(DATE)
#print(TIME)
def x(a,b):
    return a - b

df_worksheet_list = []
for index in range(0, len(worksheet_list)-1):
    if index == 0 :
        worksheet = sheet_file.get_worksheet(index)   # 시트 인덱스 번호로 지정할때
        #print(worksheet)
        column_data = worksheet.col_values(3)
        maxRow = 4 + len(column_data[3:-1])
        df = pd.DataFrame(worksheet.get_all_values())
        df_worksheet_list.append(df)

        requestOut_list = df[3:][7].values.tolist()
        realOut_list = df[3:][5].values.tolist()
        lackList = []
        colorIndexList = []
        #print(realOut_list)
        for i in range(len(realOut_list)) :
            if requestOut_list[i] == "" :
                requestOut_list[i] = 0
            if realOut_list[i] == "" :
                realOut_list[i] = 0
            diff = int(requestOut_list[i]) - int(realOut_list[i])
            lackList.append([diff])
            if diff != 0 :
                colorIndexList.append(i + 4)
        #print(lackList)
        worksheet.update('G4:G'+ str(maxRow), lackList)
        
        for index in colorIndexList :   
            worksheet.format("G" + str(index), {
                "backgroundColor": {
                  "red": 1.0,
                  "green": 1.0,
                  "blue": 0.0
                },
                "textFormat": {
                  "bold": True
                }
            })
    elif index == 1 :
        continue
    elif index == 2 :
        continue
    elif index == 3 :
        continue
    elif index == 4 :
        continue
    #df.to_excel("./" + str(worksheet).split('\'')[1] + "(" + DATE[1] + DATE[2] + "_" + TIME[0] + "시" + TIME[1]+"분).xlsx", index = False)
    #df.to_excel("./" + str(worksheet).split('\'')[1]+".xlsx", index = False)