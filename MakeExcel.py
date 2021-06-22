import pandas as pd
import openpyxl

def MakeNaverToDel(loadFilename, fileDirectory, quantityTypeVal, quantityEdit, fareDivisionEdit, saveFileName):
    df = pd.read_excel(loadFilename, skiprows=[0], usecols=[1, 8, 10, 16, 40, 41, 42, 45])
    selectcell = pd.DataFrame(index=range(0, len(df)), columns={
        '수취인명': [],
        '우편번호': [],
        '배송지': [],
        '수취인연락처1': [],
        '수취인연락처2': [],
        '수량'+'('+quantityTypeVal+'타입)': [],
        '운임': [],
        '운임구분': [],
        '상품명': [],
        '배송메세지': [],
        '주문번호': [],
        '주문자': []
    })

    print(df)

    selectcell['수취인명'] = df['수취인명']
    selectcell['배송지'] = df['배송지']
    selectcell['수취인연락처1'] = df['수취인연락처1']
    selectcell['수취인연락처2'] = df['수취인연락처2']
    selectcell['수량'+'('+quantityTypeVal+'타입)'] = quantityEdit
    selectcell['운임구분'] = fareDivisionEdit
    selectcell['상품명'] = df['상품명']
    selectcell['배송메세지'] = df['배송메세지']
    selectcell['주문번호'] = df['주문번호']
    selectcell['주문자'] = df['구매자명']

    selectcell.to_excel(fileDirectory+'/'+saveFileName+'.xlsx', index=False, )

def MakeCoupangToDel(loadFilename, fileDirectory, quantityTypeVal, quantityEdit, fareDivisionEdit, saveFileName):
    df = pd.read_excel(loadFilename, usecols=[2, 10, 24, 26, 27, 29, 30])
    selectcell = pd.DataFrame(index=range(0, len(df)), columns={
        '수취인명': [],
        '우편번호': [],
        '배송지': [],
        '수취인연락처1': [],
        '수취인연락처2': [],
        '수량' + '(' + quantityTypeVal + '타입)': [],
        '운임': [],
        '운임구분': [],
        '상품명': [],
        '배송메세지': [],
        '주문번호': [],
        '주문자': []
    })

    print(df)
    selectcell['수취인명'] = df['수취인이름']
    selectcell['배송지'] = df['수취인 주소']
    selectcell['수취인연락처1'] = df['수취인전화번호']
    selectcell['수량' + '(' + quantityTypeVal + '타입)'] = quantityEdit
    selectcell['운임구분'] = fareDivisionEdit
    selectcell['상품명'] = df['등록상품명']
    selectcell['배송메세지'] = df['배송메세지']
    selectcell['주문번호'] = df['주문번호']
    selectcell['주문자'] = df['구매자']

    selectcell.to_excel(fileDirectory + '/' + saveFileName + '.xlsx', index=False)

def CutDeliveryEx(loadFilename, fileDirectory, saveFileName):
    dfNaver = pd.read_excel(loadFilename[0], skiprows=[0],dtype = {'(수취인연락처1)': str, '(수취인연락처2)': str})
    dfCoupang = pd.read_excel(loadFilename[1])
    dfDell = openpyxl.load_workbook(loadFilename[2])
    wb = dfDell.active

    dlistnaver = []
    dlistcoupang = []
    for x in range(2, len(dfNaver)+2):
        dlistnaver.append(wb['G'+str(x)].value)

    for x in range(len(dfNaver)+2, wb.max_row+1):
        dlistcoupang.append(wb['G' + str(x)].value)

    dfnaverdel = pd.DataFrame({'송장번호': dlistnaver})
    dfcoupangdel = pd.DataFrame({'운송장번호': dlistcoupang})

    dfNaver['송장번호'] = dfnaverdel['송장번호']
    dfCoupang['운송장번호'] = dfcoupangdel['운송장번호']

    dfNaver.to_excel(fileDirectory+'/'+saveFileName+'(Naver)'+'.xlsx', sheet_name='발송처리', index=False)
    dfCoupang.to_excel(fileDirectory + '/' + saveFileName + '(Coupang)' + '.xlsx', sheet_name='발송처리', index=False)