import win32com.client # 필요한 모듈을 가져온다
 
# 크레온 플러스 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# 현재가 객체 구하기
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')   # A005930이 삼성전자의 종목코드
objStockMst.BlockRequest() # 삼성전자의 현재가 객체를 가져온다
 
# 현재가 통신 및 통신 에러 처리 
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()
 
# Blockrequest로 가져온 현재가를 각각 이름에 부여
offer = objStockMst.GetHeaderValue(16)  #매도호가 <-주식을 매도할 수 있는 시장가격, 다른 코드는 삭제해도 된다

from slacker import Slacker

# After creating the bot in Slack, add OAuth Token below
slack = Slacker('<your-slack-api-token-goes-here>')

# Send a message to #general channel
# offer는 숫자열이기 때문에 str로 감싸서 문자열로 만들어준다
slack.chat.post_message('#stock', '삼성전자 현재가:' + str(offer))