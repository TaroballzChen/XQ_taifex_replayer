import openpyxl

Replay_Data = "15/60.xlsx"
closePrice_col = "F"
SMA20_col = "H"

class Replay():
    def __init__(self,FilePath):
        self.wb = self.openFile(FilePath)
        self.wsheet = self.activeSheet()

    def openFile(self,Filepath):
        wb = openpyxl.load_workbook(Filepath)
        return wb

    def activeSheet(self):
        wsheet = self.wb.active
        return wsheet

    def getValue(self,col,index):
        value = self.wsheet["%s%d"%(col,index)].value
        return value

    def getNowPrice(self,index):
        price = self.getValue(closePrice_col,index)
        return int(price)

    def getSMA20(self,index):
        SMA20 = self.getValue(SMA20_col,index)
        return float(SMA20)

def my_strategy(index):
    buy_score = 0
    sell_score = 0

    # SMA20 strategy
    SMA_20 = play.getSMA20(index)
    if Last_price > SMA_20:
        buy_score += 1
    elif Last_price < SMA_20:
        sell_score += 1
    else:
        pass

    return buy_score, sell_score

if __name__ == '__main__':
    Status = ""
    index = 2
    play = Replay(Replay_Data)
    test_price = play.getValue("F",index)
    buy_price = 0
    sell_price = 0
    total = 0
    while (test_price != None):
        Last_price = play.getNowPrice(index)
        buy_score,sell_score = my_strategy(index)
        if Status == "buy":
            if sell_score >= 1:
                point = Last_price - buy_price
                print("獲利(下空單)：%d"%point)
                Status = ""
                total += point

        elif Status == "sell":
            if buy_score >= 1:
                point = sell_price - Last_price
                print("獲利(下多單)：%d"%point)
                Status = ""
                total += point


        if Status == "":
            if buy_score >=1:
                Status = "buy"
                buy_price = Last_price
            elif sell_score >=1:
                Status = "sell"
                sell_price = Last_price

        index += 1
        test_price = play.getValue("F", index)
    else:
        print("總獲利：",total)