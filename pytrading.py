from kiteconnect import KiteConnect
import datetime
import xlsxwriter

if __name__ == "__main__":
    api_key = open("api_key.txt", 'r').read()
    api_secret = open("api_secret_key.txt", 'r').read()

    kite = KiteConnect(api_key=api_key)

#while running the program for first time otherwise hash
    access_token = open("access_token.txt", 'r').read()
    kite.set_access_token(access_token)

#while running the program after first time otherwise hash
    # print(kite.login_url())
    # data = kite.generate_session("", api_secret=api_secret)
    # print(data['access_token'])
    # kite.set_access_token(data['access_token'])
    # with open("access_token.txt", 'w') as ak:
    #     ak.write(data['access_token'])

instruments = ["NSE:PNB", "NSE:HATHWAY",
               "NSE:NHPC", "NSE:YESBANK", "NSE:SJVN", "NSE:SUZLON", "NSE:RPOWER","NSE:HUDCO","NSE:BCG", "BSE:PNB", "BSE:HATHWAY",
               "BSE:NHPC", "BSE:YESBANK", "BSE:SJVN", "BSE:SUZLON", "BSE:RPOWER","BSE:HUDCO","NSE:BCG"]
st_price = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
avg_price = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
price = []
get = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
order = {
    "buy_order_id": ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    "sell_order_id": ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    "ebuy_order_id": ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    "esell_order_id": ["0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"]}
volume = []
m = 1
k = 1
x = 1
# y = 1

while True:

    while datetime.time(9, 15, 0, 0) < datetime.datetime.now().time() < datetime.time(9, 15, 5, 0):
        s = kite.ltp(instruments)
        for i in range(0, len(instruments)):
            st_price[i] = round(((s[instruments[i]]["last_price"]) + st_price[i] * (k - 1)) / k * 20) / 20
        k = k + 1

    while datetime.time(9, 15, 5, 0) <= datetime.datetime.now().time() < datetime.time(9, 15, 6, 0):
        if len(volume) != len(instruments):
            s = kite.quote(instruments)
            for i in range(0, len(instruments)):
                volume.append(int((s[instruments[i]]["volume"]) //4000) + 1)

    while datetime.time(9, 15, 6, 0) <= datetime.datetime.now().time() < datetime.time(9, 15, 11, 0):
        while x != k:
            r = kite.ltp(instruments)
            for i in range(0, len(instruments)):
                avg_price[i] = round(((r[instruments[i]]["last_price"]) + avg_price[i] * (x - 1)) / x * 20) / 20
            x = x + 1

        if len(price) != len(instruments):
            s = kite.ltp(instruments)
            for i in range(0, len(instruments)):
                price.append(s[instruments[i]]["last_price"])
                if avg_price[i] > st_price[i]:
                    try:
                        order["buy_order_id"][i] = kite.place_order(variety="regular",
                                                                    exchange=instruments[i][:
                                                                                            instruments[i].index(":")],
                                                                    order_type="LIMIT",
                                                                    tradingsymbol=instruments[i][
                                                                                  instruments[i].index(":") + 1:],
                                                                    transaction_type="BUY",
                                                                    quantity=1,
                                                                    price=price[i],
                                                                    validity="DAY",
                                                                    product="MIS", )
                    except Exception as e:
                        order["buy_order_id"][i] = "0"
                        print(e)
                        pass

                if avg_price[i] < st_price[i]:
                    try:
                        order["sell_order_id"][i] = kite.place_order(variety="regular",
                                                                     exchange=instruments[i][:
                                                                                             instruments[i].index(":")],
                                                                     order_type="LIMIT",
                                                                     tradingsymbol=instruments[i][
                                                                                   instruments[i].index(":") + 1:],
                                                                     transaction_type="SELL",
                                                                     quantity=1,
                                                                     price=price[i],
                                                                     validity="DAY",
                                                                     product="MIS", )
                    except Exception as e:
                        order["sell_order_id"][i] = "0"
                        print(e)
                        pass

    while datetime.time(9, 15, 11, 0) <= datetime.datetime.now().time() < datetime.time(9, 32, 0, 0):
        for i in range(0, len(instruments)):
            get[i] = price[i] * 0.0025
            if get[i] % 0.05 != 0:
                get[i] = ((get[i] // 0.05) + 1) * 5 / 100

            if order["buy_order_id"][i] != "0":
                s = kite.order_history(order["buy_order_id"][i])

                if s[-1]["status"] == "OPEN":
                    try:
                        cbuy_order_id = kite.cancel_order(variety="regular",
                                                          order_id=order["buy_order_id"][i],
                                                          )
                    except Exception as e:
                        print(e)
                        pass
                    # volume[i] = s[-1]["filled_quantity"]
                elif s[-1]["status"] == "COMPLETE" and order["esell_order_id"][i] == "0":
                    order["esell_order_id"][i] = kite.place_order(variety="regular",
                                                                  exchange=instruments[i][:
                                                                                          instruments[i].index(":")],
                                                                  order_type="LIMIT",
                                                                  tradingsymbol=instruments[i][
                                                                                instruments[i].index(":") + 1:],
                                                                  transaction_type="SELL",
                                                                  quantity=1,
                                                                  price=price[i] + get[i],
                                                                  validity="DAY",
                                                                  product="MIS", )
            if order["sell_order_id"][i] != "0":
                t = kite.order_history(order["sell_order_id"][i])
                if t[-1]["status"] == "OPEN":
                    try:
                        csell_order_id = kite.cancel_order(variety="regular",
                                                           order_id=order["sell_order_id"][i],
                                                           )
                    except Exception as e:
                        print(e)
                        pass
                    # volume[i]=t[-1]["filled_quantity"]
                elif t[-1]["status"] == "COMPLETE" and order["ebuy_order_id"][i] == "0":
                    order["ebuy_order_id"][i] = kite.place_order(variety="regular",
                                                                 exchange=instruments[i][:
                                                                                         instruments[i].index(":")],
                                                                 order_type="LIMIT",
                                                                 tradingsymbol=instruments[i][
                                                                               instruments[i].index(":") + 1:],
                                                                 transaction_type="BUY",
                                                                 quantity=1,
                                                                 price=price[i] - get[i],
                                                                 validity="DAY",
                                                                 product="MIS", )

    while datetime.time(9, 32, 0, 0) <= datetime.datetime.now().time() < datetime.time(9, 33, 0, 0):
        for i in range(0, len(instruments)):
            if order["ebuy_order_id"][i] != "0":
                s = kite.order_history(order["ebuy_order_id"][i])
                if s[-1]["status"] == "OPEN":
                    cbuy_order_id = kite.cancel_order(variety="regular",
                                                      order_id=order["ebuy_order_id"][i],
                                                      )
                    # volume[i] = s[-1]["filled_quantity"]
                    ebuy_order_id = kite.place_order(variety="regular",
                                                     exchange=instruments[i][:
                                                                             instruments[i].index(":")],
                                                     order_type="MARKET",
                                                     tradingsymbol=instruments[i][
                                                                   instruments[i].index(":") + 1:],
                                                     transaction_type="BUY",
                                                     quantity=1,
                                                     validity="DAY",
                                                     product="MIS", )
            if order["esell_order_id"][i] != "0":
                t = kite.order_history(order["esell_order_id"][i])
                if t[-1]["status"] == "OPEN":
                    csell_order_id = kite.cancel_order(variety="regular",
                                                       order_id=order["esell_order_id"][i],
                                                       )
                    # volume[i] = s[-1]["filled_quantity"]
                    esell_order_id = kite.place_order(variety="regular",
                                                      exchange=instruments[i][:
                                                                              instruments[i].index(":")],
                                                      order_type="MARKET",
                                                      tradingsymbol=instruments[i][
                                                                    instruments[i].index(":") + 1:],
                                                      transaction_type="SELL",
                                                      quantity=1,
                                                      validity="DAY",
                                                      product="MIS", )

    while datetime.time(9, 33, 0, 0) <= datetime.datetime.now().time():
        if m == 1:
            workbook = xlsxwriter.Workbook('pytrade_chart.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.write("A1", "Instrument")
            worksheet.write("B1", "Start Price")
            worksheet.write("C1", "Average Price")
            worksheet.write("D1", "Price")
            worksheet.write("E1", "Get")
            worksheet.write("F1", "volume")
            worksheet.write("G1", "buy_order_id")
            worksheet.write("H1", "sell_order_id")
            for i in range(0, len(instruments)):
                worksheet.write(i + 1, 0, instruments[i])
                worksheet.write(i + 1, 1, st_price[i])
                worksheet.write(i + 1, 2, avg_price[i])
                worksheet.write(i + 1, 3, price[i])
                worksheet.write(i + 1, 4, get[i])
                worksheet.write(i + 1, 5, volume[i])
                worksheet.write(i + 1, 6, order["buy_order_id"][i])
                worksheet.write(i + 1, 7, order["sell_order_id"][i])
            workbook.close()
            print("Done!")
            ch=input()
            m = 0
            break
    if m == 0:
        break
