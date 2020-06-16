while True:

    import xlwings as xw
    import requests
    import time
    import pprint
    import json

    excel = input("Provide the complete path of Excel file : ")
    wb = xw.Book(excel)
    sht = wb.sheets[0]

    exp = input("Provide the expiry  : ")
    start = int(input("Write the no of row from which you want to start:  "))
    end = int(input("Write the no of row at which you want to end:  "))

    sht.range("b1").value = "VOLUME"
    sht.range("C1").value = "AVG IV"
    sht.range("D1").value = "CE IV"
    sht.range("E1").value = "PE IV"
    sht.range("g1").value = "TIME TAKEN"

    for l in range(start, end + 1):
        start_time = time.time()
        stock = sht.range("a" + str(l)).value
        print(l- start + 1, ":", stock)
        symbol = {"symbol": stock}
        expiry = exp
        url = 'https://www.nseindia.com/api/option-chain-equities?'
        headers = {'User-Agent': 'Mozilla/5.0'}
        res = requests.get(url, headers= headers, params= symbol)
        res.raise_for_status()
        opt_chain = json.loads(res.text)
        
        try:
            number = len(opt_chain["records"]["data"])
        except KeyError:
            number = 0

        if number > 0:
            try: 
                ltp = opt_chain["records"]["data"][0]["CE"]["underlyingValue"]
            except KeyError:
                ltp = opt_chain["records"]["data"][0]["PE"]["underlyingValue"]

            strike_list = []
            for i in range(0, number):
                expiry_date = opt_chain["records"]["data"][i]["expiryDate"]
                if expiry_date == expiry:
                    strikes = opt_chain["records"]["data"][i]["strikePrice"] 
                    strike_list.append(strikes)

            k_list = []

            for k in range(0, number):
                expiry_date = opt_chain["records"]["data"][k]["expiryDate"]
                if expiry_date == expiry:
                    k_list.append(k)

            diff_list = []
            for j in range(0, len(strike_list)):
                diff = ltp - strike_list[j]
                diff_list.append(abs(diff))
                small = min(diff_list)

            loc = diff_list.index(small)

            if strike_list[loc] > ltp:
                otm_strike_ce = strike_list[loc]
                otm_strike_pe = strike_list[loc-1]

            else:
                otm_strike_ce = strike_list[loc + 1]
                otm_strike_pe = strike_list[loc]

            vol = []
            
            for o in range(0, len(k_list)):
                strike = opt_chain["records"]["data"][k_list[o]]["strikePrice"]
                if strike == otm_strike_pe:
                    for p in range(o-6, o+6):
                        try:
                            try:
                                pe_vol = opt_chain["records"]["data"][k_list[p]]["PE"]["totalTradedVolume"]
                            except KeyError:
                                pe_vol = 0
                        except IndexError:
                            pe_vol = 0
                        vol.append(pe_vol)
                    for m in range(o, 0, -1):
                        try:
                            pe_iv = opt_chain["records"]["data"][k_list[m]]["PE"]["impliedVolatility"]
                            if pe_iv != 0:
                                break
                        except KeyError:
                            pe_iv = 0
                elif strike == otm_strike_ce:
                    for q in range(o-6, o+6):
                        try:
                            try:
                                ce_vol = opt_chain["records"]["data"][k_list[q]]["CE"]["totalTradedVolume"]
                            except KeyError:
                                ce_vol = 0
                        except IndexError:
                            ce_vol = 0
                        vol.append(ce_vol)
                    for n in range(o, len(k_list)):
                        try:
                            ce_iv = opt_chain["records"]["data"][k_list[n]]["CE"]["impliedVolatility"]
                            if ce_iv != 0:
                                break
                        except KeyError:
                            ce_iv = 0
                        
            sht.range("b" + str(l)).value = max(vol) 
            sht.range("d" + str(l)).value = ce_iv
            sht.range("e" + str(l)).value = pe_iv
            sht.range("i" + str(l)).value = ltp
            sht.range("g" + str(l)).value = time.time() - start_time

        else:
            sht.range("b" + str(l)).value = 0 
            sht.range("d" + str(l)).value = 0
            sht.range("e" + str(l)).value = 0
            sht.range("i" + str(l)).value = 0
            sht.range("g" + str(l)).value = time.time() - start_time

    while True:

        answer = input(str('Run again? (y/n): '))
        if answer in ("y", "n"):
            break
        else:
            print("Ïnvalid Input")

    if answer == 'y':
        continue
    else:
        break

