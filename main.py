import datetime
import math
import os
import sqlite3
from multiprocessing.dummy import Pool as ThreadPool

import requests
import xlsxwriter
from bs4 import BeautifulSoup

session = requests.session()


def get_proxy():
    print("Looking for a working proxy server")
    soup = BeautifulSoup(requests.get('https://www.us-proxy.org').text, 'lxml')
    table = soup.find('table', {'id': 'proxylisttable'})
    tbody = table.find('tbody')
    proxies = []
    for tr in tbody:
        columns = tr.find_all('td')
        if columns[2].text in 'US' and columns[4].text in 'anonymous' and columns[6].text in 'yes':
            proxies.append("https://" + columns[0].text + ":" + columns[1].text)
    for p in proxies:
        try:
            proxy_line = {'https': p}
            resp = requests.get('https://www.ncl.com/search_vacations?cruise=1&cruiseTour=0&cruiseHotel=0&cruiseHotelAir=0&flyCruise=0&numberOfGuests=4294953449&state=undefined&pageSize=10&currentPage=', proxies=proxy_line, timeout=10)
            if resp.ok:
                print("Found one!")
                return proxy_line
            else:
                print(p, "Not working")
        except requests.exceptions.ProxyError:
            print(p, "Not working")
        except requests.exceptions.ConnectTimeout:
            print(p, "Not working")
        except requests.exceptions.ReadTimeout:
            print(p, "Not working")

headers = {
    "authority": "www.ncl.com",
    "method": "GET",
    "path": "/search_vacations",
    "scheme": "https",
    "accept": "application/json, text/plain, */*",
    "connection": "keep-alive",
    "referer": "https://www.ncl.com",
    "cookie": "AkaUTrackingID=5D33489F106C004C18DFF0A6C79B44FD; AkaSTrackingID=F942E1903C8B5868628CF829225B6C0F; UrCapture=1d20f804-718a-e8ee-b1d8-d4f01150843f; BIGipServerpreprod2_www2.ncl.com_http=61515968.20480.0000; _gat_tealium_0=1; BIGipServerpreprod2_www.ncl.com_r4=1957341376.10275.0000; MP_COUNTRY=us; MP_LANG=en; mp__utma=35125182.281213660.1481488771.1481488771.1481488771.1; mp__utmc=35125182; mp__utmz=35125182.1481488771.1.1.utmccn=(direct)|utmcsr=(direct)|utmcmd=(none); utag_main=_st:1481490575797$ses_id:1481489633989%3Bexp-session; s_pers=%20s_fid%3D37513E254394AD66-1292924EC7FC34CB%7C1544560775848%3B%20s_nr%3D1481488775855-New%7C1484080775855%3B; s_sess=%20s_cc%3Dtrue%3B%20c%3DundefinedDirect%2520LoadDirect%2520Load%3B%20s_sq%3D%3B; _ga=GA1.2.969979116.1481488770; mp__utmb=35125182; NCL_LOCALE=en-US; SESS93afff5e686ba2a15ce72484c3a65b42=5ecffd6d110c231744267ee50e4eeb79; ak_location=US,NY,NEWYORK,501; Ncl_region=NY; optimizelyEndUserId=oeu1481488768465r0.23231006365903206",
    "Proxy-Authorization": "Basic QFRLLTVmZjIwN2YzLTlmOGUtNDk0MS05MjY2LTkxMjdiMTZlZTI5ZDpAVEstNWZmMjA3ZjMtOWY4ZS00OTQxLTkyNjYtOTEyN2IxNmVlMjlk"
}
proxy = get_proxy()
response = requests.get("https://www.ncl.com/search_vacations?cruise=1&cruiseTour=0&cruiseHotel=0&cruiseHotelAir=0&flyCruise=0&numberOfGuests=4294953449&state=undefined&pageSize=10&currentPage=", proxies=proxy)
tmpcruise_results = response.json()
tmpline = tmpcruise_results['meta']
total_record_count = tmpline['aggregate_record_count']
total_cruise_count = total_record_count
pool = ThreadPool(5)
pool2 = ThreadPool(5)
page = ''
counter = 1
total_page_count = math.ceil(int(total_cruise_count) / 12)
cruises = []
page_counter = 1
nao = 12
to_write = []
keys = []
urls = set()

while page_counter <= int(total_page_count):

    if page_counter == 1:
        url = "https://www.ncl.com/search_vacations?"
        urls.add(url)
        page_counter += 1
    else:
        url = "https://www.ncl.com/search_vacations?cruise=1&cruiseTour=0&cruiseHotel=0&cruiseHotelAir=0&Nao=" + str(
            nao) + ""
        urls.add(url)
        page_counter += 1
        nao += 12


def send_req(link):
    response = session.post(link, proxies=proxy, headers=headers).json()
    for line in response['results']:
        cruises.append(line)


urls = list(urls)
pool.map(send_req, urls)
pool.close()
pool.join()


def convert_date(unformated):
    splitter = unformated.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    if month == 'Jan':
        month = '1'
    elif month == 'Feb':
        month = '2'
    elif month == 'Mar':
        month = '3'
    elif month == 'Apr':
        month = '4'
    elif month == 'May':
        month = '5'
    elif month == 'Jun':
        month = '6'
    elif month == 'Jul':
        month = '7'
    elif month == 'Aug':
        month = '8'
    elif month == 'Sep':
        month = '9'
    elif month == 'Oct':
        month = '10'
    elif month == 'Nov':
        month = '11'
    elif month == 'Dec':
        month = '12'
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def calculate_days(date, duration):
    dateobj = datetime.datetime.strptime(date, "%m/%d/%Y")
    calculated = dateobj + datetime.timedelta(days=int(duration))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def get_from_code(dc):
    if dc == 'CARIBBEAN':
        return ['Carib', 'C']
    if dc == 'ALASKA':
        return ['Alaska', 'A']
    if dc == 'ASIA':
        return ['Exotics', 'O']
    if dc == 'CANADA_NEW_ENGL':
        return ['Canada/New England', 'NN']
    if dc == 'GRNDX':
        return ['Grand crossings', 'TBD']
    if dc == 'EUROPE':
        return ['Europe', 'E']
    if dc == 'HAWAII':
        return ['Hawaii', 'H']
    if dc == 'PACIFIC_COASTAL':
        return ['Pacific Coastal', 'PC']
    if dc == 'PANAMA_CANAL':
        return ['Panama Canal', 'T']
    if dc == 'SOUTH_AMERICA':
        return ['S. America', 'S']
    if dc == 'TRANSATLANTIC':
        return ['Transatlantic', 'X']
    if dc == 'BERMUDA':
        return ['Bermuda', 'BM']
    if dc == 'BAHAMAS_FLORIDA':
        return ['Bahamas', 'BH']
    if dc == 'MEXICAN_RIVIERA':
        return ['Mexico', 'M']
    if dc == 'AUSTRALIA':
        return ['Australia', 'AU']
    else:
        try:
            dests = []
            for p in dc:
                if p == "WEEKEND":
                    pass
                else:
                    dests.append(p)
            if "PANAMA_CANAL" in dests:
                return ['Panama Canal', 'T']
            if 'CUBA' in dests:
                return ['Cuba', 'C']
            return [dc, dc]
        except TypeError:
            return [dc, dc]


def split_carib_auto(ports, dc, dn):
    cu = []
    wc = []
    ec = []
    bm = []
    conn = sqlite3.connect('/home/fixxxer/PycharmProjects/PortsExplorer/ports.db')
    c = conn.cursor()
    c.execute("SELECT * FROM portlist WHERE destination_name='Cuba'")
    for row in c.fetchall():
        cu.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='West Carib'")
    for row in c.fetchall():
        wc.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='East Carib'")
    for row in c.fetchall():
        ec.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='Bermuda'")
    for row in c.fetchall():
        bm.append(row[0])
    c.close()
    conn.close()
    result = []
    iscu = False
    isec = False
    iswc = False
    ports_list = []
    for i in range(len(ports)):
        if i == 0:
            pass
        else:
            ports_list.append(ports[i])
    for element in cu:
        for p in ports_list:
            if p in element or element in p:
                iscu = True
    if not iscu:
        for element in wc:
            for p in ports_list:
                if p in element or element in p:
                    iswc = True
    if not iswc:
        for element in ec:
            for p in ports_list:
                if p in element or element in p:
                    isec = True
    if iscu:
        result.append("Cuba")
        result.append("C")
        result.append("CU")
        return result
    elif iswc:
        result.append("West Carib")
        result.append("C")
        result.append("WC")
        return result
    elif isec:
        result.append("East Carib")
        result.append("C")
        result.append("EC")
        return result
    else:
        result.append(dn)
        result.append(dc)
        result.append("")
        return result


def get_from_code2(dest, ports, dc, dn):
    if "CARIBBEAN" in dest:
        destination = split_carib_auto(ports, dc, dn)
        return [destination[0], destination[1]]
    if 'Cococay' in ports or "Nassau" in ports or "Bahama Island" in ports:
        return ['Bahamas', 'BH']
    else:
        return [dc, dn]


def split_europe_auto(ports, dn, dc):
    baltic = []
    eastern_med = []
    west_med = []
    baltic.append("SOU")
    baltic.append("HAU")
    baltic.append("FLM")
    baltic.append("AES")
    baltic.append("BGO")
    baltic.append("KWL")
    baltic.append("GNR")
    baltic.append("SVG")
    baltic.append("TOS")
    baltic.append("LKN")
    baltic.append("HVG")
    conn = sqlite3.connect('/home/fixxxer/PycharmProjects/PortsExplorer/ports.db')
    c = conn.cursor()
    c.execute("SELECT * FROM portlist WHERE destination_name='Baltics'")
    for row in c.fetchall():
        baltic.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='EastMed'")
    for row in c.fetchall():
        eastern_med.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='WestMed'")
    for row in c.fetchall():
        west_med.append(row[0])
    c.close()
    conn.close()
    for element in baltic:
        for p in ports:
            if p in element or element in p:
                return ['Baltic', 'E']
            elif ports[0] in element or element in ports[0]:
                return ['Baltic', 'E']

    for element in eastern_med:
        for p in ports:
            if p in element or element in p:
                return ['Eastern Med', 'E']

    for element in west_med:
        for p in ports:
            if p in element or element in p:
                return ['Western Med', 'E']
    return [dn, dc]


# def split_europe(ports, dn, dc):
#     baltic = ['Petropavlovsk', 'Bergen', 'Flam', 'Geiranger', 'Alesund',
#               'Stavanger', 'Skjolden', 'Stockholm', 'Helsinki',
#               'St. Petersburg', 'Tallinn', 'Riga', 'Warnemunde',
#               'Copenhagen', 'Kristiansand', 'Skagen', 'Fredericia',
#               'Rostock (Berlin)', 'Nynashamn', 'Oslo', 'Amsterdam',
#               'Reykjavik',
#               'Zeebrugge (Brussels)', 'Southampton']
#     eastern_med = ['Athens (Piraeus)', 'Katakolon', 'Dubrovnik', 'Mykonos',
#                    'Rhodes', 'Chania (Souda)', 'Koper', 'Split',
#                    'Santorini', 'Zadar', 'Corfu', 'Kotor']
#     west_med = ['Catania,Sicily', 'Ajaccio', 'Alicante', 'Barcelona', 'Bilbao',
#                 'Cadiz', 'Cannes', 'Cartagena', 'Florence / Pisa (Livorno)',
#                 'Fuerteventura', 'Funchal (Madeira)', 'Genoa', 'Gibraltar',
#                 'Ibiza', 'La Coruna', 'La Spezia', 'Lanzarote',
#                 'Las Palmas', 'Lisbon', 'Malaga', 'Marseille',
#                 'Messina (Sicily)', 'Montecarlo, Monaco', 'Naples', 'Nice',
#                 'Palma De Mallorca', 'Ponta Delgada', 'Portofino', 'Provence (Toulon)',
#                 'Ravenna', 'Sete', 'St. Peter Port', 'Tenerife',
#                 'Valencia', 'Valletta, Malta', 'Venice', 'Vigo']
#     europe = ['Rome (Civitavecchia)', 'Le Havre (Paris)', 'Akureyri',
#               'Belfast', 'Cherbourg', 'Cork (Cobh)', 'Dover',
#               'Dublin', 'Edinburgh', 'Greenock (Glasgow)', 'Inverness/Loch Ness',
#               'Lerwick/Shetland', 'Liverpool',
#               'Waterford (Dunmore E.)']
#
#     ports_visited = ports
#
#     ports_list = []
#     for i in range(len(ports_visited)):
#
#         if i == 0:
#             pass
#         else:
#             ports_list.append(ports_visited[i])
#     for element in baltic:
#         for p in ports_list:
#             if p in element or element in p:
#                 return ['Baltic', 'E']
#             elif ports_visited[0] in element or element in ports_visited[0]:
#                 return ['Baltic', 'E']
#
#     for element in eastern_med:
#         for p in ports_list:
#             if p in element or element in p:
#                 return ['Eastern Med', 'E']
#
#     for element in west_med:
#         for p in ports_list:
#             if p in element or element in p:
#                 return ['Western Med', 'E']
#
#     return [dn, dc]


def parse(c):
    vessel_name = c['ship_name']
    brochure_name = c['title']
    if "with Hotel Bundle" in brochure_name:
        return
    number_of_nights = c['duration']
    destination = c['destination_code']
    vessel_id = ''
    cruise_id = ''
    cruise_line_name = 'Norwegian Cruise Lines'
    itinerary_id = ''
    destination = get_from_code(destination)
    destination_name = destination[0]
    destination_code = destination[1]

    price_grid_url = c['price_grid_url']
    price_url = "https://www.ncl.com" + price_grid_url + ""
    page = session.post(price_url, headers=headers, proxies=proxy)
    cruise_results = page.json()
    for each in cruise_results['results']:
        key = each['Record']['Properties']['p_Package_ID']
        if key in keys:
            continue
        else:
            keys.append(key)
        sail_date = (convert_date(each['Record']['Properties']['p_Sail_Date']))
        return_date = (convert_date(each['Record']['Properties']['p_Sail_End_Date']))
        price_details = (each['Record']['stateroomPriceDetails'])
        if "INSIDE" in price_details:
            interior_bucket_price = price_details['INSIDE'][0]['leastPrice']
        else:
            interior_bucket_price = "N/A"
        if "BALCONY" in price_details:
            balcony_bucket_price = price_details['BALCONY'][0]['leastPrice']
        else:
            balcony_bucket_price = "N/A"
        if "OCEANVIEW" in price_details:
            oceanview_bucket_price = price_details['OCEANVIEW'][0]['leastPrice']
        else:
            oceanview_bucket_price = "N/A"
        if "MINISUITE" in price_details:
            suite_bucket_price = price_details['MINISUITE'][0]['leastPrice']
        else:
            if "SUITE" in price_details:
                suite_bucket_price = price_details['SUITE'][0]['leastPrice']
            else:
                suite_bucket_price = "N/A"
        try:
            if "Cruisetour" in brochure_name:
                continue
        except TypeError:
            continue
        ports = cruise_results['dimensions']['ShorexPortCode'].items()
        portlist = []
        for k, v in ports:
            portlist.append(v)
        if isinstance(destination_name, list):
            destination = get_from_code2(destination_name, portlist, destination_code, destination_name)
            destination_name = destination[0]
            destination_code = destination[1]
        if destination_name == 'Carib':
            destination = split_carib_auto(portlist, destination_code, destination_name)
            destination_name = destination[0]
            destination_code = destination[1]
        if 'Europe' in destination_name:
            dest = split_europe_auto(portlist, destination_name, destination_code)
            destination_code = dest[1]
            destination_name = dest[0]
        temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name, itinerary_id,
                brochure_name, number_of_nights, sail_date, return_date,
                interior_bucket_price, oceanview_bucket_price, balcony_bucket_price, suite_bucket_price]
        tmp2 = [temp]
        print(temp)
        to_write.append(tmp2)


pool2.map(parse, cruises)
pool2.close()
pool2.join()


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')

    now = datetime.datetime.now()
    path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Norwegian Cruise Line.xlsx'
    if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 25)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 20)
    worksheet.set_column("H:H", 50)
    worksheet.set_column("I:I", 20)
    worksheet.set_column("J:J", 20)
    worksheet.set_column("K:K", 20)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 25)
    worksheet.set_column("N:N", 20)
    worksheet.set_column("O:O", 20)
    worksheet.write('A1', 'DestinationCode', bold)
    worksheet.write('B1', 'DestinationName', bold)
    worksheet.write('C1', 'VesselID', bold)
    worksheet.write('D1', 'VesselName', bold)
    worksheet.write('E1', 'CruiseID', bold)
    worksheet.write('F1', 'CruiseLineName', bold)
    worksheet.write('G1', 'ItineraryID', bold)
    worksheet.write('H1', 'BrochureName', bold)
    worksheet.write('I1', 'NumberOfNights', bold)
    worksheet.write('J1', 'SailDate', bold)
    worksheet.write('K1', 'ReturnDate', bold)
    worksheet.write('L1', 'InteriorBucketPrice', bold)
    worksheet.write('M1', 'OceanViewBucketPrice', bold)
    worksheet.write('N1', 'BalconyBucketPrice', bold)
    worksheet.write('O1', 'SuiteBucketPrice', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship in data_array:
        for l in ship:
            column_count = 0
            for r in l:
                try:
                    if column_count == 0:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 1:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 2:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 3:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 4:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 5:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 6:
                        column_count += 1
                    elif column_count == 7:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 8:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 9:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 10:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 11:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 12:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 13:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 14:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                except ValueError:
                    worksheet.write_string(row_count, column_count, r, centered)
                    column_count += 1
            row_count += 1
    workbook.close()
    pass


write_file_to_excell(to_write)
