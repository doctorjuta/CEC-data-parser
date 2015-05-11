def get_dep_datas(link, worksheet_list, worksheet_vo, row, base_url, BeautifulSoup, urlopen):
    content = urlopen(link).read()
    soup = BeautifulSoup(content)
    tds = soup.select('.td10 a')
    links_num = [link]
    if len(tds) > 0:
        for lin in tds:
            links_num.append(base_url+lin['href'].replace("¤", "&curren"))
    for lin in links_num:
        content = urlopen(lin).read()
        soup = BeautifulSoup(content)
        tables = soup.find_all('table', 't2')
        tables_rows = tables[1].find_all('tr')
        tables_rows.pop(0)
        for rows in tables_rows:
            dep_name = rows.find('a').get_text()
            dep_party = rows.find('b').get_text().split(',')
            if "ОВО" in dep_party[1]:
                worksheet_vo.cell(row=row[1],column=1,value=dep_name)
                worksheet_vo.cell(row=row[1],column=2,value=dep_party[0])
                worksheet_vo.cell(row=row[1],column=3,value=dep_party[1])
                row[1] += 1
            else:
                worksheet_list.cell(row=row[0],column=1,value=dep_name)
                worksheet_list.cell(row=row[0],column=2,value=dep_party[0])
                worksheet_list.cell(row=row[0],column=3,value=dep_party[1])
                row[0] += 1
            row[2] += 1
            print("Some datas wrote in xlsx files in row number "+str(row[2]))
    return row

def count_dep_list(worksheet_part_list, worksheet_list):
    uniq_pary = {}
    tmp_row = 1
    for row in worksheet_list.rows:
        targ_date = worksheet_list.cell(row=tmp_row,column=2).value
        if targ_date in uniq_pary:
            uniq_pary[targ_date] += 1
        else:
            uniq_pary[targ_date] = 1
        tmp_row += 1
    tmp_row = 1
    for part, cnt in uniq_pary.items():
        worksheet_part_list.cell(row=tmp_row,column=1,value=str(part))
        worksheet_part_list.cell(row=tmp_row,column=2,value=str(cnt))
        tmp_row += 1
    return tmp_row