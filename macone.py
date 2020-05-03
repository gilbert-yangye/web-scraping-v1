#! python3.6

import openpyxl,requests,bs4

wb = openpyxl.load_workbook('List 3 scrabbed.xlsx')
for page in range(27,30):
        sheet = wb.get_sheet_by_name('Agent '+ str(page))
        MAX = str(sheet.max_row + 1)
        for i in range(2,int(MAX)):
                print('running '+str(i))
                res1 = requests.get('https://www.cea.gov.sg/public-register?category=Salesperson&mobile=' + str(sheet['A'+str(i)].value))#request reg page
                reg = bs4.BeautifulSoup(res1.text,"html.parser")
	
                try:
                        judgement = reg.select('td')
                        KEOcheck = reg.select('td > a')
                        KEOinfor = KEOcheck[0].get('onclick').replace('return ShowDetailForm("','')
                        KEOID = KEOinfor.replace('",  540, 800);','')
                        KEOlink = requests.get('https://www.cea.gov.sg/Custom/CEA/PublicRegister/Page/PublicRegisterDetail.aspx?UserId=' + KEOID)
                        try:
                                KEOCHECK = (KEOlink.text).index(' - [ KEO ]')
                                sheet['F'+str(i)]= str('KEO')
                        except:
                                pass
                        sheet['D'+str(i)] = judgement[0].getText()
                        sheet['E'+str(i)] = judgement[1].getText()
                        sheet['B'+str(i)]=str('Y')
                        res2 = requests.get('https://www.google.com.sg/search?q='+ judgement[1].getText())

                        googleweb = bs4.BeautifulSoup(res2.text,"html.parser")
                        link = googleweb.select('.r a')
                        numsearch = min(5, len(link))

                        for a in range (numsearch):
                                URL = str('https://www.google.com.sg'+ link[a].get('href'))
                                res3 = requests.get(URL)
                                try:
                                        check = (res3.text).index(str(sheet['A'+str(i)].value))
                                        sheet['C'+str(i)] = URL
                                        break
                                except:
                                        pass
                
                                
                                
                except:
                        sheet['B'+str(i)]='N'
                        continue
 
        wb.save('List 3 scrabbed.xlsx')
        print(sheet.title + ' finished & saved')
        
print('All finished')
