#! python3.6
# Web scrapping for candh.xlsx
#instruction:#input data :column A of excel--phone number
            #output data1:column B of excel--check whether the number is registered on 'https://www.cea.gov.sg/public-register' or not
            #output data2:column c of excel--find the relevant personal page on google and paste it on column C
         
#input excel data into list

import openpyxl,pprint,requests,sys,os,bs4,webbrowser


wb = openpyxl.load_workbook('List 2.xlsx')
sheet = wb.get_sheet_by_name('Agent 13')



#creat a loop
print('processing')
for i in range(2,12):
        res1 = requests.get('https://www.cea.gov.sg/public-register?category=Salesperson&mobile=' + str(sheet['A'+str(i)].value))#request reg page
        res1.raise_for_status()

        reg = bs4.BeautifulSoup(res1.text,"html.parser")#pass to beautifulsoup4 and a string
        sheet['B'+str(i)]='N'#setdefault
        sheet['C'+str(i)]='None'
	
        try:#see if registered
                #needtobechecked:
                judgement = reg.select('td')
                sheet['B'+str(i)]=str('Y')
                res2 = requests.get('https://www.google.com.sg/search?q='+ judgement[0].getText())#request('https://www.google.com.sg/search?q='+ sheet['A'+i].)
                res2.raise_for_status()
                googleweb = bs4.BeautifulSoup(res2.text,"html.parser")#pass to beautifulsoup4 and a string
                link = googleweb.select('cite')
                #needtobechecked:linkshortcut = list[judgement[].getText()]
                numsearch = min(6, len(link))

                #creat a scraping loop
                for a in range (numsearch):
                        res3 = requests.get(link[a].getText())
                        targetweb = bs4.BeautifulSoup(res3.text,"html.parser")#pass to beautifulsoup4 and targetweb
                        if sheet['A'+str(i)] in targetweb:
                                sheet['C'+str(i)] = str(link[a].getText())
                                print('before if break')
                                break
                        else:
                                print('run else')
                                continue
                        break
                                
                                
        except:
                print('continue')
                continue
        print('in process'+str(i))
        break
print('the for loop broken')

wb.save('List 2 copy.xlsx')
print('finished')
