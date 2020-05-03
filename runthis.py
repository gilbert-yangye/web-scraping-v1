#! python3.6
# Web scrapping for candh.xlsx
#instruction:#input data :column A of excel--phone number
            #output data1:column B of excel--check whether the number is registered on 'https://www.cea.gov.sg/public-register' or not
            #output data2:column c of excel--find the relevant personal page on google and paste it on column C
         
#input excel data into list

import openpyxl,pprint,requests,sys,os,bs4,webbrowser,selenium

wb = openpyxl.load_workbook('List 2.xlsx')
sheet = wb.get_sheet_by_name('Agent 12')

from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
dcap = dict(DesiredCapabilities.PHANTOMJS)


dcap["phantomjs.page.settings.userAgent"] = ('Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36')
dcap["phantomjs.page.settings.loadImages"] = False
service_args = ['--proxy=127.0.0.1:9999','--proxy-type=socks5']
completepage = webdriver.PhantomJS(executable_path=r'D:\Python trail\explore\phantomjs-2.1.1-windows\bin\phantomjs.exe',desired_capabilities = dcap)
completepage.set_page_load_timeout(10)
completepage.set_script_timeout(10)

print('processing')
for i in range(755,757):
        print('running '+str(i))
        completepage.get('https://www.cea.gov.sg/public-register?category=Salesperson&mobile='+ str(sheet['A'+str(i)].value))#request reg page

#        completepage.get('https://www.cea.gov.sg/public-register?category=Salesperson&mobile=' + str(sheet['A'+str(i)].value))#request reg page
        print('2')
        completesource = completepage.page_source
        print(completesource)
        print('4')
        reg = bs4.BeautifulSoup(completesource,"html.parser")
        print(reg.select('div')[0].getText())
	
        try:
                judgement = reg.select('td')
                print('okay yet1')
                resceapath = reg.select('iframe')
                print(str(len(resceapath)))
                havenoidea = str(resceapath[0].get('onclick'))
                print(str(havenoidea.remove('return ShowDetailForm("')))#  '",  540, 800);')))
                rescea = requests.get('https://www.cea.gov.sg' + resceapath[0].get('src'))
                print('successful request')
                try:
                        rescea.raise_for_status()
                except Exception as exc:
                        continue
                try:
                        KEOcheck = (res3.text).index(' - [ KEO ]')
                        sheet['F'+str(i)] = str('KEO')
                except:
                        continue
                              
                sheet['D'+str(i)] = judgement[0].getText()
                sheet['E'+str(i)] = judgement[1].getText()
                sheet['B'+str(i)]=str('Y')
                
                res2 = requests.get('https://www.google.com.sg/search?q='+ judgement[1].getText())
                try:
                        res2.raise_for_status()
                except Exception as exc:
                        continue

                googleweb = bs4.BeautifulSoup(res2.text,"html.parser")
                link = googleweb.select('.r a')
                numsearch = min(6, len(link))
 
                
                for a in range (numsearch):
                        res3 = requests.get('https://www.google.com.sg'+ link[a].get('href'))
                        try:
                                res3.raise_for_status()
                        except Exception as exc:
                                continue
                        try:
                                check = (res3.text).index(str(sheet['A'+str(i)].value))
                                sheet['C'+str(i)] = str('https://www.google.com.sg'+ link[a].get('href'))
                                break
                        except:
                                pass
                                            
                        try:
                                check = (res3.text).index(str(sheet['D'+str(i)].value))
                                sheet['C'+str(i)] = str('https://www.google.com.sg'+ link[a].get('href'))
                                break
                        except:
                                pass
                        try:
                                check = (res3.text).index(str(sheet['E'+str(i)].value))
                                sheet['C'+str(i)] = str('https://www.google.com.sg'+ link[a].get('href'))
                                break
                        except:
                                pass
                
                                
                                
        except:
                sheet['B'+str(i)]='N'
                print(str(i)+' '+'not exist')
                continue
        
        print(str(i)+' '+'finished')
 
wb.save('List 2 copy.xlsx')
completepage.quit()
print('finished & saved')
