from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
from time import sleep
from os import path
import xlsxwriter


_RED = '\033[0;31m'
_GREEN = '\033[0;32m'
_YELLOW = '\033[0;33m'
_UNCOLOR = '\033[0m'

def kindleranker(_keywords):
    wd = uc.Chrome(version_main=102)
    try : 

        for _keyword in _keywords:
            _keyword.replace('\n','')
            _data_file = '/Users/call911/Desktop/_C_/Work/Automation/KDPBot/Output/AmzonData/KindleRanker/{}_data.xlsx'.format(_keyword)
            if path.exists(_data_file):
                continue

            workbook = xlsxwriter.Workbook(_data_file)
            worksheet = workbook.add_worksheet()

            border_format = workbook.add_format({'border': 1})
            head_format = workbook.add_format({'border': 1, 'bg_color': 'yellow'})

            print('{}[+] KDPBot :{} - Kindle Ranker - Book Ideas'.format(_GREEN, _UNCOLOR))

            try:
                wd.get('https://www.kindleranker.com/idea_search/{}/US'.format(_keyword))
            except:
                sleep(3)

            sleep(3)
            competition_char = wd.find_element(By.XPATH, '//*[@id="total_assessment"]/h1/span').text
            sleep(0.2)
            average_price = wd.find_element(By.XPATH, '/html/body/main/div[11]/div[1]/p').text
            sleep(0.2)
            average_review = wd.find_element(By.XPATH, '/html/body/main/div[11]/div[2]/p/span').get_attribute('data-original-title')#data-original-title="4.6"
            sleep(0.2)
            average_review_count = wd.find_element(By.XPATH, '/html/body/main/div[11]/div[3]/p').text
            sleep(0.2)
            average_pages = wd.find_element(By.XPATH, '/html/body/main/div[11]/div[4]/p').text
            sleep(0.2)
            average_title_words = wd.find_element(By.XPATH, '/html/body/main/div[11]/div[5]/p').text

            print('{}[+] KDPBot :{} 1er Table Data : \n\t- Keyword : {}\n\t- Competition : {}\n\t- Average Price : {}\n\t- Average Review : {}\n\t- Average Review Count : {}\n\t- Average Pages : {}\n\t- Average Title Words : {}'.format(
                _GREEN, _UNCOLOR, _keyword, competition_char, average_price, average_review, average_review_count, average_pages, average_title_words))

            # --- Start - Write 1er table --- #

            worksheet_row = 1
            worksheet_col = 1
            worksheet.write( worksheet_row,  worksheet_col,     'Keyword', head_format)
            worksheet.write( worksheet_row + 1,  worksheet_col,   _keyword.upper(), border_format)
            worksheet.write( worksheet_row,  worksheet_col + 1,     'Competition', head_format)
            worksheet.write( worksheet_row + 1,  worksheet_col + 1,     competition_char, border_format)
            worksheet.write( worksheet_row,  worksheet_col + 2,     'Average Price', head_format)
            worksheet.write( worksheet_row + 1,  worksheet_col + 2,     average_price, border_format)
            worksheet.write( worksheet_row,  worksheet_col + 3,     'Average Review', head_format)
            worksheet.write( worksheet_row + 1,  worksheet_col + 3,     average_review, border_format)
            worksheet.write( worksheet_row,  worksheet_col + 4,     'Average Review Count', head_format)
            worksheet.write( worksheet_row + 1,  worksheet_col + 4,     average_review_count, border_format)
            worksheet.write( worksheet_row,  worksheet_col + 5,     'Average Pages', head_format)
            worksheet.write( worksheet_row + 1,  worksheet_col + 5,     average_pages, border_format)
            worksheet.write( worksheet_row,  worksheet_col + 6,     'Average Title Words', head_format)
            worksheet.write( worksheet_row + 1,  worksheet_col + 6,     average_title_words, border_format)

            # --- End - Write 1er table --- #

            top5_selling_authors = wd.find_element(By.XPATH, '/html/body/main/div[12]/p').text
            sleep(0.2)
            top5_selling_publishers = wd.find_element(By.XPATH, ' /html/body/main/div[13]/p').text
            sleep(0.2)

            print('{}[+] KDPBot :{} 2nd Table Data : \n\t- Top 5 Selling Authors : {}\n\t- Top 5 Selling Publishers : {}'.format(
                _GREEN, _UNCOLOR, top5_selling_authors.replace(' ,   ','\n\t\t'), top5_selling_publishers.replace(' ,   ','\n\t\t')))


            # --- Start - Write 2nd table --- #

            worksheet_row = 5
            worksheet_col = 1
            worksheet.write( worksheet_row,  worksheet_col,     'Top 5 Selling Authors', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 2,    'Top 5 Selling Publishers', border_format)
            worksheet_row += 1
            for _au in top5_selling_authors.split(' ,   '):
                worksheet.write( worksheet_row,  worksheet_col,   _au, border_format)
                worksheet_row += 1
            worksheet_row = 5
            worksheet_row += 1
            for _au in top5_selling_publishers.split(' ,   '):
                worksheet.write( worksheet_row,  worksheet_col + 2,     _au, border_format)
                worksheet_row += 1
            
            # --- End - Write 2nd table --- #


            select_element = wd.find_element(By.XPATH, '//select[@name="dtBasicExample_length"]')
            sleep(0.2)
            wd.execute_script("arguments[0].scrollIntoView();", select_element)
            sleep(0.2)
            select = Select(select_element)
            select.select_by_index(len(select.options)-1)
            sleep(2)


            _pages = wd.find_elements(By.XPATH, '//*[@id="dtBasicExample_wrapper"]/div[3]/ul/li')
            sleep(0.2)
            last_page = int(_pages[-1].text)

            print('{}[+] KDPBot :{} 3th Table Data '.format(_GREEN, _UNCOLOR))
            # --- Start - Write 3th table --- #

            worksheet_row = 13
            worksheet_col = 1
            worksheet.write( worksheet_row,  worksheet_col,     'Category Name', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 1,     'Best Seller Rank', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 2,     '50th Book\'s Rank', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 3,     'Median Sales', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 4,     'Median Price', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 5,     'Volatility', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 6,     'New Releases', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 7,     'Self Pub', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 8,     'KDP Select', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 9,     'Competition', head_format)

            for page in range(1,last_page+1):
                _elements = wd.find_elements(By.XPATH, '//*[@id="dtBasicExample"]/tbody/tr')
                sleep(0.25)
                for _elem in _elements:
                    wd.execute_script("arguments[0].scrollIntoView();", _elem)
                    sleep(0.2)
                    worksheet_row += 1
                    worksheet_col = 1
                    _infos = _elem.find_elements(By.XPATH, './/td')
                    sleep(0.2)
                    print('\t\t{}'.format(_elem.text))
                    
                    for _info in _infos:
                        worksheet.write( worksheet_row,  worksheet_col,   _info.text, border_format)
                        worksheet_col += 1
                        #print('\n\t\t{}'.format(_info.text))
                
                is_next = True
                try:
                    next_page = wd.find_element(By.XPATH,'//li[@class="paginate_button page-item active"]//following-sibling::li')
                    sleep(0.2)
                except NoSuchElementException as e:
                    is_next = False
                    break
                    
                if is_next:
                    try:
                        next_page.click()
                        sleep(0.2)
                    except ElementNotInteractableException:
                        break

            # --- End - Write 3th table --- #

            print('{}[+] KDPBot :{} 4th Table Data '.format(_GREEN, _UNCOLOR))

            # --- Start - Write 4th table --- #
            worksheet_row = 1
            worksheet_col = 12
            worksheet.write( worksheet_row,  worksheet_col,   "Related Ideas", head_format)
            worksheet_row += 1
            related_keywords = wd.find_element(By.XPATH,'//*[@id="related_keywords"]').text.split('\n')
            sleep(0.2)
            for related_keyword in related_keywords:
                worksheet.write( worksheet_row,  worksheet_col,   related_keyword, border_format)
                worksheet_row += 1
                print('\t\t{}'.format(related_keyword))

            wd.quit()
            sleep(2)
            wd = uc.Chrome(version_main=102)

            # --- End - Write 4th table --- #

            print('{}[+] KDPBot :{} - Kindle Ranker - Keywords'.format(_GREEN, _UNCOLOR))

            try:
                wd.get('https://www.kindleranker.com/keyword_search/{}/books/US'.format(_keyword))
            except:
                sleep(3)

            while True:
                try:
                    wd.find_element(By.XPATH, '//span[@id="me-result-number"]')
                    sleep(0.2)
                    break
                except NoSuchElementException as e:
                    sleep(0.1)
            
            while wd.find_element(By.XPATH, '//*[@id="related_keywords"]').text == "":
                sleep(0.1)
            

            sleep(2)

            print('{}[+] KDPBot :{} 5th Table Data :\n\t- Suggestions : '.format(_GREEN, _UNCOLOR))
            _suggestions = wd.find_element(By.XPATH, '//*[@id="related_keywords"]').text
            sleep(0.2)

            # --- Start - Write 5th table --- #
            worksheet_row = 1
            worksheet_col = 14
            worksheet.write( worksheet_row,  worksheet_col, 'Suggestions', head_format)
            for _word in _suggestions.split(',  '):
                worksheet.write( worksheet_row + 1,  worksheet_col, _word, border_format)
                worksheet_row += 1
                print('\t\t{}'.format(_word))
            
            # --- End - Write 5th table --- #

            _count = wd.find_element(By.XPATH, '//span[@id="me-result-number"]').text
            sleep(0.2)
            select = Select(wd.find_element(By.XPATH, '//select[@name="dtBasicExample_length"]'))
            sleep(0.2)
            select.select_by_index(len(select.options)-1)
            sleep(2)
            _pages = wd.find_elements(By.XPATH, '//*[@id="dtBasicExample_wrapper"]/div[4]/ul/li')
            last_page = int(_pages[-1].text)

            print('{}[+] KDPBot :{} th Table Data '.format(_GREEN, _UNCOLOR))
            

             # --- Start - Write 6th table --- #
            worksheet_row = 13
            worksheet_col = 14
            worksheet.write( worksheet_row,  worksheet_col,     'Keywords', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 1,     'Search Volume', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 2,     'Competing books', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 3,     'Books broadly related', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 4,     'Competing authors', head_format)
            worksheet.write( worksheet_row,  worksheet_col + 5,     'Median Monthly Sales', head_format)
            for page in range(1,last_page+1):
                _elements = wd.find_elements(By.XPATH, '//tbody[@id="me-tbody"]/tr')
                sleep(0.2)
                for _elem in _elements:
                    wd.execute_script("arguments[0].scrollIntoView();", _elem)
                    sleep(0.2)
                    worksheet_row += 1
                    worksheet_col = 14
                    _infos = _elem.find_elements(By.XPATH, './/td')
                    sleep(0.2)
                    print('\t\t{}'.format(_elem.text))
                    for _info in _infos:
                        worksheet.write( worksheet_row,  worksheet_col,   _info.text, border_format)
                        worksheet_col += 1
                        #print('\n\t\t{}'.format(_info.text))
                        

                is_next = True
                try:
                    next_page = wd.find_element(By.XPATH,'//li[@class="paginate_button page-item active"]//following-sibling::li')
                    sleep(0.2)
                except NoSuchElementException as e:
                    workbook.close()
                    wd.quit()
                    sleep(2)
                    is_next = False
                    wd = uc.Chrome(version_main=102)
                    break

                    
                if is_next:
                    try:
                        next_page.click()
                        sleep(0.2)
                    except ElementNotInteractableException:
                        workbook.close()
                        wd.quit()
                        sleep(2)
                        wd = uc.Chrome(version_main=102)
                        break
                        

    except Exception as ex:
        print(ex)
        workbook.close()
        wd.quit()
    
    wd.quit()


keywords = open('keywords.txt', 'r').readlines()
kindleranker(keywords)