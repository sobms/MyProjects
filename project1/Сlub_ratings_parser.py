import requests
import os
import sys
from bs4 import BeautifulSoup
import socks
import pandas as pd
import numpy as np
import googlemaps
import design2
from PyQt5 import QtWidgets

class App(QtWidgets.QMainWindow, design2.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.work)
        self.pushButton_2.clicked.connect(self.open_file)
        
    def get_result(self, url):
        rs = requests.session()
        rs.headers['User-Agent'] = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'
        rs.headers['Accept-Language'] = 'Accept-Language: en-US,en;q=0.5'
        result = rs.get(url, timeout=1000)
        rs.close()
        return result 

    def get_work_status_yandex(self, html):
        working_status = html.find('div', class_ = 'orgpage-header-view__working-status').text
        return working_status

    def get_organization_type_yandex(self, html):
        org_statuses = [elem.text for elem in html.find_all('a', class_ = 'breadcrumbs-view__breadcrumb')]
        return org_statuses

    def get_rating_yandex(self, html):
        rating = html.find('div', class_='orgpage-header-view__rating-wrapper').find_all('span',
                                                                                     class_='business-rating-badge-view__rating-text _size_m')
        if len(rating) != 0:
            rating = rating[0].text
        else:
            rating = "No data"
        return rating
    
    def get_results_google(self, place_id, gmaps):
        if place_id is not np.nan:
            try:
                result=gmaps.place(place_id=place_id,fields=['name', 'rating', 'business_status', 'type'])
            except:
                rating = "No data"
                working_status = "No data"
                organization_type_is_correct = "No data"
            try:
                rating=result['result']['rating']
            except:
                rating = "No data"
            try:
                working_status = result['result']['business_status']
            except:
                working_status = "No data"
            try:
                if ('gym' not in result['result']['types']):
                    organization_type_is_correct = "ERROR"
                else:
                    organization_type_is_correct = "ok"
            except:
                organization_type_is_correct = "No data"
        else:
            rating = "No data"
            working_status = "No data"
            organization_type_is_correct = "No data"
        return [rating,working_status,organization_type_is_correct]
    

    def work(self):
        adr = self.lineEdit.text()
        print(adr)
        table = pd.read_excel(adr)
        
        rating_yandex = []
        working_status_yandex = []
        organization_type_is_correct_yandex = []
        for url in table['URL_yandex']:
            if url is not np.nan:
                response = self.get_result(url)
                soup = BeautifulSoup(response.text,'lxml')
                rating_yandex.append(self.get_rating_yandex(soup))
                working_status_yandex.append(self.get_work_status_yandex(soup))
                types = self.get_organization_type_yandex(soup)
                if (('Фитнес-клубы' not in types) and ('Спортзалы' not in types)):
                    organization_type_is_correct_yandex.append('ERROR')
                else:
                    organization_type_is_correct_yandex.append('ok')
            else:
                rating_yandex.append("no_data")
                working_status_yandex.append("no_data")
                organization_type_is_correct_yandex.append("no_data")
        
        file = open('key.txt', 'r')
        API_TOKEN= file.read()
        file.close()
        gmaps = googlemaps.Client(key=API_TOKEN)

        rating_google = []
        working_status_google = []
        organization_type_is_correct_google = []
        for place_id in table['place_id_google']:
            result = self.get_results_google(place_id, gmaps)
            rating_google.append(result[0])
            working_status_google.append(result[1])
            organization_type_is_correct_google.append(result[2])

        table['rating_yandex'] = rating_yandex
        table['rating_google'] = rating_google
        table['working_status_yandex'] = working_status_yandex
        table['organization_type_is_correct_yandex_(Корректно, если type = Фитнес-клубы или Спортзалы)'] = organization_type_is_correct_yandex
        table['working_status_google'] = working_status_google
        table['organization_type_is_correct_google'] = organization_type_is_correct_google

        table.to_excel('rating_list.xlsx')
        print('File is ready!')

    def open_file(self):
        os.system('rating_list.xlsx')

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = App()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()





   