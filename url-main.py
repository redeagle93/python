from bs4 import BeautifulSoup
import requests, openpyxl


class MainWindow:
    def __init__(self):
        super().__init__()
        self.ip_list = []
        self.col = 1
        try:
            file = open("input.txt", 'r')
            for line in file:
                self.ip_list.append(line)
        except:
            print("Cant found File")

        self.get_link()

    def get_link(self):
        try:
            for item in self.ip_list:
                print(item)
                self.bing_url = "http://www.bing.com/search?q=ip%3A"+item+"&qs=n&form=QBLH&pq=ip%3A"+item+"&count=50"
                self.scrap(self.bing_url)
                self.col +=1
        except:
            print("IP "+item + "may be Not True")
            print("May be Execl file is open ")

    def scrap(self,bing):
        self.bing = bing
        self.book = openpyxl.load_workbook('result.xlsx')
        self.sheet = self.book.active
        rows = 1
        page = requests.get(bing)
        soup = BeautifulSoup(page.content, 'html.parser')
        for x in soup.findAll('cite'):
            print(x.text)
            self.value = x.text
            self.sheet.cell(row=rows, column=self.col).value = self.value
            rows +=1
        self.book.save('result.xlsx')

pri = MainWindow()