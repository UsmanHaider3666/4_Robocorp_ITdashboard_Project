import datetime

from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
import time
from RPA.PDF import PDF
import os
import shutil


class CollectData:

    def __init__(self):
        self.browse = Selenium()
        self.url = "https://itdashboard.gov/"
        self.http = HTTP()
        self.excel = Files()
        self.pdf = PDF()
        self.name = []
        self.agencies_name = []
        self.agencies_investment = []
        self.final_investment = []
        self.individual = []
        self.UII = []
        self.Bureau = []
        self.Investment_Title = []
        self.Total_Spending = []
        self.Type = []
        self.CIO = []
        self.Projects = []
        self.num1 = ""
        self.links = []
        self.browse.set_download_directory('/home/usman/Python-RPA/RobocorpProject4/output1')
        self.file_name = []
        self.title = []
        self.id = []
        self.dir_path = ""

    def clear_output1(self):
        self.dir_path = '/home/usman/Python-RPA/RobocorpProject4/output1'
        if os.path.exists(self.dir_path) and os.path.isdir(self.dir_path):
            shutil.rmtree(self.dir_path)

    def open_browser(self):
        self.browse.open_available_browser(self.url, maximized=True, )
        self.browse.click_element_when_visible('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
        self.browse.auto_close = False

    def create_excel_file(self):
        self.excel.create_workbook("data.xlsx", fmt="xlsx")
        self.excel.save_workbook("data.xlsx")

    def extract_data(self):
        name = self.browse.find_elements('//span[@class="h4 w200"]')
        for i in name:
            nam1 = i.text
            if nam1 in self.agencies_name:
                break
            else:
                self.agencies_name.append(nam1)
        investment = self.browse.find_elements('//span[@class=" h1 w900"]')
        for i in investment:
            nam = i.text
            self.agencies_investment.append(nam)
        middle = len(self.agencies_investment) // 2
        self.final_investment = self.agencies_investment[:middle]

    def put_data_to_excel(self):
        dic = {"name": self.agencies_name, "investment": self.final_investment}
        self.excel.open_workbook("data.xlsx")
        self.excel.append_rows_to_worksheet(dic, header=True)
        self.excel.save_workbook()

    def create_individual_excel_file(self):
        self.excel.create_workbook("individual.xlsx", fmt="xlsx")
        self.excel.save_workbook("individual.xlsx")

    def collect_individual_data(self):
        self.browse.click_element_when_visible(
            '//*[@id="agency-tiles-widget"]/div/div[4]/div[2]/div/div/div/div[1]/a/span[1]')
        time.sleep(5)
        self.browse.click_element_when_visible('//*[@id="investments-table-object_length"]/label/select/option[4]')
        time.sleep(10)
        # self.browse.wait_until_element_is_visible('',timeout=datetime.timedelta(seconds=10))
        number = self.browse.find_element('//*[@id="investments-table-object_info"]')
        num = number.text
        self.num1 = num.split(' ')[-2]
        for i in range(1, 8):
            a = self.browse.find_element(
                f'//*[@id="investments-table-object_wrapper"]/div[3]/div[1]/div/table/thead/tr[2]/th[{i}]')
            self.individual.append(a.text)
        for j in range(1, int(self.num1) + 1):
            for x in range(1, 8):
                ans = self.browse.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{j}]/td[{x}]')
                if x == 1:
                    self.UII.append(ans.text)
                elif x == 2:
                    self.Bureau.append(ans.text)
                elif x == 3:
                    self.Investment_Title.append(ans.text)
                elif x == 4:
                    self.Total_Spending.append(ans.text)
                elif x == 5:
                    self.Type.append(ans.text)
                elif x == 6:
                    self.CIO.append(ans.text)
                elif x == 7:
                    self.Projects.append(ans.text)
        print(self.UII)
        print(self.Investment_Title)

    def put_individual_data(self):
        dic = {f'{self.individual[0]}': self.UII, f'{self.individual[1]}': self.Bureau,
               f'{self.individual[2]}': self.Investment_Title,
               f'{self.individual[3]}': self.Total_Spending, f'{self.individual[4]}': self.Type,
               f'{self.individual[5]}': self.CIO,
               f'{self.individual[6]}': self.CIO}
        self.excel.open_workbook("individual.xlsx")
        self.excel.append_rows_to_worksheet(dic, header=True)
        self.excel.save_workbook()

    def get_links(self):
        for j in range(1, int(self.num1) + 1):
            try:
                for x in range(1, 2):
                    link = self.browse.get_element_attribute(
                        f'//*[@id="investments-table-object"]/tbody/tr[{j}]/td[{x}]/a', attribute="href")
                    self.links.append(link)
            except:
                pass

    def get_pdf(self):
        for i in self.links:
            self.browse.go_to(f'{i}')
            time.sleep(5)
            self.browse.click_element_when_visible('//*[@id="business-case-pdf"]/a')
            time.sleep(10)

    def get_pdf_file_name(self):
        for filename in os.listdir("/home/usman/Python-RPA/RobocorpProject4/output1/"):
            if filename.endswith(".pdf"):
                self.file_name.append(filename)
            else:
                continue

    def read_pdf(self):
        for i in self.file_name:
            try:
                text = self.pdf.get_text_from_pdf(f'/home/usman/Python-RPA/RobocorpProject4/output1/{i}', pages="1")
                new_list = list(text.values())
                new_string = "".join(new_list)
                investment = new_string.split(':', 13)[-2]
                investment_title = investment.split('2')[-2]
                self.title.append(investment_title)

                investment2 = new_string.split(':', 14)[-2]
                investment_id = investment2.split('S')[-2]
                self.id.append(investment_id)
            except:
                pass
        print(self.id)
        print(len(self.id))
        print(self.title)
        print(len(self.title))

    def compare_values(self):
        for i in self.id:
            print(f"checking for {i}")
            if i in self.UII:
                print(f"{i} is present in UII")
        for j in self.title:
            print(f"checking for {j}")
            if j in self.Investment_Title:
                print(f"{j} is present in Investment_Title")


if __name__ == '__main__':
    res = CollectData()
    res.clear_output1()
    res.open_browser()
    res.create_excel_file()
    res.extract_data()
    res.put_data_to_excel()
    res.create_individual_excel_file()
    res.collect_individual_data()
    res.put_individual_data()
    res.get_links()
    res.get_pdf()
    res.get_pdf_file_name()
    res.read_pdf()
    res.compare_values()
