from datetime import timedelta
import os
from config import OUTPUT, create_directory, URL
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF


class DashBoard:

    def __init__(self):
        create_directory("output")
        self.files = Files()
        self.pdf = PDF()
        self.browser = Selenium()
        self.browser.set_download_directory(OUTPUT)
        self.browser.open_available_browser(URL, maximized=True)
        self.urls = []

    def dive_in_agencies(self):

        """
        This function is used to click "Dive In" scrape name and Total FYI 2021 Spending and
        add the data in Excel sheet
        :return: None

        """

        agencies = []
        amounts = []

        self.browser.wait_until_page_contains('DIVE IN')
        self.browser.find_element('//a[@aria-controls="home-dive-in"]').click()
        self.browser.wait_until_page_contains_element("//div[@id='agency-tiles-container']")
        agency_list = self.browser.find_elements('//span[@class="h4 w200"]')
        amount_list = self.browser.find_elements('//span[@class=" h1 w900"]')

        for agency in agency_list:
            if agency.text != "":
                agencies.append(agency.text)
        for amount in amount_list:
            amounts.append(amount.text)

        excel_data = {'Agencies': agencies, 'Amounts': amounts}
        self.files.create_workbook(os.path.join(os.getcwd(), 'output/data.xlsx'))
        self.files.append_rows_to_worksheet(content=excel_data, header=True)
        self.files.rename_worksheet("Sheet", "Agencies")
        self.files.save_workbook()
        self.files.close_workbook()

    def agency_table_scrape(self, agency_name):
        """

        This function perform click operation on single agency and scroll down to
        Individual Investment table scrape all the table and add the data to an Excel Sheet

        :return: None

        """

        uii_list = []
        bureau_list = []
        investment_title_list = []
        spending_list = []
        type_list = []
        cio_rating_list = []
        no_of_projects_list = []
        pdf_matched = []
        self.browser.wait_until_page_contains_element(f'//img[@alt="Seal of the {agency_name}"]')
        self.browser.find_element(f'//img[@alt="Seal of the {agency_name}"]').click()
        self.browser.wait_until_page_contains_element('//table[@id="investments-table-object"]', timeout=timedelta(seconds=20))
        self.browser.select_from_list_by_value('//div[@class="dataTables_length"]//label//select', "-1")
        row_range = self.browser.find_element('//div[@id="investments-table-object_info"]').text
        row_range_int = int(row_range.split(' ')[5])
        self.browser.wait_until_page_contains_element(f'//table[@id="investments-table-object"]//tbody//tr[{row_range_int}]',
                                                      timeout=timedelta(seconds=20))
        for i in range(1, row_range_int + 1):
            uii = self.browser.find_element(f'//table[@id="investments-table-object"]//tbody//tr[{i}]//td[1]')
            uii_list.append(uii.text)
            bureau = self.browser.find_element(f'//table[@id="investments-table-object"]//tbody//tr[{i}]//td[2]')
            bureau_list.append(bureau.text)
            investment_title = self.browser.find_element(f'//table[@id="investments-table-object"]//tbody//tr[{i}]/td[3]')
            investment_title_list.append(investment_title.text)
            spending = self.browser.find_element(f'//table[@id="investments-table-object"]//tbody//tr[{i}]/td[4]')
            spending_list.append(spending.text)
            types = self.browser.find_element(f'//table[@id="investments-table-object"]//tbody//tr[{i}]/td[5]')
            type_list.append(types.text)
            cio_rating = self.browser.find_element(f'//table[@id="investments-table-object"]//tbody//tr[{i}]/td[6]')
            cio_rating_list.append(cio_rating.text)
            no_of_projects = self.browser.find_element(f'//table[@id="investments-table-object"]//tbody//tr[{i}]/td[7]')
            no_of_projects_list.append(no_of_projects.text)

            try:
                url = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]').find_element_by_tag_name(
                    "a").get_attribute("href")
                self.pdf_download(url, uii_list[-1])
                match = self.data_matching(uii_list[-1], investment_title_list[-1])
                if match:
                    pdf_matched.append("Matched")
                # self.urls.append({'URL': url, 'Investment Title': investment_title_list, 'UII': uii_list})
            except:
                pdf_matched.append("Not Matched")

        table_data = {"UII": uii_list, "Bureau": bureau_list, "Investment Title": investment_title_list,
                      'Total FYI 2021 Spendings': spending_list, "Type": type_list, 'CIO Rating': cio_rating_list,
                      '# of Projects': no_of_projects_list, "pdf matched": pdf_matched
                      }

        self.files.open_workbook('output/data.xlsx')
        self.files.create_worksheet('Individual Investment')
        self.files.append_rows_to_worksheet(content=table_data, header=True)
        self.files.save_workbook()

    def pdf_download(self, url, uii):
        """
        This function perform PDF downloading Operation in output folder
        :return: None

        """
        self.browser.execute_javascript(f"window.open('{url}');")
        tabs = self.browser.get_window_handles()
        self.browser.switch_window(tabs[-1])
        self.browser.wait_until_page_contains_element('//a[text()="Download Business Case PDF"]')
        self.browser.find_element('//a[text()="Download Business Case PDF"]').click()
        while True:
            if os.path.exists(f"{OUTPUT}/{uii}.pdf"):
                break
            else:
                pass
        self.browser.close_window()
        self.browser.switch_window(tabs[0])

    def data_matching(self, uii, investment):
        """
        This function performs the bonus task, open PDF and start matching the UII Code
        and Investment Title.

        :return: None
        """
        pdf_file = f'output/{uii}.pdf'
        pdf_name = self.pdf.get_text_from_pdf(pdf_file, pages=1)[1].split("Name of this Investment:")[1].split(".")[0]
        pdf_uii_code = self.pdf.get_text_from_pdf(pdf_file, pages=1)[1].split("Unique Investment Identifier (UII)")[1].split("Section")[0]

        if uii in pdf_uii_code and investment in pdf_name:
            return True
        return False

    def close_browser(self):
        """
        This Function close the browser when task is completed
        :return: None

        """
        self.browser.close_browser()
