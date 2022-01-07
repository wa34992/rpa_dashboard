import time
from datetime import timedelta
import os
from config import OUTPUT, create_directory, URL
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from config import AGENCY_NAME


class DashBoard:

    def __init__(self):
        create_directory("output")
        self.files = Files()
        self.pdf = PDF()
        self.browser = Selenium()
        self.urls = []
        self.agencies_data = {}
        self.agency_table = {}

    def dive_in(self):
        """
        Navigate to itdashboard.org and click on Dive in
        :return:
        """
        self.browser.set_download_directory(OUTPUT)
        self.browser.open_available_browser(URL, maximized=True)
        self.browser.wait_until_page_contains("Dive in")
        self.browser.find_element('//a[@aria-controls="home-dive-in"]').click()
        self.browser.wait_until_page_contains_element('//div[@id="agency-tiles-container"]')

    def scrap_agencies(self):
        """
        read names and amount of all the agencies that appear on the page after clicking on Dive in.
        :return:
        """
        agencies = self.browser.find_elements('//div[@id="agency-tiles-container"]//span[@class="h4 w200"]')
        amounts = self.browser.find_elements('//div[@id="agency-tiles-container"]//span[@class="h1 w900"]')
        while True:
            if len(amounts) == 0 or len(agencies) == 0:
                time.sleep(3)
                agencies = self.browser.find_elements('//div[@id="agency-tiles-container"]//span[@class="h4 w200"]')
                amounts = self.browser.find_elements('//div[@id="agency-tiles-container"]//span[@class=" h1 w900"]')
            else:
                break
        agency_names = []
        agency_amounts = []
        for agency, amount in zip(agencies, amounts):
            agency_names.append(agency.text)
            agency_amounts.append(amount.text)

        self.agencies_data["Agency"] = agency_names
        self.agencies_data["Amount"] = agency_amounts

    def scrap_agency_table(self, agency):
        """
        scrap
        :return:
        """
        pdf_links = []
        pdf_index = []
        self.browser.find_element(f"//span[text()='{agency}']").click()
        self.browser.wait_until_page_contains_element("//table[@id='investments-table-object']",
                                                      timeout=timedelta(seconds=10))
        table_headers = self.get_headers()
        for header in table_headers:
            self.agency_table[header] = []
        self.agency_table["PDF"] = []
        self.browser.select_from_list_by_value("//select[@name='investments-table-object_length']", "-1")
        total_rows = self.browser.find_element("//div[@class='dataTables_info']").text.split(" ")[-2]
        self.browser.wait_until_page_contains_element(f"//table[@id='investments-table-object']//tr[{total_rows}]",
                                                      timeout=timedelta(seconds=20))
        rows = self.browser.find_elements(f"//table[@id='investments-table-object']//tbody//tr")
        row_list = []
        for count, row in enumerate(rows):
            row_cells = row.find_elements_by_tag_name("td")
            raw_data = []
            for cell, header in zip(row_cells, table_headers):
                try:
                    link = cell.find_element_by_tag_name("a").get_attribute("href")
                    pdf_links.append(link)
                    pdf_index.append(count)
                except:
                    pass
                self.agency_table[header].append(cell.text)
            row_list.append(raw_data)
        self.download_pdfs(pdf_links, pdf_index)
        self.match_pdfs(pdf_index)
        self.write_data_to_workbook()

    def download_pdfs(self, pdf_links, pdf_index):
        """
        download pdfs where link is found in table
        :param pdf_links: list of pdf links to download
        :param pdf_index: index of pdf row in table
        :return:
        """
        for link, index in zip(pdf_links, pdf_index):
            self.browser.execute_javascript(f"window.open('{link}');")
            tabs = self.browser.get_window_handles()
            self.browser.switch_window(tabs[-1])
            self.browser.wait_until_page_contains_element('//a[text()="Download Business Case PDF"]',
                                                          timeout=timedelta(seconds=10))
            self.browser.find_element('//a[text()="Download Business Case PDF"]').click()
            while True:
                if os.path.exists(f"{OUTPUT}/{self.agency_table['UII'][index]}.pdf"):
                    break
                else:
                    pass

    def match_pdfs(self, pdf_index):
        """
        This function performs the bonus task, open PDF and start matching the UII Code
        and Investment Title.

        """
        for index in pdf_index:
            unique_id = f"{self.agency_table['UII'][index]}"
            name = f"{self.agency_table['Investment Title'][index]}"
            pdf_file = f"{OUTPUT}/{self.agency_table['UII'][index]}.pdf"
            pdf_name = self.pdf.get_text_from_pdf(pdf_file, pages=1)[1].split(
                "Name of this Investment:")[1].split(".")[0]
            pdf_uii_code = self.pdf.get_text_from_pdf(pdf_file, pages=1)[1].split(
                "Unique Investment Identifier (UII)")[1].split("Section")[0]
            if unique_id in pdf_uii_code and name in pdf_name:
                self.agency_table["PDF"].append("matched")
            else:
                self.agency_table["PDF"].append("not matched")

    def get_headers(self):
        headers = []
        raw_headers = self.browser.find_elements(f"//div[@id='investments-table-object_wrapper']//table/thead/tr[2]/th")
        for head in raw_headers:
            headers.append(head.text)
        return headers

    def write_data_to_workbook(self):
        self.files.create_workbook(f"{OUTPUT}/agencies.xlsx")
        self.files.remove_worksheet("Sheet")
        self.files.create_worksheet("AGENCIES")
        self.files.append_rows_to_worksheet(self.agencies_data, header=True)
        self.files.create_worksheet("UII TABLE")
        self.files.append_rows_to_worksheet(self.agency_table, header=True)
        self.files.save_workbook()
        self.files.close_workbook()
