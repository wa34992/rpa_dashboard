from process import DashBoard
from config import AGENCY_NAME

if __name__ == "__main__":
    it_dashboard = DashBoard()
    it_dashboard.dive_in_agencies()
    it_dashboard.agency_table_scrape(AGENCY_NAME)
    it_dashboard.close_browser()
