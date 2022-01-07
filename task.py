from process import DashBoard
from config import AGENCY_NAME

if __name__ == "__main__":
    it_dashboard = DashBoard()
    it_dashboard.dive_in()
    it_dashboard.scrap_agencies()
    it_dashboard.scrap_agency_table(AGENCY_NAME)
