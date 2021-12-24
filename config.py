import os

OUTPUT = F"{os.getcwd()}/output"
URL = "https://itdashboard.gov/"
AGENCY_NAME = "Social Security Administration"


def create_directory(directory_name):
    if not os.path.exists(f"{os.getcwd()}/{directory_name}"):
        os.mkdir(f"{os.getcwd()}/{directory_name}")
