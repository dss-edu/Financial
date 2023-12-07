# from finance.connect import connect
from finance.backend import update_fy
from config.settings import SCHOOLS

if __name__ == "__main__":
    year = ""
    for school in SCHOOLS.keys():
        update_fy(school, year)



