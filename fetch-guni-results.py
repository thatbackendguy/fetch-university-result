# Before running the script, run the line below in terminal
# pip install selenium webdriver_manager pandas openpyxl

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import pandas as pd


bdaEnrolls = ["add list of BDA enrollment numbers here"]
cbaEnrolls = ["add list of CBA enrollment numbers here"]
csEnrolls = ["add list of CS enrollment numbers here"]

enrollmentList = []
filename = ""
# input from user
branch = int(
    input("Select your branch:\n1. BDA\n2. CBA\n3. CS\nEnter your input: "))
if (branch == 1):  # BDA
    branch = 127
    enrollmentList = bdaEnrolls
    filename = "BDA"
elif (branch == 2):  # CBA
    branch = 125
    enrollmentList = cbaEnrolls
    filename = "CBA"
elif (branch == 3):  # CS
    branch = 207
    enrollmentList = csEnrolls
    filename = "CS"
else:
    print("Invalid input!")
    exit(1)

sem = int(input("Enter your semester: "))
examID = int(input("Enter your exam id: "))

df = pd.DataFrame(columns=['Name', 'Enrollment', 'SGPA', 'CGPA'])

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

url = "https://result.ganpatuniversity.ac.in/"
driver.get(url)
driver.implicitly_wait(10)


# fetching
for i in range(len(enrollmentList)):
    institute = Select(driver.find_element(By.ID, 'ddlInst'))
    driver.implicitly_wait(2)
    institute.select_by_value('17')
    driver.implicitly_wait(2)

    degree = Select(driver.find_element(By.ID, 'ddlDegree'))
    driver.implicitly_wait(2)
    degree.select_by_value(str(branch))
    driver.implicitly_wait(2)

    semester = Select(driver.find_element(By.ID, 'ddlSem'))
    driver.implicitly_wait(2)
    semester.select_by_value(str(sem))
    driver.implicitly_wait(2)

    exam = Select(driver.find_element(By.ID, 'ddlScheduleExam'))
    driver.implicitly_wait(2)
    exam.select_by_value(str(examID))
    driver.implicitly_wait(2)

    enrollment = driver.find_element(By.ID, "txtEnrNo")
    enrollment.send_keys(enrollmentList[i])
    driver.implicitly_wait(2)

    search = driver.find_element(By.ID, "btnSearch")
    search.click()
    driver.implicitly_wait(2)

    sgpa = driver.find_element(By.ID, "uclGrd_lblSGPA").text
    cgpa = driver.find_element(By.ID, "uclGrd_lblPrgCGPA").text
    name = driver.find_element(By.ID, "uclGrd_lblStudentName").text

    data = [name, enrollmentList[i], sgpa, cgpa]
    df.loc[len(df)] = data
    backBTN = driver.find_element(By.ID, "btnBack")
    backBTN.click()

driver.close()

df.sort_values(by=['SGPA'])
df.to_excel(f"{filename}-results.xlsx", index=False)
print(f"==> {filename}-results.csv exported successfully!!")
