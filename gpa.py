import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd


def login(driver, username, password):
    driver.get("http://exams.mictech.ac.in/Login.aspx")

    # Locate the username and password fields
    username_input = driver.find_element(By.ID, "ContentPlaceHolder1_txtUsername")
    password_input = driver.find_element(By.ID, "ContentPlaceHolder1_txtPassword")

    # Enter username and password
    username_input.send_keys(username)
    password_input.send_keys(password)

    # Click login button
    login_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnLogin")
    login_button.click()

    # Wait for the dashboard page to load after login
    WebDriverWait(driver, 10).until(
        EC.url_to_be("http://exams.mictech.ac.in/StdDashBoard.aspx")
    )


def get_results(driver, hall_ticket_number):
    # Navigate to the dashboard page after login
    driver.get("http://exams.mictech.ac.in/StdDashBoard.aspx")

    # Wait for the SGPA/CGPA table to load
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "cBody_gvSGPA_CGPA"))
    )

    # Parse the page source
    soup = BeautifulSoup(driver.page_source, "html.parser")

    # Find the table with SGPA and CGPA
    table = soup.find(id="cBody_gvSGPA_CGPA")
    rows = table.find_all("tr")

    # Initialize variables to store GPA and CGPA
    semester_gpa = None
    semester_cgpa = None

    # Loop through rows to find 6th semester data
    for row in rows:
        cells = row.find_all("td")
        if len(cells) > 0 and cells[0].text.strip() == "VI":  # Check for 6th semester
            semester_gpa = cells[1].text.strip()
            semester_cgpa = cells[2].text.strip()
            break

    return {
        "Hall Ticket Number": hall_ticket_number,
        "6th Semester GPA": semester_gpa,
        "6th Semester CGPA": semester_cgpa,
    }


def main():
    results = []
    start_number = 6101
    end_number = 6166

    # Initialize WebDriver
    driver = webdriver.Chrome()

    for i in range(start_number, end_number + 1):
        hall_ticket_number = f"21H71A{i:04d}"

        # Skip 21H71A6107
        if (
            hall_ticket_number == "21H71A6109"
            or hall_ticket_number == "21H71A6135"
            or hall_ticket_number == "21H71A6160"
        ):
            print(f"Skipping hall ticket number: {hall_ticket_number}")
            continue

        print(f"Fetching results for: {hall_ticket_number}")

        # Login with hall ticket number as username and password
        login(driver, hall_ticket_number, hall_ticket_number)

        # Get results including GPA and CGPA for 6th semester
        result = get_results(driver, hall_ticket_number)
        results.append(result)

    # Close the WebDriver
    driver.quit()

    # Convert results to DataFrame
    results_df = pd.DataFrame(results)

    # Save results to Excel
    results_df.to_excel("student_results.xlsx", index=False)

    print("Results saved to student_results.xlsx")


if __name__ == "__main__":
    main()
