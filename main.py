import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time


def get_results(hall_ticket_number):
    # Setup WebDriver
    # options = webdriver.ChromeOptions()
    # options.add_argument("--headless")  # Run in headless mode (no UI)
    driver = webdriver.Chrome()

    try:
        driver.get("http://exams.mictech.ac.in/Result.aspx?Id=280")

        # Locate the hall ticket input field and submit button
        hall_ticket_input = driver.find_element(By.ID, "ContentPlaceHolder1_txtRegNo")
        get_result_button = driver.find_element(
            By.ID, "ContentPlaceHolder1_btnGetResult"
        )

        # Enter the hall ticket number and submit the form
        hall_ticket_input.send_keys(hall_ticket_number)
        get_result_button.click()

        # Wait for the results to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "dvContents"))
        )

        # Get the page source and parse it with BeautifulSoup
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # Extract data from the form
        result_data = {
            "Hall Ticket Number": hall_ticket_number,
            "Student Name": soup.find(id="ContentPlaceHolder1_txtStudentName").get(
                "value"
            ),
            "Course": soup.find(id="ContentPlaceHolder1_txtCourse").get("value"),
            "Branch Name": soup.find(id="ContentPlaceHolder1_txtGRP").get("value"),
        }

        # Extract data from the table
        table = soup.find(id="ContentPlaceHolder1_dgvStudentHistory")
        rows = table.find_all("tr")[1:]  # Skip the header row

        subjects = []
        for row in rows:
            cols = row.find_all("td")
            subject_data = {
                "Sub.Code": cols[0].find("span").text,
                "Sub Name": cols[1].find("span").text,
                "CR": cols[2].find("span").text,
                "MRK/GR": cols[3].find("span").text,
                "Result": cols[4].find("span").text,
            }
            subjects.append(subject_data)

        result_data["Subjects"] = subjects

    finally:
        # Close the WebDriver
        driver.quit()

    return result_data


def main():
    results = []
    start_number = 6101
    end_number = 6166

    for i in range(start_number, end_number + 1):
        hall_ticket_number = f"21H71A{i:04d}"
        result = get_results(hall_ticket_number)
        print(f"Fetching results for: {hall_ticket_number}")
        results.append(result)
        time.sleep(1)

    with open("Student_results.txt", "w") as f:
        for line in results:
            f.write(f"{line}\n")


if __name__ == "__main__":
    main()
