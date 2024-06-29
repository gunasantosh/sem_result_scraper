import pandas as pd
import ast  # For literal evaluation of string to dictionary


# Function to read text file and convert each row to a dictionary
def read_records_from_file(filename):
    records = []
    with open(filename, "r") as file:
        for line in file:
            record_dict = ast.literal_eval(line.strip())  # Convert string to dictionary
            records.append(record_dict)
    return records


# Function to convert records to DataFrame
def records_to_dataframe(records):
    data = []
    subject_names = {}

    # Extract subject names and initialize data dictionary
    for record in records:
        for subject in record["Subjects"]:
            subject_code = subject["Sub.Code"]
            subject_name = subject["Sub Name"]
            if subject_code not in subject_names:
                subject_names[subject_code] = subject_name

    for record in records:
        hall_ticket_number = record["Hall Ticket Number"]
        student_name = record["Student Name"]

        # Initialize a dictionary for the student
        student_data = {
            "Hall Ticket Number": hall_ticket_number,
            "Student Name": student_name,
        }

        # Initialize subject columns with empty values
        for subject_code, subject_name in subject_names.items():
            student_data[subject_name] = ""

        # Extract subjects and MRK/GR for each subject
        for subject in record["Subjects"]:
            subject_code = subject["Sub.Code"]
            subject_name = subject["Sub Name"]
            mrk_gr = subject["MRK/GR"]

            # Mark failed subjects with an indicator (e.g., 'Fail')
            if mrk_gr == "":
                mrk_gr = "Fail"

            # Add MRK/GR for each subject to the student's data
            student_data[subject_name] = mrk_gr

        # Append student's data to the main list
        data.append(student_data)

    # Convert list of dictionaries to DataFrame
    df = pd.DataFrame(data)

    # Reorder columns to desired format (Hall Ticket Number, Student Name, Sub1, Sub2, ..., Result)
    columns_order = ["Hall Ticket Number", "Student Name"] + list(
        subject_names.values()
    )
    df = df[columns_order]

    return df


# Main function
def main():
    filename = "student_records.txt"  # Replace with your file name
    records = read_records_from_file(filename)
    df = records_to_dataframe(records)

    # Export to Excel
    df.to_excel("student_records.xlsx", index=False)


if __name__ == "__main__":
    main()
