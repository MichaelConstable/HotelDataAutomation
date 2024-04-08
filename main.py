import os
import re
import datetime
import shutil
from pypdf import PdfReader


def delete_folder(folder_path):
    """Delete a folder and its contents if it exists."""
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
        print(f"Folder '{folder_path}' and its contents have been deleted.")
    else:
        print(f"Folder '{folder_path}' does not exist.")


def extract_trial_balance_data():
    # Ask the user to input the file location
    file_location = input("Enter the file location for Trial Balance PDF: ")
    # Extract text from PDF
    text = extract_text_from_pdf(file_location)
    # Split the text into lines
    lines = text.splitlines()
    # Get output file path
    output_file_path = get_output_file_path(file_location, "Trial_Balance")
    with open(output_file_path, "w") as output_file:
        # Loop through each line
        for line in lines:
            # Split the line by multiple spaces to separate the columns
            line_without_whitespace = re.sub(r'\s+', '', line)
            # if first character in line is a digit we know it's a code to identify the money amount
            if line_without_whitespace[0].isdigit():
                # then we can remove the white space
                line_with_spaces = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', line_without_whitespace)
                # and then remove the digit code because we know in the middle column we get the text
                # representative of the code
                data = line_with_spaces[4:]
                pattern = re.compile(r'([a-zA-Z*/]+)([-+]?[\d,]+(?:\.\d+)?)')
                # Match the pattern in each line of data
                matches = pattern.findall(data)
                # Write the extracted text and number parts to the file
                for match in matches:
                    text_part = match[0]
                    number_part = match[1].replace(',', '')  # Remove commas from the number part
                    output_file.write(f"{text_part}\n")
                    output_file.write(f"{number_part}\n")
    print(f"Data has been extracted and saved to: {output_file_path}")


def extract_manager_flash_data():
    # Ask the user to input the file location
    file_location = input("Enter the file location for Manager Flash PDF: ")
    # Extract text from PDF
    text = extract_text_from_pdf(file_location)
    # Split the text into lines
    lines = text.splitlines()
    # Get output file path
    output_file_path = get_output_file_path(file_location, "Manager_Flash")
    with open(output_file_path, "w") as output_file:
        # Loop through each line
        for line_number, line in enumerate(lines, start=1):
            if 15 <= line_number <= 42:
                # clean lines to get format string and joined number, strip all after 38 characters
                index = next((i for i, char in enumerate(line) if char.isdigit()), None)
                cleaned_line = line[:index].strip() + line[index:]
                cleaned_line = cleaned_line[:43]
                modified_line = ''  # Initialize modified line
                prev_char_is_digit = False  # Flag to track if previous character was a digit
                # Iterate through each character in the line
                for char in cleaned_line:
                    # Check if the character is a digit
                    if char.isdigit():
                        # If previous character was not a digit, insert a newline character
                        if not prev_char_is_digit:
                            modified_line += "\n"
                        modified_line += char  # Add the digit to the modified line
                        prev_char_is_digit = True
                    else:
                        modified_line += char  # Add non-digit characters to the modified line
                        prev_char_is_digit = False
                # Write the modified line to the output file
                output_file.write(modified_line.strip() + "\n")
    # Print the path where the extracted data is saved
    print(f"Lines 15 to 25 have been extracted and saved to: {output_file_path}")


def extract_tax_exempt_data():
    # Ask the user to input the file location
    file_location = input("Enter the file location for Tax Exempt PDF: ")
    # Extract text from PDF
    text = extract_text_from_pdf(file_location)
    # Split the text into lines
    lines = text.splitlines()

    # Get output file path
    output_file_path = get_output_file_path(file_location, "Tax_Exempt")
    with open(output_file_path, "w") as output_file:
        for line_number, line in enumerate(lines, start=1):
            # Skip lines until reaching the 7th line
            if line_number < 8:
                continue

            # Strip leading whitespaces from the line
            line = line.lstrip()

            # Check if the line contains "Tax Type Total"
            if "Tax Type Total" in line:
                # If it does, break out of the loop
                break

            # Check if the line starts with "T-" or "Routing Instruction"
            if not (line.startswith("T-") or line.startswith("Routing Instruction")):
                # If it doesn't start with "T-" or "Routing Instruction", extract and write the data to the file
                extracted_GuestName = line[0:40]
                extracted_RoomNum = line[46:49]
                extracted_Price = line[133:139]
                output_file.write(f"{extracted_GuestName}\n")
                output_file.write(f"{extracted_RoomNum}\n")
                output_file.write(f"{extracted_Price}\n")

            # Check if the line starts with "Routing Instruction"
            if line.startswith("Routing Instruction"):
                line = line[37:]
                line = line.split(":", 1)[0]  # Split at the first occurrence of ":" and take the first part
                output_file.write(f"{line}\n")

    print(f"Extracted data has been saved to: {output_file_path}")


# Helper functions to modularize task
def get_output_file_path(file_location, file_type):
    """Generate the output file path based on file location and type."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    folder_path = os.path.join(desktop_path, "Audit Data Processing")
    create_folder(folder_path)  # Ensure the folder exists
    file_name = os.path.basename(file_location)
    file_name_without_extension = os.path.splitext(file_name)[0]
    return os.path.join(folder_path, f"{file_name_without_extension}_{file_type}.txt")


def create_folder(folder_path):
    """Create a folder if it doesn't exist."""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Folder '{folder_path}' has been created.")
    else:
        print(f"Folder '{folder_path}' already exists.")


def extract_text_from_pdf(file_location):
    """Extract text from a PDF file."""
    if os.path.exists(file_location):
        with open(file_location, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            page = pdf_reader.pages[0]
            return page.extract_text(extraction_mode="layout", layout_mode_space_vertically=False)
    else:
        print(f"File '{file_location}' does not exist.")
        return ""


def main():
    # Delete the folder "Audit Data Processing" and its contents if it exists
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    folder_to_delete = os.path.join(desktop_path, "Audit Data Processing")

    delete_folder(folder_to_delete)
    extract_trial_balance_data()
    extract_manager_flash_data()
    extract_tax_exempt_data()


if __name__ == "__main__":
    main()
