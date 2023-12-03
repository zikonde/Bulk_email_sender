import requests
from bs4 import BeautifulSoup
import openpyxl

def scrape_secondary_url(secondary_url):
    # Make a request to the secondary URL
    try:
        response = requests.get(secondary_url)
    except Exception as e:
        print(f"Error: {e}")
        return
    else:
        if response.status_code != 200:
            print(f"Failed to fetch the Secondary URL. Status code: {response.status_code}")
            return

        # Parse HTML content
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract teacher's name
        teacher_name = ''
        for email in soup.select('.userBaseName'):
            teacher_name = email.text.strip()

        # Extract teacher's email
        teacher_email = ''
        for email in soup.select('.email'):
            teacher_email = email.text.strip().split("\n")[1]

        # Extract teacher's field
        teacher_field = ''
        for field in soup.select('.second_research'):
            teacher_field = field.text.strip()
        
        return secondary_url, teacher_name, teacher_email, teacher_field

def scrape_teachers_and_secondary_url(primary_url):
    # Make a request to the primary URL
    response = requests.get(primary_url)
    if response.status_code != 200:
        print(f"Failed to fetch the primary URL. Status code: {response.status_code}")
        return

    # Parse HTML content
    soup = BeautifulSoup(response.text, 'html.parser')

    # Extract teacher names and their corresponding secondary URLs
    teacher_info = []
    for teacher_entry in soup.select('.xl75 a'):
        secondary_url = teacher_entry['href']
        all_info = scrape_secondary_url(secondary_url)
        if all_info:
            teacher_info.append(all_info)
        else:
            teacher_info.append((secondary_url,'','',''))
    

    # Fetch emails from secondary URLs
    '''for index, (teacher_name, secondary_url) in enumerate(teacher_info):
        secondary_response = requests.get(secondary_url)
        if secondary_response.status_code != 200:
            print(f"Failed to fetch the secondary URL for {teacher_name}. Status code: {secondary_response.status_code}")
            continue

        secondary_soup = BeautifulSoup(secondary_response.text, 'html.parser')
        email_element = secondary_soup.find('a', {'href': 'mailto:'})
        teacher_email = email_element.text if email_element else "N/A"
        teacher_info[index] = (teacher_name, teacher_email)'''

    return teacher_info

def write_to_excel(data, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Teachers Information"

    # Write headers
    sheet['A1'] = "网页"
    sheet['B1'] = "姓名"
    sheet['C1'] = "Email"
    sheet['D1'] = "Interest"
    sheet['E1'] = "Field"

    print(data)
    # Write data
    for row_index, (teacher_page, teacher_name, teacher_email, teacher_field) in enumerate(data, start=2):
        sheet[f'A{row_index}'] = teacher_page
        sheet[f'B{row_index}'] = teacher_name
        sheet[f'C{row_index}'] = teacher_email
        sheet[f'E{row_index}'] = teacher_field

    # Save to Excel file
    workbook.save(output_file)
    print(f"Data written to {output_file}")

if __name__ == "__main__":
    primary_url = "http://www.cst.zju.edu.cn/szdw/list.htm"
    output_excel_file = "teachers_information.xlsx"

    teacher_data = scrape_teachers_and_secondary_url(primary_url)
    if teacher_data:
        write_to_excel(teacher_data, output_excel_file)