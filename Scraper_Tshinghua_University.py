import requests
from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup
import openpyxl
import re

def scrape_secondary_url(secondary_url):
    # Make a request to the secondary URL
    secondary_url = "https://www.thss.tsinghua.edu.cn/"+secondary_url.replace('../','')
    
    try:
        uClient = uReq(secondary_url)
        page_html = uClient.read().decode('utf-8')
        uClient.close()
        response = requests.get(secondary_url)
    except Exception as e:
        print(f"Error: {e}")
        return secondary_url, '', '', ''
    else:
        if response.status_code != 200:
            print(f"Failed to fetch the Secondary URL. Status code: {response.status_code}")
            return secondary_url, '', '', ''

        # Parse HTML content
        soup = BeautifulSoup(page_html, 'html.parser')
        

        # Extract teacher's name
        # Find the span element with the text "姓名"
        name_element = soup.find('span', string="姓名:")
        if not name_element:
            name_element = soup.find('span', string="姓名：")
            if not name_element:
                name_element = soup.find('span', string="姓   名:")

        teacher_name = ''
        # Extract the paragraph after the spam element
        if name_element:
            next_paragraph = name_element.find_next('span')
            if next_paragraph:
                teacher_name = next_paragraph.get_text(strip=True)

        # Extract teacher's email
        teacher_email = ""
        email_element = soup.select('a[href^="mailto:"]')
        if(email_element):
            email_element = email_element[0]
            teacher_email = email_element.text if email_element else ""
        else:
            email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
            teacher_email = email_pattern.findall(page_html)
            if(len(teacher_email) > 0):
                teacher_email = teacher_email[0]
            else:
                teacher_email = ""

        # Find the span element with the text "研究领域"
        field_element = soup.find('span', string="研究领域:")
        if not field_element:
            field_element = soup.find('span', string="研究领域：")

        teacher_field = ''
        # Extract the paragraph after the spam element
        if field_element:
            next_paragraph = field_element.find_next('span')
            if next_paragraph:
                teacher_field = next_paragraph.get_text(strip=True)
        
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
    for teacher_entry in soup.select('.name-container a'):
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
    sheet['C1'] = "姓"
    sheet['D1'] = "Email"
    sheet['E1'] = "Interest"
    sheet['F1'] = "Field"

    print(data)
    # Write data
    for row_index, (teacher_page, teacher_name, teacher_email, teacher_field) in enumerate(data, start=2):
        sheet[f'A{row_index}'] = teacher_page
        sheet[f'B{row_index}'] = teacher_name
        sheet[f'C{row_index}'] = f"=LEFT($B{row_index},1)"
        sheet[f'D{row_index}'] = teacher_email
        sheet[f'F{row_index}'] = teacher_field

    # Save to Excel file
    workbook.save(output_file)
    print(f"Data written to {output_file}")

if __name__ == "__main__":
    primary_url = "https://www.thss.tsinghua.edu.cn/szdw/jsml.htm"
    output_excel_file = "teachers_information.xlsx"

    teacher_data = scrape_teachers_and_secondary_url(primary_url)
    if teacher_data:
        write_to_excel(teacher_data, output_excel_file)