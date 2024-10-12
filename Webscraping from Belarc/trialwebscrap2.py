import os
from tkinter import Tk, Button, Label, filedialog
from openpyxl import Workbook
from bs4 import BeautifulSoup
import subprocess

def parse_filename(filename):
    # Split the filename based on underscores
    parts = filename.split('_')

    if len(parts) < 6:
        return None

    system_name = parts[0]  # System name
    department = parts[-5]  # Department name
    employee_name = parts[-4]  # Employee name
    branch_name = parts[-3]  # Branch name
    location = parts[-2]  # Location of PC
    port_number = parts[-1].split('.')[0]  # Port number or additional identifier (remove '.html')

    return {
        'System Name': system_name,
        'Department': department,
        'Employee Name': employee_name,
        'Branch Name': branch_name,
        'Location': location,
        'Port Number': port_number
    }

def extract_system_info_from_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

    # Extracting required data from HTML
    computer_name = soup.find('table', class_="reportHeader").find_all('tr')[1].find('td').text.strip()
    os = soup.find_all('div', class_='reportSection rsLeft')[0].find('td').get_text(separator='\n', strip=True).split('\n')[0]
    system_model = soup.find_all('div', class_="reportSection rsRight")[0].find('td').get_text(separator='\n', strip=True).split('\n')[0]
    processor = soup.find_all('div', class_="reportSection rsLeft")[1].find('td').get_text(separator='\n', strip=True).split('\n')[0]
    board = ''.join(soup.find_all('div', class_="reportSection rsRight")[1].find('td').get_text(separator='\n', strip=True).split('\n')[0].split(' ')[1:])
    hard_disk = soup.find_all('div', class_='reportSection rsLeft')[2].find('td').get_text(separator='\n', strip=True).split('\n')[0].split(' ')[0] + 'GB'
    memory = str(int(soup.find_all('div', class_="reportSection rsRight")[2].find('td').get_text(separator='\n', strip=True).split('\n')[0].split(' ')[0]) / 1000) + 'GB'
    ram_slots = soup.find_all('div', class_="reportSection rsRight")[2].find('td').get_text(separator='\n', strip=True).count('Slot')
    graphics = soup.find_all('div', class_="reportSection rsRight")[4].find('td').get_text(separator='\n', strip=True).split('[')[0]
    monitor = soup.find_all('div', class_="reportSection rsRight")[4].find('td').contents[-1].strip().split('[')[0]
    return {
        'Computer Name': computer_name,
        'OS': os,
        'System Model': system_model,
        'Processor': processor,
        'Board': board,
        'Hard Disk': hard_disk,
        'Memory': memory,
        'RAM Slots': ram_slots,
        'Graphic Card': graphics,
        'Monitor': monitor
    }

def process_folder(folder_path, output_file, status_label):
    # List all files in the directory
    files = os.listdir(folder_path)

    # Create a new Workbook
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'System Information'

    # Write headers
    headers = ['System Name', 'Department', 'Employee Name', 'Branch Name', 'Location', 'Port Number',
               'Computer Name', 'Operating System', 'System Model', 'Processor', 'Board', 'Hard Disk', 'Memory', 'RAM Slots','Graphic Card','Monitor']
    sheet.append(headers)

    # Process each file
    for filename in files:
        if filename.endswith('.html'):
            file_path = os.path.join(folder_path, filename)
            file_info = parse_filename(filename)

            if file_info:
                system_info = extract_system_info_from_html(file_path)
                row = [file_info['System Name'], file_info['Department'], file_info['Employee Name'],
                       file_info['Branch Name'], file_info['Location'], file_info['Port Number'],
                       system_info['Computer Name'], system_info['OS'], system_info['System Model'],
                       system_info['Processor'], system_info['Board'], system_info['Hard Disk'],
                       system_info['Memory'], system_info['RAM Slots'], system_info['Graphic Card'], system_info['Monitor']]
                sheet.append(row)

    # Save the workbook
    wb.save(output_file)
    status_label.config(text=f"Excel file saved successfully: {output_file}")

def select_folder_and_process():
    # Create Tkinter window
    root = Tk()
    root.title("HTML to Excel Converter")

    # Label to instruct user
    label = Label(root, text="Select folder containing HTML files:")
    label.pack(pady=10)

    # Function to handle button click
    def select_folder():
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            output_file = os.path.join(folder_selected, 'output.xlsx')
            status_label.config(text="Processing...")
            process_folder(folder_selected, output_file, status_label)
            subprocess.Popen(['start', '', output_file], shell=True)
            status_label.config(text=f"Excel file saved successfully: {output_file}")

    # Button to select folder
    button = Button(root, text="Select Folder", command=select_folder)
    button.pack(pady=10)

    # Label to show status
    status_label = Label(root, text="")
    status_label.pack(pady=10)

    # Run Tkinter main loop
    root.mainloop()

if __name__ == "__main__":
    select_folder_and_process()
