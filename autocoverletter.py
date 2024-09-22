from docxtpl import DocxTemplate
import datetime
import os
def converttopdf():
    # Choose the path where template is located
    doc = DocxTemplate(r"C:\Users\alexc\Documents\Cover_Letters\coverlettertemplate.docx")

    doc.render(context)
    # Choose the name and path of the output file you are creating
    doc.save(f"C:/Users/alexc/Documents/Cover_Letters/Cover_Letter_" + company_name + "_" + position_name + ".docx")

    from docx2pdf import convert

    # Specify the path to your .docx file
    docx_file = "Cover_Letter_" + company_name + "_" + position_name + ".docx"
    print(docx_file)
    convert(file_path)

# Inputs that correspond to the fields you would like to change in the template
company_name = input("Enter the name of the company ")
position_name = input("Enter the name of the position ")
add_line1 = input("Enter the first line of the address ")
add_line2 = input("Enter the second line of the address ")

today_date = datetime.datetime.today().strftime('%B %d, %Y')

context = {
    'today_date': today_date,
    'company_name': company_name,
    'position_name': position_name,
    'add_line1': add_line1,
    'add_line2': add_line2
}

file_name = "Cover_Letter_" + company_name + "_" + position_name + ".docx"
file_path = os.path.join("C:/Users/alexc/Documents/Cover_Letters", file_name)

# If file already exists, you will be prompted to choose whether or not you would like to overwrite, else it will convert .docx to PDF and save
if os.path.exists(file_path):
    overwrite_response = input(f"File exists, do you want to overwrite? yes or no ").lower()
    if overwrite_response == "yes":
        converttopdf()
    else:
        print("Change inputs and run again")
else:
    converttopdf()



