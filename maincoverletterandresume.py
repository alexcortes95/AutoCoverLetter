from docxtpl import DocxTemplate
import datetime
import os

# This function writes to the doc file and then converts it into a pdf for the cover letter
def converttopdfcl():
    # Choose the path where template is located
    doccl = DocxTemplate(r"C:\Users\alexc\Documents\Cover_Letters\coverlettertemplate.docx")

    doccl.render(contextcl)
    # Choose the name and path of the output file you are creating
    doccl.save(f"C:/Users/alexc/Documents/Cover_Letters/Cover_Letter_" + company_name + "_" + position_name + ".docx")

    from docx2pdf import convert

    # Specify the path to your .docx file
    docx_file = "Cover_Letter_" + company_name + "_" + position_name + ".docx"
    print(docx_file)
    convert(file_pathcl)

# This function writes to the doc file and then converts it into a pdf for the resume
def converttopdfr():
    # Choose the path where template is located
    docr = DocxTemplate(r"C:\Users\alexc\Documents\Custom_Resumes\DataAnalystResumeTemplate.docx")

    docr.render(contextr)
    # Choose the name and path of the output file you are creating
    docr.save(f"C:/Users/alexc/Documents/Custom_Resumes/Resume_" + company_name + "_" + position_name + ".docx")

    from docx2pdf import convert

    # Specify the path to your .docx file
    docx_file = "Resume" + company_name + "_" + position_name + ".docx"
    print(docx_file)
    convert(file_pathr)

# Fields for cover letter template
company_name = input("Enter the name of the company ")
position_name = input("Enter the name of the position ")

# Fields for resume template
needresume = input("Do you need a custom resume? (y/n) ")
if needresume == "y":
    need_GC_point1 = input("Need GC Point? (y/n) ")
    if need_GC_point1 == "y":
        GC_point1 = input("Add point here ")
    else:
        GC_point1 = "Collaborated with senior management to achieve key financial targets—including EBITDA, sales, and margin objectives—resulting in an 8% year-over-year increase in sales through data analysis and cleanliness."
    need_FCI_point1 = input("Need FCI Point? (y/n) ")
    if need_FCI_point1 == "y":
        FCI_point1 = input("Add point here ")
    else:
        FCI_point1 = "Identified and implemented strategies for data quality improvements, ensuring customer data accuracy."
    need_Prog_1 = input("Need to add Programs? (y/n) ")
    if need_Prog_1 == "y":
        Prog_1 = ", " + input("Add Programs here ")
    else:
        Prog_1 = ""
    need_Skill_1 = input("Need to add Skills? (y/n) ")
    if need_Skill_1 == "y":
        Skill_1 = ", " + input("Add skills here ")
    else:
        Skill_1 = ""
    need_Courses_1 = input("Need to add a course? (y/n) ")
    if need_Courses_1 == "y":
        Courses_1 = "•  " + input("Add course here ")
        need_Courses_2 = input("Need to add course 2? (y/n) ")
        if need_Courses_2 == "y":
            Courses_2 = "•  " + input("Add course 2 here ")
            need_Courses_3 = input("Need to add course 3? (y/n) ")
            if need_Courses_3 == "y":
                Courses_3 = "•  " + input("Add course 3 here ")
                need_Courses_4 = input("Need to add course 4? (y/n) ")
                if need_Courses_4 == "y":
                    Courses_4 = "•  " + input("Add course 4 here ")
                else:
                    Courses_4 = ""
            else:
                Courses_3 = ""
                Courses_4 = ""
        else:
            Courses_2 = ""
            Courses_3 = ""
            Courses_4 = ""
    else:
        Courses_1 = ""
        Courses_2 = ""
        Courses_3 = ""
        Courses_4 = ""
else:
    GC_point1 = "Collaborated with senior management to achieve key financial targets—including EBITDA, sales, and margin objectives—resulting in an 8% year-over-year increase in sales through data analysis and cleanliness."
    FCI_point1 = "Identified and implemented strategies for data quality improvements, ensuring customer data accuracy."
    Prog_1 = ""
    Skill_1 = ""
    Courses_1 = ""
    Courses_2 = ""
    Courses_3 = ""
    Courses_4 = ""

# Associating fields to cover letter
today_date = datetime.datetime.today().strftime('%B %d, %Y')
contextcl = {
    'today_date': today_date,
    'company_name': company_name,
    'position_name': position_name
    #'add_line1': add_line1,
    #'add_line2': add_line2
}

# Associating fields to resume
contextr = {
    'GC_point1': GC_point1,
    'FCI_point1': FCI_point1,
    'Prog_1': Prog_1,
    'Skill_1': Skill_1,
    'Courses_1': Courses_1,
    'Courses_2': Courses_2,
    'Courses_3': Courses_3,
    'Courses_4': Courses_4
}

# Path and naming conventions for saving Cover Letter
file_namecl = "Cover_Letter_" + company_name + "_" + position_name + ".docx"
file_pathcl = os.path.join("C:/Users/alexc/Documents/Cover_Letters", file_namecl)

# Path and naming conventions for saving Resume
file_namer = "Resume_" + company_name + "_" + position_name + ".docx"
file_pathr = os.path.join("C:/Users/alexc/Documents/Custom_Resumes", file_namer)

#For Cover Letter
# If file already exists, you will be prompted to choose whether or not you would like to overwrite, else it will convert .docx to PDF and save
if os.path.exists(file_pathcl):
    overwrite_response = input(f"File exists, do you want to overwrite? yes or no ").lower()
    if overwrite_response == "yes":
        converttopdfcl()
    else:
        print("Change inputs and run again")
else:
    converttopdfcl()

# For resume, only if you choose to create a resume
if needresume == "y":
    converttopdfr()





