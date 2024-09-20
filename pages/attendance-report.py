import streamlit as st
import pandas as pd
import base64 
from pathlib import Path
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle,Paragraph,Image
import io
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
import textwrap
import zipfile
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE
from email import encoders

from datetime import date

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)


def generate_pdf(df, row,Branch_Choice,test_choice,submission_d,semester,no_of_subjects,note):  

    from datetime import datetime
    today = datetime.today()
    day = today.day
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    date_of_generation = today.strftime(f"%d{suffix} %b, %Y")
    date_of_generation = str(date_of_generation)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0, leftMargin=50, rightMargin=50, bottomMargin=0)
    
    styles = getSampleStyleSheet()

    # Creating a bold and capitalized Times New Roman style
    bold_times_style = styles["Heading1"]
    bold_times_style.fontName = "Times-Bold"
    bold_times_style.fontSize = 12 
    bold_times_style.alignment = 1
    bold_times_style.textTransform = "uppercase"
    bold_times_style.spaceAfter = 1
    bold_times_style.spaceBefore = 1

    bold_style = styles["Heading2"]
    bold_style.fontName = "Times-Bold"
    bold_style.fontSize = 10
    bold_style.spaceAfter = 1
    bold_style.spaceBefore = 1

    elements = [] 

    header_path = "Images/Header_RV.png"
    image = Image(header_path, width=8*inch, height=1.6445*inch)
    image.vAlign = "TOP"
    elements.append(image)

    heading = Paragraph('<u>'+Branch_Choice+'</u>', bold_times_style)
    elements.append(heading)
 
    heading = Paragraph('<u>'+test_choice+'</u>', bold_times_style)
    elements.append(heading)

    style_sheet = getSampleStyleSheet() #date
    style = style_sheet['Normal']
    text = date_of_generation
    para = Paragraph(text, style)
    elements.append(para)

    style_sheet = getSampleStyleSheet() #date
    style = style_sheet['Normal']
    text = "\u00a0"
    para = Paragraph(text, style)
    elements.append(para)

    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "To, "
    para = Paragraph(text, style)
    elements.append(para)
    
    father = str(df.iloc[row, 3])
    heading = Paragraph("\u00a0 \u00a0 \u00a0Mr/Mrs \u00a0"+father+",", bold_style)
    elements.append(heading)

    student_name = df.iloc[row,1]
    USN = df.iloc[row,2]
    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "\u00a0 \u00a0 \u00a0 \u00a0 \u00a0 \u00a0The Attendance report of your ward <b>"+str(student_name)+",\u00a0"+str(USN)+"</b> studying in <b>"+str(semester)+"</b> is given below : "
    para = Paragraph(text, style)
    elements.append(para)

    wrapped_sl = textwrap.fill("Sl. No", width=3)
    wrapped_attendance  = textwrap.fill("Attendance Percentage", width=10)
    wrapped_classheld  = textwrap.fill("Classes Held", width=7)
    wrapped_classattended = textwrap.fill("Classes Attended", width=9)
    wrapped_testmarks = textwrap.fill(str(df.iloc[1,10]), width=10)
    wrapped_assignment = textwrap.fill(str(df.iloc[1,11]), width=10)
    data = [[wrapped_sl,"Subject Name",wrapped_classheld,wrapped_classattended,wrapped_attendance]]

    for i in range(no_of_subjects):
        subject = df.iloc[0, 8 + i * 2]
        try:
            classesheld = int(df.iloc[row, 8 + i * 2])
        except ValueError:
            classesheld = 0
        try:
            classattended = int(df.iloc[row, 9 + i * 2])
        except ValueError:
            classattended = 0
    
        # Check if both classesheld and classattended are zero
        if classesheld == 0 and classattended == 0:
            classesheld = '-'
            classattended = '-'
            attendance = '-'
        else:
            try:
                attendance = int(classattended / classesheld * 100)
            except ZeroDivisionError:
                attendance = 0
    
      
        wrapped_subject = textwrap.fill(subject, width=30)
    
        data.append([str(i + 1), wrapped_subject, classesheld, classattended, "{}%".format(attendance)])

    table = Table(data, splitByRow=1, spaceBefore=10, spaceAfter=10, cornerRadii=[1.5,1.5,1.5,1.5])
    

    table.setStyle(TableStyle([      
    
    ('BACKGROUND', (0, 0), (-1, 0), '#FFFFFF'),
    ('TEXTCOLOR', (0, 0), (-1, 0), '#000000'),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('fontsize', (-1,-1), (-1,-1), 14),
    ('ALIGNMENT', (1, 1), (1, -1), 'LEFT'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
    ('BACKGROUND', (0, 1), (-1, -1), '#FFFFFF'),
    ('GRID', (0, 0), (-1, -1), 1, "black")
  ]))

    elements.append(table)

    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "<b>Remarks:</b>\u00a0"+str(df.iloc[row,7])+""
    para = Paragraph(text, style)
    elements.append(para)

    style = style_sheet['Normal']
    text = "\u00a0 "
    para = Paragraph(text, style)
    elements.append(para)

    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "<b>Note:</b>\u00a0"+str(note)+""
    para = Paragraph(text, style)
    elements.append(para)


    style = style_sheet['Normal']
    text = "\u00a0 "
    para = Paragraph(text, style)
    elements.append(para)


    counsellor_mail = str(df.iloc[row,6])
    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "Please sign and send the report to “"+counsellor_mail+"” on or before "+submission_d+"."
    para = Paragraph(text, style)
    elements.append(para)

    image_path = "Images/default.png"

 
    if Branch_Choice == "COMPUTER SCIENCE & ENGINEERING":
        image_path = "Images/CSE_Signature.png"
    elif Branch_Choice == "INFORMATION SCIENCE & ENGINEERING":
        image_path = "Images/ISE_Signature.png"
    elif Branch_Choice == "ELECTRONICS & COMMUNICATION ENGINEERING":
        image_path = "Images/ECE_Signature.png"
    elif Branch_Choice == "MECHANICAL ENGINEERING":
        image_path = "Images/ME_Signature.png"
    elif Branch_Choice == "MASTER OF COMPUTER APPLICATIONS":
        image_path = "Images/MCA_Signature.png"
    image = Image(image_path, width=7*inch, height=1.4155*inch)
    elements.append(image)

    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "\u00a0" 
    para = Paragraph(text, style)
    elements.append(para)

    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "\u00a0"
    para = Paragraph(text, style)
    elements.append(para)

    style_sheet = getSampleStyleSheet()
    style = style_sheet['Normal']
    text = "This report was auto-generated through EDUSTACK RVITM"
    para = Paragraph(text, style)
    elements.append(para)

    doc.build(elements)
    
    buffer.seek(0)
    return buffer

def progress_pdf():

    st.markdown("<div style='text-align:center;'><h2> </h2></div>", unsafe_allow_html=True,)

    st.markdown("<div style='text-align:center;'><h3> 📈 ATTENDANCE REPORT GENERATOR </h3></div>", unsafe_allow_html=True,)
    st.markdown("<div style='text-align:center;'><h1> </h1></div>", unsafe_allow_html=True,)

    Branch_Choice = st.selectbox("Choose Branch: ",["COMPUTER SCIENCE & ENGINEERING","INFORMATION SCIENCE & ENGINEERING","ELECTRONICS & COMMUNICATION ENGINEERING", "MECHANICAL ENGINEERING","MASTER OF COMPUTER APPLICATIONS"])

    test_choice = st.selectbox("Choose the test: ",["ATTENDANCE REPORT"])   

    submission_d = st.date_input("The Ward Should Submit the Signed Attendance Report to Counsellor Before:", date.today())
    day = submission_d.day
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    submission_d = submission_d.strftime(f"%d{suffix} %b, %Y")
    
    semester = st.selectbox("Select the Semester: ",[" I Semester BE  "," II Semester BE  ", " III Semester BE  "," IV Semester BE ", "V Semester BE", "VI Semester BE","VII Semester BE"," VIII Semester BE", "I Semester MCA", "II Semester MCA", "III Semester MCA", "IV Semester MCA","V Semester MCA","VI Semester MCA"])   
    no_of_subjects = st.selectbox("Select the no of Subjects: ",[1,2,3,4,5,6,7,8,9,10,11])   
    note = st.text_area("General Note (If any*):",placeholder="example: Attendace considered up till 17th March 2023")

    uploaded_file = st.file_uploader("Upload the Marks Sheet Excel File for the test:", type=["xlsx"])   

    
    if uploaded_file is not None:
      tab1,tab2, tab3 = st.tabs(["Generate & Download Report", 'Preview Report' ,"Confirm & Send Email"])
      with tab1:
        df = pd.read_excel(uploaded_file)

        progress_bar = st.progress(0, text = 'Generating Report')
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for i in range(2, df.shape[0]):
                buffer = generate_pdf(df, i, Branch_Choice, test_choice, submission_d, semester, no_of_subjects, note)
                file_name = f"{df.iloc[i, 1]}.pdf"
                zip_file.writestr(file_name, buffer.getvalue())
        
                progress_value = int((i - 1) / (df.shape[0] - 2) * 100)
                progress_bar.progress(progress_value, text = 'Generating Report')
        
        # Generate a download link for the zip file
        zip_name = ""+test_choice+"."+semester+".zip"
        b64 = base64.b64encode(zip_buffer.getvalue()).decode()
        download_link = f'<a href="data:application/zip;base64,{b64}" download="{zip_name}">click here to begin download</a>'
        st.markdown(download_link, unsafe_allow_html=True)


      with tab2:
      
        df = pd.read_excel(uploaded_file)
        st.write("Generating Preview of Attendance Report...")
        
        # Show a progress bar while the PDFs are being generated
        progress_bar = st.progress(0)
        
        # Generate the PDFs for each student and store it in a dictionary with the student name as the key
        pdfs = {}
        for i in range(2, df.shape[0]):
            buffer = generate_pdf(df, i, Branch_Choice, test_choice, submission_d,semester,no_of_subjects,note)
            file_name = f"{df.iloc[i, 1]}.pdf"
         
            b64 = base64.b64encode(buffer.getvalue()).decode()
            pdfs[file_name] = b64
            
            progress_value = int((i - 1) / (df.shape[0] - 2) * 100)
            progress_bar.progress(progress_value)
        
        # Show a selectbox to select the PDF to preview
        selected_pdf = st.selectbox("Select a student", list(pdfs.keys()))
        if selected_pdf is not None:
            b64 = pdfs[selected_pdf]
            st.write("""
            <iframe
                src="data:application/pdf;base64,{b64}"
                style="border: none; width: 100%; height: 970px;"
            ></iframe>
            """.format(b64=b64), unsafe_allow_html=True)

      with tab3:

       df = pd.read_excel(uploaded_file)
       SMTP_SERVER = "smtp.gmail.com"
       SMTP_PORT = 587
       with st.form("login_form"):
         st.write("Enter the mail ID login from which you want to send the mail:")
         
         SMTP_USERNAME = st.text_input('Input mail ID',help="Credentials are safe and not stored anywhere")
         SMTP_PASSWORD = st.text_input('Input password',type='password')
         st.checkbox("I confirm that the Report generated are correct")
         submitted = st.form_submit_button("Confirm & send email")

       if submitted:
        st.write("Sending Email...")
        total_emails = df.shape[0] - 2
        email_sent = 0
        progress_bar = st.progress(0)

        for i in range(2, df.shape[0]):
            buffer = generate_pdf(df, i, Branch_Choice, test_choice, submission_d,semester,no_of_subjects,note)
            file_name = f"{df.iloc[i, 1]}.pdf"
            email = df.iloc[i, 4]
            cc_email = df.iloc[i,6]
            father = str(df.iloc[i, 3])
            student_name = str(df.iloc[i, 1])

            msg = MIMEMultipart()
            msg['From'] = SMTP_USERNAME
            msg['To'] = COMMASPACE.join([email])
            msg['Cc'] = COMMASPACE.join([cc_email])
            msg['Subject'] = ""+test_choice+"\u00a0 "+semester
            
            body = "Dear <b>"+father+"</b> ,<br><br>Herewith enclosed the <b>"+semester+" "+test_choice+"</b>\u00a0  of your ward <b>"+student_name+"</b><br><br>Thanks & Regards,<br><b>RVITM</b>"
            text = MIMEText(body,'html')
            msg.attach(text)
        
            # Attach the generated PDF
            part = MIMEBase('application', "octet-stream")
            part.set_payload((buffer.getvalue()))
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=file_name)
            msg.attach(part)
        
            # Connect to the SMTP server and send the email
            smtpObj = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            smtpObj.ehlo()
            smtpObj.starttls()
            smtpObj.login(SMTP_USERNAME, SMTP_PASSWORD)
            smtpObj.sendmail(SMTP_USERNAME, [email,cc_email], msg.as_string())
            smtpObj.quit()
        
            st.write("Email sent to\u00a0"+student_name+"\u00a0 parent's mail - ", email)

            email_sent += 1
            progress_bar.progress(email_sent / total_emails)
        st.success("All Attendance Reports sent successfully")


progress_pdf()
