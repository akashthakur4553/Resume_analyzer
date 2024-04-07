from flask import *
import nltk
from spire.doc import *
from spire.doc.common import *
import PyPDF2
from docx2python import docx2python


nltk.data.path.append(
    "https://github.com/akashthakur4553/flask_deploy/tree/main/stopwords/stopwords"
)
nltk.data.path.append(
    "https://github.com/akashthakur4553/flask_deploy/tree/main/nltk_data/stopwords"
)
nltk.download("stopwords")
# nltk.download('stopwords', download_dir='/opt/render/nltk_data')
import spacy
from pyresparser import ResumeParser
import os
import xlsxwriter

workbook = xlsxwriter.Workbook("output.xlsx")
worksheet = workbook.add_worksheet()
# from docx import Document
worksheet.write(0, 0, "Email")
worksheet.write(0, 1, "Phone Number")
worksheet.write(0, 2, "Name")
worksheet.write(0, 3, "Designation")
worksheet.write(0, 4, "Skills")
worksheet.write(0, 5, "Entire text")
app = Flask(__name__)


@app.route("/")
def main():
    return render_template("index.html")


@app.route("/download")
def download():
    return send_file("output.xlsx", as_attachment=True)


@app.route("/upload", methods=["POST"])
def upload():
    if request.method == "POST":
        data_l = []
        # Get the list of files from the webpage
        files = request.files.getlist("file")
        row = 1

        # Iterate for each file in the files List, and Save them
        for file in files:
            file.save(file.filename)
            filed = file.filename
            print(filed)
            if ".doc" in filed:
                document = Document()
                document.LoadFromFile(filed)
                document.SaveToFile("WordToPdf.pdf", FileFormat.PDF)
                document.Close()
                filed = "WordToPdf.pdf"
            if ".pdf" in filed:
                pdf_file = open(filed, "rb")
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                num_pages = len(pdf_reader.pages)
                pdf_text = ""
                for page in range(num_pages):
                    page_obj = pdf_reader.pages[page]
                    pdf_text += page_obj.extract_text()
                pdf_file.close()
            elif ".docx" in filed:
                doc_result = docx2python(filed)
                pdf_text = doc_result.text

            # print(pdf_text)

            # emails = re.findall(email_pattern, pdf_text)
            data = ResumeParser(filed).get_extracted_data()
            data_l.append(data["name"])
            worksheet.write(row, 0, str(data["email"]))
            worksheet.write(row, 1, str(data["mobile_number"]))
            worksheet.write(row, 2, str(data["name"]))
            # worksheet.write(row, 1, str(emails))
            worksheet.write(row, 3, str(data["designation"]))
            worksheet.write(row, 4, str(data["skills"]))
            worksheet.write(
                row,
                5,
                str(
                    pdf_text.replace(
                        "Evaluation Warning: The document was created with Spire.Doc for Python.",
                        "",
                    )
                ),
            )
            row = row + 1
            # Delete the saved file
            os.remove(file.filename)
        workbook.close()
        print(data_l)
        return render_template("download.html")


# if __name__ == "__main__":
#     app.run(debug=True)
