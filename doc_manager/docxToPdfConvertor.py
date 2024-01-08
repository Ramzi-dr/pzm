import asyncio
import os
from docx import Document
from docx2pdf import convert
import subprocess
import time
from payloadCollection import PayloadCollection
from dateCalculator import calculate_week_and_day
from emailManager import send_email


class DocxToPdf:
    def __init__(self):
        try:
            (
                self.current_year,
                self.current_month,
                self.current_week,
                self.current_day,
            ) = calculate_week_and_day()
            self.docx_folder = "./doors_wordFiles"

            self.week_file_list = []
            self.month_file_list = []
            self.year_file_list = []
            self.get_fileToConvert()
            self.all_docx_list = (
                self.week_file_list + self.month_file_list + self.year_file_list
            )
        except Exception as e:
            send_email(subject='Error',message=f'error in Pzm Event Server in __init__: at docxToPdfConvertor.py\n {e}  ')
            pass

    def get_fileToConvert(self):
        try:
            last_year = self.current_year - 1
            last_week = self.current_week - 1 if self.current_week > 1 else 52
            last_month_index = (
                PayloadCollection.german_months.index(self.current_month) - 1
            )
            last_month = PayloadCollection.german_months[last_month_index]
            str_last_year = str(last_year)
            str_current_year = self.current_year
            str_last_week = str(last_week)
            str_current_week = self.current_week
            str_current_year = str(self.current_year)
            correctYear_for_lastWeek = (
                str_last_year if last_week == 52 else str_current_year
            )
            correctYear_for_lastMonth = (
                str_last_year if last_month == "Dezember" else str_current_year
            )

            if os.path.exists(self.docx_folder):
                for files in os.listdir(self.docx_folder):
                    words = files.split()
                    if all(
                        item in words
                        for item in ["KW", str_last_week, correctYear_for_lastWeek]
                    ):
                        self.week_file_list.append(files)
                    if all(
                        item in words
                        for item in [correctYear_for_lastMonth, last_month]
                    ):
                        self.month_file_list.append(files)
                    if str_last_year in words and not any(
                        item in words
                        for item in [
                            self.current_month,
                            last_month,
                            "KW",
                            str_last_week,
                            str_current_week,
                            str_current_year,
                        ]
                    ):
                        self.year_file_list.append(files)
            
        except Exception as e:
            send_email(subject='Error',message=f'error in Pzm Event Server in get_fileToConvert() at docxToPdfConvertor.py\n {e}  ')
            pass

    async def convert(self, docx_path, pdf_filename,pdf_output_dir):
        try:
            pdf_path = os.path.join(pdf_filename)
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
              #  print("removing")
            pdf_path =os.path.join(pdf_output_dir, pdf_filename)
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', pdf_output_dir, docx_path])
             # Rename the generated PDF file to the specified filename
            pdf_output_path = os.path.join(pdf_output_dir, pdf_filename)
            os.rename(os.path.join(pdf_output_dir, f"{os.path.splitext(os.path.basename(docx_path))[0]}.pdf"), pdf_output_path)
            await asyncio.sleep(1)
        except Exception as e:
          #  print(f"Error in convert: {e}")
            send_email(subject='Error',message=f'error in Pzm Event Server in convert() at docxToPdfConvertor.py\n {e}  ')
            pass

    async def doc_convertor(self, task):
        try:
            pdfFiles_dir = ""
            docx_list_to_convert = []
            email_subjet = ''
         #   print(f' task is :{task} ')
            if task == "week_pdf":
                pdfFiles_dir = "week_PdfFiles"
                docx_list_to_convert = self.week_file_list
                email_subjet = 'Wochen-Auswertung:'  

            if task == "month_pdf":
             #   print('in month')
                pdfFiles_dir = "month_PdfFiles"
                docx_list_to_convert = self.month_file_list
                email_subjet = 'Monats-Auswertung:' 
            elif task == "year_pdf":
                pdfFiles_dir = "year_PdfFiles:" 
                email_subjet = 'Jahres-Auswertung:' 
                docx_list_to_convert = self.year_file_list

            pdf_output_dir = f"doors_pdfFiles/{pdfFiles_dir}"
            if not os.path.exists(pdf_output_dir):
                os.makedirs(pdf_output_dir)

            tasks = []
            if docx_list_to_convert:
                for docx in docx_list_to_convert:
                    try:
                        docx_path = f"{self.docx_folder}/{docx}"
                        word_basename = os.path.splitext(os.path.basename(docx))[0]
                        document_name_forEmail = word_basename[:-1]
                        pdf_filename = word_basename[:-1] + ".pdf"
                        
                        
                        tasks.append(self.convert(docx_path=docx_path,pdf_filename=pdf_filename,pdf_output_dir=pdf_output_dir))
                    except Exception as e:
                     #   print(f"Error in doc_convertor loop: {e}")
                        send_email(subject='Error',message=f'error in Pzm Event Server in doc_converter() at docxToPdfConvertor.py\n {e}  ')
                        pass

                await asyncio.gather(*tasks)

                # List to keep track of processed documents
                processed_documents = []
                for docx in docx_list_to_convert:
                    word_basename = os.path.splitext(os.path.basename(docx))[0]
                    document_name_forEmail = word_basename[:-1]
                    pdf_filename = word_basename[:-1] + ".pdf"
                    output_pdf_path = os.path.join(f"{pdf_output_dir}/{pdf_filename}")
                    subject = f'Türen {email_subjet} {word_basename} '

                    if (
                        os.path.exists(output_pdf_path)
                        and document_name_forEmail not in processed_documents
                    ):
                        
                        send_email(
                            subject=subject,
                            message=PayloadCollection.pdf_mail_text,
                            attachment_path=output_pdf_path,
                        )
                        # Add the document to the processed list
                        processed_documents.append(document_name_forEmail)

                    else:
                        send_email(subject='Error',message=f'error in Pzm Event Server in doc_converter() at docxToPdfConvertor.py\n {e}  ')
                

        except Exception as e:
            send_email(subject='Error',message=f'error in Pzm Event Server in doc_converter() at docxToPdfConvertor.py\n {e}  ')
            pass
