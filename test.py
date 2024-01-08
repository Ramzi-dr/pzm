import asyncio
import os
from docx2pdf import convert
import time
from payloadCollection import PayloadCollection
from dateCalculator import calculate_week_and_day


class DocxToPdf:
    def __init__(self):
        try:
            (
                self.current_year,
                self.current_month,
                self.current_week,
                self.current_day,
            ) = calculate_week_and_day()
            self.docx_folder = ".\doors_wordFiles"

            self.week_file_list = []
            self.month_file_list = []
            self.year_file_list = []
            self.get_fileToConvert()
            self.all_docx_list = (
                self.week_file_list + self.month_file_list + self.year_file_list
            )
        except Exception as e:
            print(f"Error in __init__: {e}")
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
            else:
                print("dont exist")
        except Exception as e:
            print(f"Error in get_fileToConvert: {e}")
            pass

    async def convert(self, docx_path, pdf_output):
        try:
            pdf_path = os.path.join(pdf_output)
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
                print("removing")
            convert(docx_path, pdf_output)
            await asyncio.sleep(1)
        except Exception as e:
            print(f"Error in convert: {e}")
            pass

    async def doc_convertor(self, task):
        try:
            pdfFiles_dir = ""
            docx_list_to_convert = []

            if task == "week_pdf":
                pdfFiles_dir = "week_PdfFiles"
                docx_list_to_convert = self.week_file_list

            if task == "month_pdf":
                pdfFiles_dir = "month_PdfFiles"
                docx_list_to_convert = self.month_file_list
            elif task == "year_pdf":
                pdfFiles_dir = "year_PdfFiles"
                docx_list_to_convert = self.year_file_list

            pdf_directory = f"doors_pdfFiles/{pdfFiles_dir}"
            if not os.path.exists(pdf_directory):
                os.makedirs(pdf_directory)
                print(f"{pdf_directory}  is created")

            tasks = []
            if docx_list_to_convert:
                for docx in docx_list_to_convert:
                    try:
                        docx_path = f"{self.docx_folder}/{docx}"
                        word_basename = os.path.splitext(os.path.basename(docx))[0]
                        pdf_filename = word_basename[:-1] + ".pdf"
                        output_pdf_path = os.path.join(
                            f"{pdf_directory}/{pdf_filename}"
                        )
                        tasks.append(self.convert(docx_path, output_pdf_path))
                    except Exception as e:
                        print(f"Error in doc_convertor loop: {e}")
                        pass

                start_time = time.time()
                print(f"starting converting 10 docx files to .pdf at: {start_time} ")
                await asyncio.gather(*tasks)
                if os.path.exists(output_pdf_path):
                    end_time = time.time()
                    print(f"Finishing converting 10 docx files to .pdf at: {end_time} ")
                    time_difference = int(end_time - start_time)
                    print(
                        f"Ready to send email after {time_difference} seconds of converting time"
                    )
                else:
                    print("list is empty")
        except Exception as e:
            print(f"Error in doc_convertor: {e}")
            pass
