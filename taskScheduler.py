import asyncio
from datetime import datetime, timedelta, time
from timerTasks import TimerTasks
import datetime
from doc_manager.docxToPdfConvertor import DocxToPdf
from emailManager import send_email
import locale
locale.setlocale(locale.LC_TIME, "de_CH") 

async def virtual_openClose_tasks():
    timerTasker = TimerTasks()
    
    target_times = ["23:59:55", "07:59:55", "18:29:55"]
    
    while True:
        now = datetime.datetime.now().time()
        now_str = now.strftime("%H:%M:%S")  # Format as HH:MM:SS
        if now_str in target_times: 
            try:
                await timerTasker.stillOpen_doorList()    
            except Exception as e:     
                send_email(subject='Error',message=f'error in Pzm Event Server in virtual_openClose_tasks() at taskScheduler.py.py\n {e}  ')   
                pass
        await asyncio.sleep(1)  # Check every second


async def send_emailWithPdf_tasks():
    timerTasker = TimerTasks()
    monday_email_send = False
    month_email_send = False
    year_email_send = False

    while True:
        now = datetime.datetime.now()
        current_time = now.replace(second=0, microsecond=0).time()
        day_of_week = now.strftime("%A")
        first_day_of_month = now.day == 1
        first_day_of_year = now.month == 1 and now.day == 1
        monday_emailTime = time(7, 0)
        reset_monday_emailTime = time(7,2 )
        firstDay_ofTheMonth_emailTime = time(7, 5)
        reset_firstDay_ofTheMonth_emailTime = time(7, 7)
        firstDay_ofTheYear_emailTime = time(7, 10)
        reset_firstDay_ofTheYear_emailTime = time(7, 12)

        # print(monday_emailTime)
        # print(current_time)
        # print(reset_monday_emailTime)
        # print(monday_email_send)
        # print(day_of_week)
        # print("Doing something at 07:00 on Monday")
        # docx_to_pdf = DocxToPdf()
        # tasks =['week_pdf','month_pdf','year_pdf'] 
        # for task  in tasks:
        #     await docx_to_pdf.doc_convertor(task=task)
        
        if (
            day_of_week == "Montag"
            and monday_emailTime == current_time
            and not monday_email_send
        ):
            try:
              #  print("sending Monday email")
                docx_to_pdf = DocxToPdf()
                await docx_to_pdf.doc_convertor(task="week_pdf")
                monday_email_send = True
            except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in Monday Sending Email Task  at taskScheduler.py\n {e}  ')
                pass
        if (
            day_of_week == "Montag"
            and reset_monday_emailTime == current_time
            and monday_email_send
        ):
            monday_email_send = False

        #  send email first day of the month
        if (
            first_day_of_month
            and firstDay_ofTheMonth_emailTime==current_time
            and not month_email_send
        ):
            try:
                docx_to_pdf = DocxToPdf()
                await docx_to_pdf.doc_convertor(task="month_pdf")
                month_email_send =True
            except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in first day  of the month Sending Email Task  at taskScheduler.py\n {e}  ')
                pass

        if (
            first_day_of_month
            and current_time == reset_firstDay_ofTheMonth_emailTime
            and month_email_send
        ):
            month_email_send = False

        # send email the first day of january
        if (
            first_day_of_year
            and firstDay_ofTheYear_emailTime==current_time 
            and not year_email_send
        ):
            try:
                docx_to_pdf = DocxToPdf()
                await docx_to_pdf.doc_convertor(task="year_pdf")
                year_email_send = True
            except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in first day of January Sending Email Task  at taskScheduler.py\n {e}  ')
                pass
        if (
            first_day_of_year
            and current_time == reset_firstDay_ofTheYear_emailTime
            and year_email_send
        ):
            year_email_send = False

        await asyncio.sleep(20)  # Check every 20 second
