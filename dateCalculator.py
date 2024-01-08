import datetime
import locale
from datetime import datetime as dt
import calendar

from emailManager import send_email
locale.setlocale(locale.LC_TIME, "de_CH") 

def swiss_date():
    current_date = dt.now()
    swissDate = current_date.strftime("%d.%m.%Y")
    return swissDate


def calculate_week_and_day():
    date = swiss_date()
    locale.setlocale(locale.LC_TIME, "de_CH") 
    # Calculate the week number using isocalendar
    year, week_num, weekday = dt.strptime(date, "%d.%m.%Y").isocalendar()
    week_number = week_num
    current_year = year
    # Get the day of the week
    day_of_week = dt.strptime(date, "%d.%m.%Y").strftime("%A")
    # get the month in german
    current_month = dt.now().strftime("%B")

    return current_year, current_month, week_number, day_of_week


def get_dates_from_week():
    date_list = []
    current_year, current_month, week_number, day_of_week = calculate_week_and_day()

    locale.setlocale(locale.LC_TIME, "de_CH") 

    first_day = datetime.date.fromisocalendar(current_year, week_number, 1)
    dates = [first_day + datetime.timedelta(days=i) for i in range(7)]
    for i, date in enumerate(dates):
        the_date = date.strftime("%d.%m.%y")
        if the_date not in date_list:
            date_list.append(the_date)

    return date_list, week_number


def get_days_in_month():
    try:
        # Get the current year and month
        current_date = datetime.date.today()
        year = current_date.year
        month = current_date.month

        # Get the number of days in the current month
        days_in_month = calendar.monthrange(year, month)[1]

        return days_in_month
    except (ValueError, TypeError) as e:
        send_email(subject='Error',message=f'error in Pzm Event Server in get_days_in_month() at dateCalculator.py\n {e}  ')
                
        return 0


def get_days_in_year():
    try:
        current_date = datetime.date.today()
        year = current_date.year
        total_days = sum(
            calendar.monthrange(int(year), month)[1] for month in range(1, 13)
        )
        return total_days
    except (ValueError, TypeError) as e:
        send_email(subject='Error',message=f'error in Pzm Event Server in get_days_in_year() at dateCalculator.py\n {e}  ')
        return 0
