import calendar
import datetime as dt


def number_of_month_days():
    current_date = dt.date.today()
    # Get the number of days in the current month
    number_of_days = calendar.monthrange(current_date.year, current_date.month)[1]
    return int(number_of_days)


def number_of_year_days():
    current_year = dt.date.today().year
    # Check if it's a leap year
    if (current_year % 4 == 0 and current_year % 100 != 0) or (current_year % 400 == 0):
        number_of_days = 366  # Leap year has 366 days
    else:
        number_of_days = 365  # Non-leap year has 365 days

    return int(number_of_days)
