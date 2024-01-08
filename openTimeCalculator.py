import datetime as dt
import locale
from datetime import datetime
from calendarControl import number_of_month_days, number_of_year_days
from doc_manager.docManager import DocManager


async def calculate_total_open_time(
    door_data,
    door_Fullname,
    door_name,
    door_number,
    door_floor,
    door_house,
    year,
    month,
    week_number,
    date,
    day_of_week,
    open_list,
    close_list,
):
    async def calculator():
        total_open_time = door_data[door_Fullname][year][month][week_number][date][
            day_of_week
        ]["24H day total open (sec)"]
        total_open_work_time = door_data[door_Fullname][year][month][week_number][date][
            day_of_week
        ]["08:00-18:30 day total open (sec)"]
        total_open_inWeek = 0.0
        total_openWork_inWeek = 0.0
        total_open_inMonth = 0.0
        total_openWork_inMonth = 0.0
        total_open_inYear = 0.0
        workYear_open = 0.0
        time_format = "%H:%M:%S"
        start_work_time = datetime.strptime("08:00:00", time_format)
        end_work_time = datetime.strptime("18:30:00", time_format)

        # Assuming open_list[-1] and close_list[-1] contain the time in the format HH:MM:SS
        new_open_time_str = open_list[-1]
        new_close_time_str = close_list[-1]

        # Parse the time strings into datetime objects with a dummy date (e.g., 1900-01-01)
        new_open = datetime.strptime(new_open_time_str, time_format)
        new_close = datetime.strptime(new_close_time_str, time_format)
        new_open_time = (new_close - new_open).total_seconds()
        total_open_time += new_open_time
        door_data[door_Fullname][year][month][week_number][date][day_of_week][
            "24H day total open (sec)"
        ] = total_open_time
        percentage = (total_open_time / 86400) * 100
        door_data[door_Fullname][year][month][week_number][date][day_of_week][
            "24H day percentage"
        ] = percentage
        if (
            start_work_time.time() <= new_open.time() <= end_work_time.time()
            and new_close.time() <= end_work_time.time()
        ):
            openTime_in_workTime = (new_close - new_open).total_seconds()
            total_open_work_time += openTime_in_workTime
            door_data[door_Fullname][year][month][week_number][date][day_of_week][
                "08:00-18:30 day total open (sec)"
            ] = total_open_work_time
            percentage_workTime = (total_open_work_time / 37800) * 100

            door_data[door_Fullname][year][month][week_number][date][day_of_week][
                "08:00-18:30 day percentage"
            ] = percentage_workTime

        # week calculation
        for v in door_data[door_Fullname][year][month][week_number].values():
            if type(v) is dict:
                for day in v.values():
                    day_total = day["24H day total open (sec)"]
                    workDay_open = day["08:00-18:30 day total open (sec)"]
                    total_open_inWeek += day_total
                    total_openWork_inWeek += workDay_open
        door_data[door_Fullname][year][month][week_number][
            "24H week total open (sec)"
        ] = total_open_inWeek
        week_percentage = (total_open_inWeek / (7 * 86400)) * 100
        door_data[door_Fullname][year][month][week_number][
            "24H week percentage"
        ] = week_percentage

        door_data[door_Fullname][year][month][week_number][
            "08:00-18:30 week total open (sec)"
        ] = total_openWork_inWeek
        workWeek_percentage = (total_openWork_inWeek / (7 * 37800)) * 100
        door_data[door_Fullname][year][month][week_number][
            "08:00-18:30 week percentage"
        ] = workWeek_percentage
        # month calculation
        for ke, va in door_data[door_Fullname][year][month].items():
            if ke not in [
                "24H month percentage",
                "24H month total open (sec)",
                "08:00-18:30 month percentage",
                "08:00-18:30 month total open (sec)",
            ]:
                week_open = va["24H week total open (sec)"]
                workWeek_open = va["08:00-18:30 week total open (sec)"]
                total_open_inMonth += week_open
                total_openWork_inMonth += workWeek_open
        door_data[door_Fullname][year][month][
            "24H month total open (sec)"
        ] = total_open_inMonth
        month_percentage = (total_open_inMonth / (number_of_month_days() * 86400)) * 100
        door_data[door_Fullname][year][month]["24H month percentage"] = month_percentage

        door_data[door_Fullname][year][month][
            "08:00-18:30 month total open (sec)"
        ] = total_openWork_inMonth
        workMonth_percentage = (
            total_openWork_inMonth / (number_of_month_days() * 37800)
        ) * 100
        door_data[door_Fullname][year][month][
            "08:00-18:30 month percentage"
        ] = workMonth_percentage
        # year calculation
        for k, v in door_data[door_Fullname][year].items():
            if k not in [
                "24H year percentage",
                "24H year total open (sec)",
                "08:00-18:30 year percentage",
                "08:00-18:30 year total open (sec)",
            ]:
                month_open = v["24H month total open (sec)"]
                workMonth_open = v["08:00-18:30 month total open (sec)"]
                total_open_inYear += month_open
                workYear_open += workMonth_open

        door_data[door_Fullname][year]["24H year total open (sec)"] = total_open_inYear
        year_percentage = (total_open_inYear / (number_of_year_days() * 86400)) * 100
        door_data[door_Fullname][year]["24H year percentage"] = year_percentage

        door_data[door_Fullname][year][
            "08:00-18:30 year total open (sec)"
        ] = workYear_open
        workYear_percentage = (workYear_open / (number_of_year_days() * 37800)) * 100
        door_data[door_Fullname][year][
            "08:00-18:30 year percentage"
        ] = workYear_percentage
        docManager = DocManager(
            door_Fullname=door_Fullname,
            door_name=door_name,
            door_number=door_number,
            door_house=door_house,
            door_floor=door_floor,
            current_year=year,
            current_month=month,
            current_week_number=week_number,
            dayTotal_openWorkTime=door_data[door_Fullname][year][month][week_number][
                date
            ][day_of_week]["08:00-18:30 day total open (sec)"],
            dayPercentage_openWorkTime=door_data[door_Fullname][year][month][
                week_number
            ][date][day_of_week]["08:00-18:30 day percentage"],
            weekTotal_openWorkTime=door_data[door_Fullname][year][month][week_number][
                "08:00-18:30 week total open (sec)"
            ],
            weekPercentage_openWorkTime=door_data[door_Fullname][year][month][
                week_number
            ]["08:00-18:30 week percentage"],
            monthTotal_openWorkTime=door_data[door_Fullname][year][month][
                "08:00-18:30 month total open (sec)"
            ],
            monthPercentage_openWorkTime=door_data[door_Fullname][year][month][
                "08:00-18:30 month percentage"
            ],
            yearTotal_openWorkTime=door_data[door_Fullname][year][
                "08:00-18:30 year total open (sec)"
            ],
            yearPercentage_openWorkTime=door_data[door_Fullname][year][
                "08:00-18:30 year percentage"
            ],
        )
        await docManager.timeAndPercentage_formatter()
        await docManager.createOrUpdate_Docx()

    if len(open_list) == len(close_list) and open_list and close_list:
        await calculator()
