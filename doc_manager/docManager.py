from dateCalculator import get_days_in_month, get_days_in_year
from doc_manager.wordFileManager import WordFileManager


class DocManager:
    def __init__(
        self,
        door_Fullname,
        door_name,
        door_number,
        door_house,
        door_floor,
        current_year,
        current_month,
        current_week_number,
        dayTotal_openWorkTime,
        dayPercentage_openWorkTime,
        weekTotal_openWorkTime,
        weekPercentage_openWorkTime,
        monthTotal_openWorkTime,
        monthPercentage_openWorkTime,
        yearTotal_openWorkTime,
        yearPercentage_openWorkTime,
    ):
        self.door_Fullname = door_Fullname
        self.door_name = door_name
        self.door_house = door_house
        self.door_number = door_number
        self.door_floor = door_floor
        self.current_year = current_year
        self.current_month = current_month
        self.current_week_number = current_week_number
        self.dayTotal_openWorkTime = dayTotal_openWorkTime
        self.workDay_inSeconds = 37800
        self.workWeek_inSeconds = 264600
        self.dayPercentage_openWorkTime = dayPercentage_openWorkTime
        self.weekTotal_openWorkTime = weekTotal_openWorkTime
        self.weekPercentage_openWorkTime = weekPercentage_openWorkTime
        self.monthTotal_openWorkTime = monthTotal_openWorkTime
        self.monthPercentage_openWorkTime = monthPercentage_openWorkTime
        self.yearTotal_openWorkTime = yearTotal_openWorkTime
        self.yearPercentage_openWorkTime = yearPercentage_openWorkTime
        self.day_openTime = None
        self.day_closeTime = None
        self.week_openTime = None
        self.week_closeTime = None
        self.month_openTime = None
        self.month_closeTime = None
        self.year_openTime = None
        self.year_closeTime = None
        self.day_percentOpen = None
        self.day_percentClose = None
        self.week_percentageOpen = None
        self.month_percentageOpen = None
        self.year_percentageOpen = None
        self.week_percentageClose = None
        self.month_percentageClose = None
        self.year_percentageClose = None

    async def time_formatter(self, sec):
        hours = sec // 3600
        minutes = (sec % 3600) // 60
        seconds = sec % 60
        # Use an f-string to format the time string without leading zeros
        formatted_time = f"{int(hours)}:{int(minutes)}:{int(seconds)}"
        return formatted_time

    async def percentage_formatter(self, percentage):
        rounded_percentage = round(percentage, 2)
        formatted_percentage = f"{rounded_percentage}%"
        return formatted_percentage

    async def timeAndPercentage_formatter(self):
        self.day_openTime = await self.time_formatter(self.dayTotal_openWorkTime)
        self.day_closeTime = await self.time_formatter(
            self.workDay_inSeconds - self.dayTotal_openWorkTime
        )

        workWeek_inSeconds = self.workDay_inSeconds * 7
        self.week_openTime = await self.time_formatter(self.weekTotal_openWorkTime)
        self.week_closeTime = await self.time_formatter(
            workWeek_inSeconds - self.weekTotal_openWorkTime
        )

        days_inMonth = get_days_in_month()
        workMonth_inSeconds = self.workDay_inSeconds * days_inMonth
        self.month_openTime = await self.time_formatter(self.monthTotal_openWorkTime)
        self.month_closeTime = await self.time_formatter(
            workMonth_inSeconds - self.monthTotal_openWorkTime
        )

        days_inYear = get_days_in_year()
        workYear_inSeconds = self.workDay_inSeconds * days_inYear
        self.year_openTime = await self.time_formatter(self.yearTotal_openWorkTime)
        self.year_closeTime = await self.time_formatter(
            workYear_inSeconds - self.yearTotal_openWorkTime
        )

        self.day_percentOpen = await self.percentage_formatter(
            self.dayPercentage_openWorkTime
        )
        self.day_percentClose = await self.percentage_formatter(
            100 - self.dayPercentage_openWorkTime
        )
        self.week_percentageOpen = await self.percentage_formatter(
            self.weekPercentage_openWorkTime
        )
        self.week_percentageClose = await self.percentage_formatter(
            100 - self.weekPercentage_openWorkTime
        )
        self.month_percentageOpen = await self.percentage_formatter(
            self.monthPercentage_openWorkTime
        )
        self.month_percentageClose = await self.percentage_formatter(
            100 - self.monthPercentage_openWorkTime
        )
        self.year_percentageOpen = await self.percentage_formatter(
            self.yearPercentage_openWorkTime
        )

        self.year_percentageClose = await self.percentage_formatter(
            100 - self.yearPercentage_openWorkTime
        )

    async def createOrUpdate_Docx(self):
        wordDoc = WordFileManager(
            door_Fullname=self.door_Fullname,
            door_name=self.door_name,
            door_number=self.door_number,
            door_house=self.door_house,
            door_floor=self.door_floor,
            current_year=self.current_year,
            current_month=self.current_month,
            current_week_number=self.current_week_number,
        )
        wordDoc.createOrUpdate_File(
            close_time=[
                self.day_closeTime,
                self.week_closeTime,
                self.month_closeTime,
                self.year_closeTime,
            ],
            open_time=[
                self.day_openTime,
                self.week_openTime,
                self.month_openTime,
                self.year_openTime,
            ],
            percent_close=[
                self.day_percentClose,
                self.week_percentageClose,
                self.month_percentageClose,
                self.year_percentageClose,
            ],
            percent_open=[
                self.day_percentOpen,
                self.week_percentageOpen,
                self.month_percentageOpen,
                self.year_percentageOpen,
            ],
        )
