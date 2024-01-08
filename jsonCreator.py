import asyncio
import json
import aiofiles
import locale
from dateCalculator import calculate_week_and_day, swiss_date
from emailManager import send_email
from openTimeCalculator import calculate_total_open_time
locale.setlocale(locale.LC_TIME, "de_CH") 

class JsonData_creator:
    def __init__(self, door_Fullname, door_name, door_number, door_floor, door_house):
        self.door_FULLname = door_Fullname
        self.door_name = door_name
        self.door_number = door_number
        self.door_floor = door_floor
        self.door_house = door_house
        self.file_path = f"/home/adminbst/Documents/pzmCode/doorsJsonFiles/{door_Fullname}.json"
        self.door_data = {}
        self.lock = asyncio.Lock()

    async def initialize_data(self):
        try:
            self.door_data = await self.load_from_file()
        except Exception as e:
            send_email(subject=f'the exeption in initialize data{e}\n and the door_data in{self.door_data} ')
          #  print(f'the exeption in initialize data{e}\n and the door_data in{self.door_data}  ')
    async def update_data(self, action, time):
        year, month, week_number, day_of_week = calculate_week_and_day()
        year = str(year)
        week_number = f"KW {week_number}"
        date = swiss_date()

        async with self.lock:
            if self.door_FULLname not in self.door_data:
                self.door_data[self.door_FULLname] = {}

            # year configuration
            if year not in self.door_data[self.door_FULLname]:
                self.door_data[self.door_FULLname][year] = {}
            if "24H year percentage" not in self.door_data[self.door_FULLname][year]:
                self.door_data[self.door_FULLname][year]["24H year percentage"] = 0.0
            if (
                "24H year total open (sec)"
                not in self.door_data[self.door_FULLname][year]
            ):
                self.door_data[self.door_FULLname][year][
                    "24H year total open (sec)"
                ] = 0.0
            if (
                "08:00-18:30 year percentage"
                not in self.door_data[self.door_FULLname][year]
            ):
                self.door_data[self.door_FULLname][year][
                    "08:00-18:30 year percentage"
                ] = 0.0
            if (
                "08:00-18:30 year total open (sec)"
                not in self.door_data[self.door_FULLname][year]
            ):
                self.door_data[self.door_FULLname][year][
                    "08:00-18:30 year total open (sec)"
                ] = 0.0

            if month not in self.door_data[self.door_FULLname][year]:
                self.door_data[self.door_FULLname][year][month] = {}
            if (
                "24H month percentage"
                not in self.door_data[self.door_FULLname][year][month]
            ):
                self.door_data[self.door_FULLname][year][month][
                    "24H month percentage"
                ] = 0.0
            if (
                "24H month total open (sec)"
                not in self.door_data[self.door_FULLname][year][month]
            ):
                self.door_data[self.door_FULLname][year][month][
                    "24H month total open (sec)"
                ] = 0.0
            if (
                "08:00-18:30 month percentage"
                not in self.door_data[self.door_FULLname][year][month]
            ):
                self.door_data[self.door_FULLname][year][month][
                    "08:00-18:30 month percentage"
                ] = 0.0
            if (
                "08:00-18:30 month total open (sec)"
                not in self.door_data[self.door_FULLname][year][month]
            ):
                self.door_data[self.door_FULLname][year][month][
                    "08:00-18:30 month total open (sec)"
                ] = 0.0

            if week_number not in self.door_data[self.door_FULLname][year][month]:
                self.door_data[self.door_FULLname][year][month][week_number] = {}
            if (
                "24H week percentage"
                not in self.door_data[self.door_FULLname][year][month][week_number]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][
                    "24H week percentage"
                ] = 0.0
            if (
                "24H week total open (sec)"
                not in self.door_data[self.door_FULLname][year][month][week_number]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][
                    "24H week total open (sec)"
                ] = 0.0
            if (
                "08:00-18:30 week percentage"
                not in self.door_data[self.door_FULLname][year][month][week_number]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][
                    "08:00-18:30 week percentage"
                ] = 0.0
            if (
                "08:00-18:30 week total open (sec)"
                not in self.door_data[self.door_FULLname][year][month][week_number]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][
                    "08:00-18:30 week total open (sec)"
                ] = 0.0

            if date not in self.door_data[self.door_FULLname][year][month][week_number]:
                self.door_data[self.door_FULLname][year][month][week_number][date] = {}

            if (
                day_of_week
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ] = {}
            if (
                "24H day percentage"
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["24H day percentage"] = 0.0
            if (
                "24H day total open (sec)"
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["24H day total open (sec)"] = 0.0
            if (
                "08:00-18:30 day percentage"
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["08:00-18:30 day percentage"] = 0.0
            if (
                "08:00-18:30 day total open (sec)"
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["08:00-18:30 day total open (sec)"] = 0.0

            if (
                "open"
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["open"] = []
            if (
                "close"
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["close"] = []
            if (
                "is open"
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["is open"] = False
            open_list = self.door_data[self.door_FULLname][year][month][week_number][
                date
            ][day_of_week].get("open", [])
            close_list = self.door_data[self.door_FULLname][year][month][week_number][
                date
            ][day_of_week].get("close", [])

            if len(open_list) < len(close_list):
                theLastCloseTime = close_list[-1]

                close_list = close_list[: -(len(close_list) - (len(open_list) - 1))]
                close_list.append(theLastCloseTime)

                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["close"] = close_list
            if len(open_list) > (len(close_list)) and action == "open":
                open_list = open_list[: -(len(open_list) - (len(close_list)))]
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ]["open"] = open_list

            if (
                time
                not in self.door_data[self.door_FULLname][year][month][week_number][
                    date
                ][day_of_week][action]
            ):
                self.door_data[self.door_FULLname][year][month][week_number][date][
                    day_of_week
                ][action].append(time)
                if action == "open":
                    self.door_data[self.door_FULLname][year][month][week_number][date][
                        day_of_week
                    ]["is open"] = True
                else:
                    self.door_data[self.door_FULLname][year][month][week_number][date][
                        day_of_week
                    ]["is open"] = False

            if action == "close":
                await calculate_total_open_time(
                    self.door_data,
                    self.door_FULLname,
                    self.door_name,
                    self.door_number,
                    self.door_floor,
                    self.door_house,
                    year,
                    month,
                    week_number,
                    date,
                    day_of_week,
                    open_list,
                    close_list,
                )

    async def save_to_file(self):
        async with self.lock:
            try:
                async with aiofiles.open(self.file_path, "w") as json_file:
                    await json_file.write(json.dumps(self.door_data, indent=4))
            except Exception as e:
                send_email(subject='Error',message=f'exception on save_to_file() at jsoncreator.py:{e} ')
                

    async def load_from_file(self):
        try:
            async with aiofiles.open(self.file_path, "r") as json_file:
                content = await json_file.read()
                return json.loads(content)
        except json.JSONDecodeError as e:
            send_email(subject='Error', message=f'exception on load_from_file() at jsoncreator.py:{e} ')
            
            return {}
        except FileNotFoundError:
            # Create an empty file if it doesn't exist
            async with aiofiles.open(self.file_path, "w") as json_file:
                await json_file.write("{}")
            return {}

