import asyncio
import json
import os
import aiofiles
from pathlib import Path
from dateCalculator import calculate_week_and_day
from messageFilter import Message_manager
from payloadCollection import PayloadCollection
from emailManager import send_email


class TimerTasks:
    def __init__(self):
        self.folder_path = "doorsJsonFiles"
        self.folder_list = []
        self.current_year = None
        self.current_month = None
        self.current_week = None
        self.day_of_week = None
        # Introduce a semaphore to control concurrent calls
       # self.semaphore = asyncio.Semaphore(1)

    def get_time(self):
        try:
            (
                self.current_year,
                self.current_month,
                self.current_week,
                self.day_of_week,
            ) = calculate_week_and_day()
        except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in get_time() at timerTasks.py\n {e}  ')
                pass

    def get_door_List(self):
        try:
            doorList = []
            if os.path.exists(self.folder_path):
                
                for file_name in os.listdir(self.folder_path):
                    if os.path.isfile(os.path.join(self.folder_path, file_name)):
                        self.folder_list.append(file_name)
            return doorList
        except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in get_door_List() at timerTasks.py\n {e}  ')
                pass

    def get_doorId_list(self, doorNameToCloseOpen_list):
        try:
            doorIdList = []
            for door in PayloadCollection.doorDict:
                for house, door_data in door.items():
                    for doorName in doorNameToCloseOpen_list:
                        for k, v in door_data.items():
                            if {"floor", "number"}.issubset(door_data.keys()):
                                floor = door_data["floor"]
                                number = door_data["number"]
                                doorName_inJsonFile = f"{house} {floor} {k} {number}"
                                if doorName == doorName_inJsonFile:
                                    if door_data[k] not in doorIdList:
                                        doorIdList.append(door_data[k])
            return doorIdList
        except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in get_doorId_list() at timerTasks.py\n {e}  ')
                pass


    async def virtual_close_open(self, doorIdToCloseList):
        if doorIdToCloseList is not None:
            try:
                for doorId in doorIdToCloseList:
                    virtual_close_message = PayloadCollection.close_door(deviceid=doorId)
                    virtual_open_message = PayloadCollection.open_door(deviceid=doorId)
                    messageInstance = Message_manager()
                    await messageInstance.data_manager(data=virtual_close_message)
                    await asyncio.sleep(5)
                    await messageInstance.data_manager(data=virtual_open_message)
                    
                    
            except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in virtual_close() at timerTasks.py\n {e}  ')
                pass
                

    async def stillOpen_doorList(self):        
        self.get_time()
        self.get_door_List()
        doors_data = []
        doorNameToCloseOpen_list = []
        for doors in self.folder_list:
            doors_data.append(await self.load_from_file(f"/home/adminbst/Documents/pzmCode/doorsJsonFiles/{doors}"))          
        for entry in doors_data:
            try:
                for door, door_data in entry.items():
                    for year, year_data in door_data.items():
                        if int(year) == self.current_year:
                            for month, month_data in year_data.items():
                                if month == self.current_month:
                                    for week, week_data in month_data.items():
                                        str_current_week = f"KW {self.current_week}"
                                        if week == str_current_week:
                                            for date, date_data in week_data.items():
                                                if date not in [
                                                    "week percentage",
                                                    "week total open",
                                                    "24H week percentage",
                                                    "24H week total open (sec)",
                                                    "08:00-18:30 week percentage",
                                                    "08:00-18:30 week total open (sec)",
                                                ]:
                                                    for day, day_data in date_data.items():
                                                        if "is open" in day_data:
                                                            if day_data["is open"]:
                                                                if (
                                                                door
                                                                not in doorNameToCloseOpen_list
                                                            ):
                                                                    doorNameToCloseOpen_list.append(
                                                                        door
                                                                    )                                                               
            except Exception as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in stillOpen_doorList() at timerTasks.py\n {e}  ')
                pass
        if doorNameToCloseOpen_list:
            doorsIdToCloseList = self.get_doorId_list(
                doorNameToCloseOpen_list=doorNameToCloseOpen_list
            )
            # Create a list of coroutine objects (tasks) without starting them
            tasks = [
                self.virtual_close_open(doorIdToCloseList=[doorId])
                for doorId in doorsIdToCloseList
            ]
            # Gather tasks concurrently
            await asyncio.gather(*tasks)
            # Add a delay between tasks if needed
            await asyncio.sleep(5)

    async def load_from_file(self, file_path):
        file_path = Path(file_path)
        if file_path.exists():  
            try:
                async with aiofiles.open(file_path, "r") as json_file:
                    content = await json_file.read()
                    return json.loads(content)
            except FileNotFoundError as e:
                send_email(subject='Error',message=f'error in Pzm Event Server in load_from_file() at timerTasks.py\n {e}\n this is the file path :{file_path}   ')
                pass
        else:
            pass

