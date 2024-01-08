from datetime import datetime
import pytz
import locale
from jsonCreator import JsonData_creator
from payloadCollection import PayloadCollection
from emailManager import send_email
locale.setlocale(locale.LC_TIME, "de_CH") 

class Message_manager:
    
    def __init__(self):
        self.data = None
        self.door_Fullname = "No door full name"
        self.door_name = "No door name"
        self.door_number = "No door number"
        self.door_floor = "No floor"
        self.door_house = "No house"
        self.action = None

    async def message_filter(self, message):
        try:
            if type(message) is dict:
                for key, value in message.items():
                    if key == "params":

                        if value[0] == "ObservedStates":
                            self.data = value[1]["data"]
                            await self.data_manager(data=self.data)
        except Exception as e:
            send_email(subject='Error',message=f'error in Pzm Event Server in message_filter() at messageFilter.py\n {e}  ')
            pass
            

    async def data_manager(self, data):
        try:
            for key, value in data.items():
                if key == "deviceid":
                    for door in PayloadCollection.doorDict:
                        for house, door_data in door.items():
                            for k, v in door_data.items():
                                if value == v:
                                    if "floor" in door_data.keys():
                                        floor = door_data["floor"]
                                    else:
                                        floor = "No floor"
                                    if "number" in door_data.keys():
                                        number = door_data["number"]
                                    else:
                                        number = "no number"
                                    self.door_Fullname = f"{house} {floor} {k} {number}"
                                    self.door_house = house
                                    self.door_name = k
                                    self.door_number = number
                                    self.door_floor = floor

                if key == "events":
                    condition = value[0]
                    for con, conValue in condition.items():
                        if con == "condition":
                            if conValue == "Falling Edge":
                                self.action = "close"
                            if conValue == "Rising Edge":
                                self.action = "open"
                    await self.add_data_to_json()
        except Exception as e:
            send_email(subject='Error',message=f'error in Pzm Event Server in data_manager() at messageFilter.py\n {e}  ')
            pass

    async def add_data_to_json(self):
        try:
            creatorInstance = JsonData_creator(
                self.door_Fullname,
                self.door_name,
                self.door_number,
                self.door_floor,
                self.door_house,
            )
        
            await creatorInstance.initialize_data()            
            await creatorInstance.update_data(
                    action=self.action,
                    time=self.time(),
                )            
            await creatorInstance.save_to_file()
        except Exception as e:
            print(f'error in outter exep :{e} ')
            send_email(subject='Error',message=f'error in Pzm Event Server in add_data_to_json() at messageFilter.py\n {e}  ')
            pass

    def time(self):
        try:
            zuri = pytz.timezone("Europe/Zurich")
            locale.setlocale(locale.LC_TIME, "de_CH")
            current_time = datetime.now(zuri)
            time_format = "%H:%M:%S"
            the_time = current_time.strftime(time_format)
            return the_time
        except Exception as e:
            send_email(subject='Error',message=f'error in Pzm Event Server in time() at messageFilter.py\n {e}  ')
            pass
            




