#!/usr/bin/env python
import asyncio
import json
import pytz
import websockets
import locale
from messageFilter import Message_manager
from payloadCollection import PayloadCollection
from taskScheduler import send_emailWithPdf_tasks, virtual_openClose_tasks
from emailManager import send_email


swiss_timezone = pytz.timezone("Europe/Zurich")
locale.setlocale(locale.LC_TIME, "de_CH")


async def listen():
    uri = PayloadCollection.webServer_Url
    

    while True:
        try:
            async with websockets.connect(uri) as websocket:
                await websocket.send(json.dumps(PayloadCollection.message))
                while True:
                    message = json.loads(await websocket.recv())
                    messageInstance = Message_manager()
                    await messageInstance.message_filter(message=message)
        except Exception as e:
         #   print(f"An error of type {type(e).__name__} occurred while connection: {e}")
            send_email(subject='Error',message=f'there is Exception in listen()  main.py\n{e}  ')
            pass
        await asyncio.sleep(3)


async def main():
    await asyncio.gather(listen(), virtual_openClose_tasks(), send_emailWithPdf_tasks())


if __name__ == "__main__":
    asyncio.run(main())
