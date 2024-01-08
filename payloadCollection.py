import json
import os


class PayloadCollection:
    username = os.environ.get("GLUTZ_BST_USER")
    password = os.environ.get("GLUTZ_BST_PASS")
    email_user = os.environ.get("EMAIL_BST_USER")
    email_pass = os.environ.get("EMAIL_BST_PASS")
    headers = {"Content-Type": "application/json"}
    GlutzUrl = "31.24.10.138"
    # RpcServerUrl = f"http://{username}:{password}@{GlutzUrl}:8331/rpc/" # there is  no license
    webServer_Url = f"ws://{username}:{password}@{GlutzUrl}:8331"
    IO_Module_Type = 103
    E_Reader_IP55_Type = 102
    E_Reader_Type = 101
    IO_Extender_Type = 80
    IO_ModuleRelay_1 = 2
    IO_ModuleRelay_2 = 4

    doorDict = [
        {
            "PZM-KPA Haus 25": {
                "Hauptzugang": "556.523.178",
                "number": "25.005-2",
                "floor": "Erdgeschoss",
            }
        },
        {
            "PZM-KPA Haus 24": {
                "Hauptzugang": "560.456.939",
                "number": "24.435-8",
                "floor": "Obergeschoss",
            }
        },
    ]
    german_months = [
        "Januar",
        "Februar",
        "März",
        "April",
        "Mai",
        "Juni",
        "Juli",
        "August",
        "September",
        "Oktober",
        "November",
        "Dezember",
    ]
    weekTable_first_row = [
        "Totalisierung",
        "Montag",
        "Dienstag",
        "Mittwoch",
        "Donnerstag",
        "Freitag",
        "Samstag",
        "Sonntag",
    ]

    message = {
        "method": "registerObserver",
        "params": [
            [
                "Media",
                "Devices",
                "AccessPoints",
                "AuthorizationPoints",
                "AuthorizationPointRelations",
                "DeviceEvents",
                "Rights",
                "ObservedStates",
                "DeviceStatus",
                "RouteTree",
                "Properties",
                "PropertyValueSpecs",
                "DevicePropertyData",
                "DeviceStaticPropertyData",
                "SystemPropertyData",
                "SubsystemPropertyData",
                "AccessPointPropertyData",
                "UserPropertyData",
                "DeviceUpdates",
            ]
        ],
        "jsonrpc": "2.0",
    }

    def open_door(deviceid):
        return {
            "deviceid": deviceid,
            "events": [{"condition": "Rising Edge", "event": "Input 1"}],
            "noEventsLost": True,
            "time": "2023-10-12T16:27:59Z",
            "type": "deviceEvent",
        }

    def close_door(deviceid):
        return {
            "deviceid": deviceid,
            "events": [{"condition": "Falling Edge", "event": "Input 1"}],
            "noEventsLost": True,
            "time": "2023-10-12T16:28:05Z",
            "type": "deviceEvent",
        }
    pdf_mail_text = '''Dies ist eine automatisch generierte Auswertung der Türen. Im Anhang befindet sich das entsprechende PDF.\nBei Rückfragen zu diesem Mail oder wenn Sie dieses Mail nicht mehr wünschen, melden Sie sich bitte bei der internen Sicherheitsabteilung.\nSchliessanlage@pzmag.ch'''