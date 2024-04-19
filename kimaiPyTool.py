#!/usr/bin/python3
# PYTHON_ARGCOMPLETE_OK

import argparse
import argcomplete
import os
import pathlib
import json
import sys
import typing
import requests
import datetime
import types
import locale
import dataclasses
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


JsonValue = typing.Union[None, int, str, bool, list["JsonValue"], "JsonObject"]
JsonObject = dict[str, JsonValue]
JsonList = list[JsonValue]
RequestParam = dict[str, str]

APP_NAME = "kimaiPyTool"
KIMAI_TAG_FOR_GOOGLE_CALENDAR = "gcalendar"
GOOGLE_API_SCOPES = ['https://www.googleapis.com/auth/calendar']
GOOGLE_CLIENT_SECRET_FILE = 'google_api_secret.json'
GOOGLE_TOKEN_FILE = 'google_token.json'


T = typing.TypeVar("T")
def jsonObject2NamedTuple(cls: type[T], jsonObject: JsonObject) -> T:
    if type(jsonObject) is not dict:
        raise TypeError('Unexpected type for {}, expected {}, get {}'
                .format(jsonObject, dict, type(jsonObject)))
    d:dict[str, typing.Any] = dict()
    for fieldName, fieldType in cls.__annotations__.items():
        if fieldName not in jsonObject:
            if (typing.get_origin(fieldType) in [typing.Union, types.UnionType]
                    and isinstance(None, typing.get_args(fieldType))):
                d[fieldName] = None
                continue
            raise ValueError("{} not found in {}".format(fieldName, jsonObject))
        typeOk = False
        if typing.get_origin(fieldType) is None:
            typeOk = isinstance(jsonObject[fieldName], fieldType)
        elif typing.get_origin(fieldType) in [typing.Union, types.UnionType]:
            typeOk = isinstance(jsonObject[fieldName], typing.get_args(fieldType))
        elif typing.get_origin(fieldType) is list and len(typing.get_args(fieldType)) == 1:
            typeOk = (type(jsonObject[fieldName]) is list and
                all(isinstance(x, typing.get_args(fieldType)[0]) for x in jsonObject[fieldName])) # type: ignore
        else:
            raise NotImplementedError("The type {} is not suported yet".format(typing.get_origin(fieldType)))
        if not typeOk:
            if jsonObject[fieldName] is None:
                print("The field {} cannot be None in {}".format(fieldName, jsonObject),
                      file=sys.stderr)
                sys.exit(1)
            raise TypeError('Unexpected type for field "{}" in {}, expected {}, get {}'
                    .format(fieldName, jsonObject, fieldType, type(jsonObject[fieldName])))
        d[fieldName] = jsonObject[fieldName]
    return cls(**d)


class Kimai:
    def __init__(self, url: str, username:str, password: str):
        self._url = url
        self._username = username
        self._password = password

    def _buildheader(self):
        return {"X-AUTH-USER": self._username, "X-AUTH-TOKEN": self._password}

    def _runRequest(self, method:str, urlSufix: str, params:RequestParam|None={}, data: JsonObject|None=None):
        response = requests.request(method, self._url + "/" + urlSufix, params=params, headers=self._buildheader(), data=data)
        if response.status_code != 200:
            print('For request "{}" get {}'.format(response.url, response.json()), file=sys.stderr)
            sys.exit(1)
        if response.headers.get("X-Total-Pages", "1") != "1":
            print('For request "{}" get too much result: {}'.format(response.url, response.headers["X-Total-Count"]), file=sys.stderr)
            sys.exit(1)
        return response.json()

    def getCustomers(self) -> JsonList:
        return self._runRequest("get", "customers")

    def getProjects(self) -> JsonList:
        return self._runRequest("get", "projects")

    def getActivities(self) -> JsonList:
        return self._runRequest("get", "activities")

    def getActivity(self, id: int) -> JsonObject:
        return self._runRequest("get", "activities/{}".format(id))

    def updateActivity(self, id: int, data: JsonObject) -> JsonObject:
        return self._runRequest("patch", "activities/{}".format(id), data=data)

    def getTimesheets(self, begin:str|None=None, maxItem: int|None=None) -> JsonList:
        params:RequestParam = RequestParam()
        if begin is not None:
            params["begin"] = begin
        if maxItem is not None:
            params["size"] = str(maxItem)
        return self._runRequest("get", "timesheets", params)

    def addTimesheet(self, userId: int, projectId: int, activityId: int, begin: str, end: str,
            description: str) -> JsonObject:
        data = JsonObject()
        data["user"] = userId
        data["project"] = projectId
        data["activity"] = activityId
        data["begin"] = begin
        data["end"] = end
        data["description"] = description
        return self._runRequest("post", "timesheets", data=data)

    def addTimesheetTag(self, timeSheetId: int, tags: list[str]) -> JsonObject:
        data = JsonObject()
        data["tags"] = typing.cast(JsonList, tags)
        return self._runRequest("patch", "timesheets/{}".format(timeSheetId), data=data)


class KimaiCustomer(typing.NamedTuple):
    id: int
    name: str
    number: str
    comment: str
    visible: bool
    billable: bool
    currency: str


class KimaiCustomers:
    def __init__(self, jsonList: JsonList):
        self._customersById: dict[int, KimaiCustomer] = dict()
        self._idsByName: dict[str, int] = dict()
        for jsonValue in jsonList:
            if type(jsonValue) is not typing.get_origin(JsonObject):
                raise TypeError('Unexpected type for {}, expected {}, get {}'
                        .format(jsonValue, JsonObject, type(jsonValue)))
            jsonObject = typing.cast(JsonObject, jsonValue)
            customer = jsonObject2NamedTuple(KimaiCustomer, jsonObject)
            self._customersById[customer.id] = customer
            self._idsByName[customer.name] = customer.id

    def get(self, id: int) -> KimaiCustomer:
        return self._customersById[id]

    def getIdByName(self, name: str) -> int:
        return self._idsByName[name]


class KimaiProject(typing.NamedTuple):
    parentTitle: str
    customer: int
    id: int
    name: str
    start: str|None
    end: str|None
    comment: str|None
    visible: bool
    billable: bool


class KimaiProjects:
    def __init__(self, jsonList: JsonList):
        self._projectsById: dict[int, KimaiProject] = dict()
        self._idsByName: dict[str, int] = dict()
        self._idsByCustomerId: dict[int, list[int]] = dict()
        for jsonValue in jsonList:
            if type(jsonValue) is not typing.get_origin(JsonObject):
                raise TypeError('Unexpected type for {}, expected {}, get {}'
                        .format(jsonValue, JsonObject, type(jsonValue)))
            jsonObject = typing.cast(JsonObject, jsonValue)
            project = jsonObject2NamedTuple(KimaiProject, jsonObject)
            self._projectsById[project.id] = project
            if project.name in self._idsByName:
                raise ValueError('The project name "{}" already exist'.format(project.name))
            self._idsByName[project.name] = project.id
            if project.customer not in self._idsByCustomerId:
                self._idsByCustomerId[project.customer] = list()
            self._idsByCustomerId[project.customer].append(project.id)

    def get(self, id: int) -> KimaiProject:
        return self._projectsById[id]

    def getIdByName(self, name: str) -> int:
        return self._idsByName[name]

    def getIdsByCustomerId(self, customerId: int) -> list[int]:
        return self._idsByCustomerId[customerId]


class KimaiActivity(typing.NamedTuple):
    parentTitle: str
    project: int
    id: int
    name: str
    comment: str|None
    visible: bool
    billable: bool


class KimaiActivities:
    def __init__(self, jsonList: JsonList):
        self._activitiesById: dict[int, KimaiActivity] = dict()
        self._idsByName: dict[str, int] = dict()
        self._idsByProjectId: dict[int, list[int]] = dict()
        for jsonValue in jsonList:
            if type(jsonValue) is not typing.get_origin(JsonObject):
                raise TypeError('Unexpected type for {}, expected {}, get {}'
                        .format(jsonValue, JsonObject, type(jsonValue)))
            jsonObject = typing.cast(JsonObject, jsonValue)
            activity = jsonObject2NamedTuple(KimaiActivity, jsonObject)
            self._activitiesById[activity.id] = activity
            if activity.name in self._idsByName:
                raise ValueError('The activity name "{}" already exist'.format(activity.name))
            self._idsByName[activity.name] = activity.id
            if activity.project not in self._idsByProjectId:
                self._idsByProjectId[activity.project] = list()
            self._idsByProjectId[activity.project].append(activity.id)

    def get(self, id: int) -> KimaiActivity:
        return self._activitiesById[id]

    def getIdByName(self, name: str) -> int:
        return self._idsByName[name]

    def getIdsByProjectId(self, projectId: int) -> list[int]:
        return self._idsByProjectId[projectId]


class KimaiTimeSheet(typing.NamedTuple):
    activity: int
    project: int
    user: int
    id: int
    begin: str
    end: str
    duration: int
    description: str
    exported: bool
    billable: bool
    tags: list[str]


class KimaiTimeSheets:
    def __init__(self, jsonList: JsonList):
        self._timeSheetsById: dict[int, KimaiTimeSheet] = dict()
        for jsonValue in jsonList:
            if type(jsonValue) is not typing.get_origin(JsonObject):
                raise TypeError('Unexpected type for {}, expected {}, get {}'
                        .format(jsonValue, JsonObject, type(jsonValue)))
            jsonObject = typing.cast(JsonObject, jsonValue)
            timeSheet = jsonObject2NamedTuple(KimaiTimeSheet, jsonObject)
            self._timeSheetsById[timeSheet.id] = timeSheet

    def get(self, id: int) -> KimaiTimeSheet:
        return self._timeSheetsById[id]

    def values(self):
        return self._timeSheetsById.values()


def getConfigPath(fileName: str) -> str:
    configPath = os.path.join(pathlib.Path.home(), ".config", APP_NAME, fileName)
    if not os.path.exists(os.path.dirname(configPath)):
        os.mkdir(os.path.dirname(configPath))
    return configPath


@dataclasses.dataclass
class Config:
    kimaiUrl: str|None = None
    kimaiUsername: str|None = None
    kimaiToken: str|None = None
    gCalendarEmail: str|None = None

    def toJson(self) -> JsonObject:
        toReturn: JsonObject = dict()
        if self.kimaiUrl is not None:
            toReturn["kimaiUrl"] = self.kimaiUrl
        if self.kimaiUsername is not None:
            toReturn["kimaiUsername"] = self.kimaiUsername
        if self.kimaiToken is not None:
            toReturn["kimaiToken"] = self.kimaiToken
        if self.gCalendarEmail is not None:
            toReturn["gCalendarEmail"] = self.gCalendarEmail
        return toReturn


def googleApiGetCredentials(secretFilePath: str, tokenFilePath: str):
    credentials = None
    if os.path.exists(tokenFilePath):
        credentials = Credentials.from_authorized_user_file(tokenFilePath, GOOGLE_API_SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(secretFilePath, GOOGLE_API_SCOPES)
            credentials = flow.run_local_server(port=0)
            with open(tokenFilePath, "w") as token:
                token.write(credentials.to_json())
    return credentials


class GCalendarEvent(typing.NamedTuple):
    summary: str
    start: str
    end: str
    description: str|None

    @classmethod
    def fromKimaiTimeSheet(cls, timeSheet: KimaiTimeSheet, clientName: str, projectName: str,
            activityName: str):
        return cls(" - ".join([clientName, projectName, activityName]), timeSheet.begin,
                timeSheet.end, timeSheet.description)

    def toJson(self) -> JsonObject:
        toReturn: JsonObject = dict()
        toReturn["summary"] = self.summary
        toReturn["start"] = dict()
        toReturn["start"]["dateTime"] = self.start
        toReturn["end"] = dict()
        toReturn["end"]["dateTime"] = self.end
        if self.description is not None:
            toReturn["description"] = self.description
        return toReturn


def googleApiPushEventToCalendar(event: GCalendarEvent, calendarEmail: str, service):
    try:
        event = service.events().insert(calendarId=calendarEmail, body=event.toJson()).execute()
        print('Event created: {}'.format(event.get('htmlLink')))
    except HttpError as err:
        message = str(err)
        if err.resp.get('content-type', '').startswith('application/json'):
            message = json.loads(err.content).get('error').get('errors')[0].get('message')
        print('Error while pushing event "{}" at {}: "{}"'.format(event.summary, event.start, message))

def importEventFile(timesheetsFilePath: str, kimaiUserId: int, kimai: Kimai):
        with open(timesheetsFilePath, 'r') as eventsFile:
            eventsData = json.loads(eventsFile.read())
        projects = KimaiProjects(kimai.getProjects())
        activities = KimaiActivities(kimai.getActivities())
        for eventData in eventsData:
            projectName = eventData["projectName"]
            projectId = projects.getIdByName(projectName)
            activityId = None
            if "activityName" in eventData:
                activityName = eventData["activityName"]
                activityId = activities.getIdByName(activityName)
            else:
                activitiesIds = activities.getIdsByProjectId(projectId)
                if len(activitiesIds) != 1:
                    print("For project {} we have no or several activities {}".format(projectName),
                            file=sys.stderr)
                    sys.exit(1)
                activityId = activitiesIds[0]
            timeSheet = kimai.addTimesheet(kimaiUserId, projectId, activityId,
                    eventData["begin"], eventData["end"], eventData["description"])
            print(timeSheet)


@dataclasses.dataclass
class CraItem:
    duration: int = 0
    description: set[str] = dataclasses.field(default_factory=set)


def kimaiToGCalendar(begin: datetime.datetime, kimai: Kimai, gCalendarEmail: str):
    beginStr = datetime.datetime.isoformat(begin)
    timeSheets = KimaiTimeSheets(kimai.getTimesheets(begin=beginStr))
    googleCalendarService = None
    customers = KimaiCustomers(kimai.getCustomers())
    projects = KimaiProjects(kimai.getProjects())
    activities = KimaiActivities(kimai.getActivities())
    for timeSheet in timeSheets.values():
        if KIMAI_TAG_FOR_GOOGLE_CALENDAR not in timeSheet.tags:
            if googleCalendarService is None:
                if not os.path.exists(getConfigPath(GOOGLE_CLIENT_SECRET_FILE)):
                    print("You pust generate a google API OAuth 2.0 token file from "
                            "https://developers.google.com/google-apps/calendar/quickstart/python#prerequisites"
                            " and copy it in {}".format(getConfigPath(GOOGLE_CLIENT_SECRET_FILE)),
                            file=sys.stderr)
                    sys.exit(1)
                googleCredentials = googleApiGetCredentials(getConfigPath(
                    GOOGLE_CLIENT_SECRET_FILE), getConfigPath(GOOGLE_TOKEN_FILE))
                googleCalendarService = build("calendar", "v3", credentials=googleCredentials)
            project = projects.get(timeSheet.project)
            googleCalendarEvent = GCalendarEvent.fromKimaiTimeSheet(timeSheet, customers.get(project.customer).name,
                    project.name, activities.get(timeSheet.activity).name)
            googleApiPushEventToCalendar(googleCalendarEvent, gCalendarEmail,
                    googleCalendarService)
            tags = timeSheet.tags
            tags.append(KIMAI_TAG_FOR_GOOGLE_CALENDAR)
            kimai.addTimesheetTag(timeSheet.id, tags)


def generateCraFiles(begin: datetime.datetime, kimai: Kimai):
    beginStr = datetime.datetime.isoformat(begin)
    end = datetime.date.today()
    timeSheets = KimaiTimeSheets(kimai.getTimesheets(begin=beginStr, maxItem=100))
    customers = KimaiCustomers(kimai.getCustomers())
    projects = KimaiProjects(kimai.getProjects())
    activities = KimaiActivities(kimai.getActivities())
    craByCustomerDateProjectActivity: dict[str, dict[datetime.date, dict[str, dict[str, CraItem]]]] = dict()
    locale.setlocale(locale.LC_ALL, locale.getlocale())
    dateFormat = locale.nl_langinfo(locale.D_FMT)
    activitiesByCustomerProject: dict[str, dict[str, set[str]]] = dict()
    for timeSheet in timeSheets.values():
        # Get time sheet data
        date = datetime.datetime.fromisoformat(timeSheet.begin).date()
        project = projects.get(timeSheet.project)
        customerName = customers.get(project.customer).name
        activityName = activities.get(timeSheet.activity).name
        # Add to CRA
        if customerName not in craByCustomerDateProjectActivity:
            craByCustomerDateProjectActivity[customerName] = dict()
        if date not in craByCustomerDateProjectActivity[customerName]:
            craByCustomerDateProjectActivity[customerName][date] = dict()
        if project.name not in craByCustomerDateProjectActivity[customerName][date]:
            craByCustomerDateProjectActivity[customerName][date][project.name] = dict()
        if activityName not in craByCustomerDateProjectActivity[customerName][date][project.name]:
            craByCustomerDateProjectActivity[customerName][date][project.name][activityName] = CraItem()
        craByCustomerDateProjectActivity[customerName][date][project.name][activityName].duration += timeSheet.duration
        for descriptionLine in timeSheet.description.replace("\r", "").split("\n"):
            craByCustomerDateProjectActivity[customerName][date][project.name][activityName].description.add(descriptionLine)
        # Add to activitiesByCustomerProject
        if customerName not in activitiesByCustomerProject:
            activitiesByCustomerProject[customerName] = dict()
        if project.name not in activitiesByCustomerProject[customerName]:
            activitiesByCustomerProject[customerName][project.name] = set()
        activitiesByCustomerProject[customerName][project.name].add(activityName)
    # Write file by customer
    for customerName, craByDateProjectActivity in craByCustomerDateProjectActivity.items():
        with open("{}_CRA_{}.tsv".format(end.strftime("%Y-%m"), customerName), "w") as craFile:
            # write header
            headerLine1 = "\t"
            headerLine2 = "Date\tTotal (hour)"
            for projectName, activities in activitiesByCustomerProject[customerName].items():
                for activityName in activities:
                    headerLine1 += "\t"+projectName
                    headerLine2 += "\t"+activityName
            headerLine2 += "\tDescription"
            craFile.write(headerLine1+"\n")
            craFile.write(headerLine2+"\n")
            # Write a line by date
            for date, craByProjectActivity in sorted(craByDateProjectActivity.items()):
                durationSum = 0.0
                dataLine = ""
                description: set[str] = set()
                for projectName, activities in activitiesByCustomerProject[customerName].items():
                    for activityName in activities:
                        craItem = craByProjectActivity.get(projectName, dict()).get(activityName, CraItem())
                        durationSum += craItem.duration
                        description.update(craItem.description)
                        dataLine += "\t"
                        if craItem.duration != 0:
                            dataLine += locale.str(craItem.duration/3600)
                craFile.write("{}\t{}{}\t{}\n".format(date.strftime(dateFormat),
                        locale.str(durationSum/3600), dataLine, ", ".join(description).replace("\t", " ")))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Kimai cli tool")
    groupAction = parser.add_mutually_exclusive_group(required=True)
    groupAction.add_argument('--configure', action='store_true', help="Save params in config file")
    groupAction.add_argument('--getCustomers', action='store_true',
            help="Display a json list of customers")
    groupAction.add_argument('--getProjects', action='store_true',
            help="Display a json list of projects")
    groupAction.add_argument('--getActivities', action='store_true',
            help="Display a json list of activities")
    groupAction.add_argument('--getActivity', type=int,
            help="Display a json object of the given activity")
    groupAction.add_argument('--updateActivity', type=int,
            help="Update the given activity")
    groupAction.add_argument('--getTimesheets', action='store_true',
            help="Display a json list of timesheets")
    groupAction.add_argument('--setTimesheets', type=str, help="Import events from file")
    groupAction.add_argument('--toGCalendar', type=lambda s: datetime.datetime.strptime(s,
            '%Y-%m-%d'), help="Copy events from the given date in format YYYY-MM-DD not tag {} to "
            "Google calendar".format(KIMAI_TAG_FOR_GOOGLE_CALENDAR))
    groupAction.add_argument('--cra', type=lambda s: datetime.datetime.strptime(s,
            '%Y-%m-%d'), help="Generate a file in current dir for each customer with one line by "
            "day, project and activity with duration in hour and description")
    parser.add_argument("--kimaiUrl", type=str, help="The Kimai url with protocol and /api "
            "exemple http://nas.local:8001/api (can be saved in config file)")
    parser.add_argument("--kimaiUsername", type=str, help="The Kimai username (can be saved in "
            "config file)")
    parser.add_argument("--kimaiToken", type=str, help="The Kimai API token (can be saved in config"
            "file)")
    parser.add_argument("--gCalendarEmail", type=str, help="The email address of the google "
            "calendar (can be saved in config file)")
    parser.add_argument("--kimaiUserId", type=int, help="The user identifier to set timesheets")
    parser.add_argument("--timeBudget", type=float, help="When use --updateActivity update the time "
            "budget with the given time in hour")
    argcomplete.autocomplete(parser)
    args = parser.parse_args()

    configData = Config()
    configPath = getConfigPath(APP_NAME + ".json")
    if os.path.exists(configPath):
        with open(configPath, 'r') as configFile:
            configJson = json.loads(configFile.read())
            configData = jsonObject2NamedTuple(Config, configJson)
    if args.kimaiUrl:
        configData.kimaiUrl = args.kimaiUrl
    if args.kimaiUsername:
        configData.kimaiUsername = args.kimaiUsername
    if args.kimaiToken:
        configData.kimaiToken = args.kimaiToken
    if args.gCalendarEmail:
        configData.gCalendarEmail = args.gCalendarEmail

    if args.configure:
        with open(configPath, 'w') as configFile:
            json.dump(configData.toJson(), configFile, indent=4)
            configFile.write('\n')
        os.chmod(configPath, 0o600)
        sys.exit(0)
    
    if configData.kimaiUrl == None:
        print("kimai URL is not defined", file=sys.stderr)
        sys.exit(1)
    if configData.kimaiUsername == None:
        print("kimai username is not defined", file=sys.stderr)
        sys.exit(1)
    if configData.kimaiToken == None:
        print("kimai password is not defined", file=sys.stderr)
        sys.exit(1)
    kimai = Kimai(configData.kimaiUrl, configData.kimaiUsername, configData.kimaiToken)

    if args.getCustomers:
        json.dump(kimai.getCustomers(), sys.stdout, indent=4)

    if args.getProjects:
        json.dump(kimai.getProjects(), sys.stdout, indent=4)

    if args.getActivities:
        json.dump(kimai.getActivities(), sys.stdout, indent=4)

    if args.getActivity:
        json.dump(kimai.getActivity(args.getActivity), sys.stdout, indent=4)

    if args.updateActivity:
        data = JsonObject()
        if args.timeBudget:
            data["timeBudget"] = args.timeBudget
        if len(data)==0:
            print("You must use at least one argument associate with the command updateActivity",
                    file=sys.stderr)
            sys.exit(1)
        json.dump(kimai.updateActivity(args.updateActivity, data), sys.stdout, indent=4)

    if args.getTimesheets:
        json.dump(kimai.getTimesheets(), sys.stdout, indent=4)

    if args.setTimesheets:
        if not args.kimaiUserId:
            print("You must define the kimai user id to import data with console argument",
                    file=sys.stderr)
            sys.exit(1)
        importEventFile(args.setTimesheets, args.kimaiUserId, kimai)

    if args.toGCalendar:
        if configData.gCalendarEmail is None:
            print("google calendar email addess not defined", file=sys.stderr)
            sys.exit(1)
        kimaiToGCalendar(args.toGCalendar, kimai, configData.gCalendarEmail)

    if args.cra:
        generateCraFiles(args.cra, kimai)
