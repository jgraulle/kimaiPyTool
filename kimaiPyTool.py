#!/usr/bin/python3

import argparse
import os
import pathlib
import json
import sys
import typing
import requests
from dataclasses import dataclass


JsonValue = typing.Union[None, int, str, bool, list["JsonValue"], "JsonObject"]
JsonObject = dict[str, JsonValue]
JsonList = list[JsonValue]


APP_NAME = "kimaiPyTool"


T = typing.TypeVar("T")
def jsonObject2NamedTuple(cls: type[T], jsonObject: JsonObject) -> T:
    if type(jsonObject) is not dict:
        raise TypeError('Unexpected type for {}, expected {}, get {}'
                .format(jsonObject, dict, type(jsonObject)))
    d = dict() # type: ignore
    for fieldName, fieldType in cls.__annotations__.items():
        if fieldName not in jsonObject:
            raise ValueError("{} not found in {}".format(fieldName, jsonObject))
        if not isinstance(jsonObject[fieldName], fieldType):
            raise TypeError('Unexpected type for field {} in {}, expected {}, get {}'
                    .format(fieldName, jsonObject, fieldType, type(jsonObject[fieldName])))
        d[fieldName] = jsonObject[fieldName]
    return cls(**d) # type: ignore


class Kimai:
    def __init__(self, url: str, username:str, password: str):
        self._url = url
        self._username = username
        self._password = password

    def _buildheader(self):
        return {"X-AUTH-USER": self._username, "X-AUTH-TOKEN": self._password}

    def _runGetRequest(self, method: str):
        return requests.get(self._url + "/" + method, headers=self._buildheader())

    def _runPostRequest(self, method: str, data: JsonObject):
        return requests.post(self._url + "/" + method, headers=self._buildheader(), data=data).json()

    def getCustomers(self) -> JsonList:
        return self._runGetRequest("customers").json()

    def getProjects(self) -> JsonList:
        return self._runGetRequest("projects").json()

    def getActivities(self) -> JsonList:
        return self._runGetRequest("activities").json()

    def getTimesheets(self) -> JsonList:
        return self._runGetRequest("timesheets").json()

    def addTimesheet(self, userId: int, projectId: int, activityId: int, begin: str, end: str,
            description: str) -> JsonObject:
        data = JsonObject()
        data["user"] = userId
        data["project"] = projectId
        data["activity"] = activityId
        data["begin"] = begin
        data["end"] = end
        data["description"] = description
        return self._runPostRequest("timesheets", data)


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
    start: None|str
    end: None|str
    comment: None|str
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


def getConfigPath(fileName: str) -> str:
    configPath = os.path.join(pathlib.Path.home(), ".config", APP_NAME, fileName)
    if not os.path.exists(os.path.dirname(configPath)):
        os.mkdir(os.path.dirname(configPath))
    return configPath


@dataclass
class Config:
    kimaiUrl: str|None = None
    kimaiUsername: str|None = None
    kimaiToken: str|None = None


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
    groupAction.add_argument('--getTimesheets', action='store_true',
            help="Display a json list of timesheets")
    groupAction.add_argument('--setTimesheets', type=str, help="Import events from file")
    parser.add_argument("--kimaiUrl", type=str, help="The Kimai url with protocol and /api "
            "exemple http://nas.local:8001/api")
    parser.add_argument("--kimaiUsername", type=str, help="The Kimai username")
    parser.add_argument("--kimaiToken", type=str, help="The Kimai API token")
    parser.add_argument("--kimaiUserId", type=int, help="The user identifier to set timesheets")
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

    if args.configure:
        with open(configPath, 'w') as configFile:
            json.dump(configData, configFile, indent=4)
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
    kiman = Kimai(configData.kimaiUrl, configData.kimaiUsername, configData.kimaiToken)

    if args.getCustomers:
        json.dump(kiman.getCustomers(), sys.stdout, indent=4)

    if args.getProjects:
        json.dump(kiman.getProjects(), sys.stdout, indent=4)

    if args.getActivities:
        json.dump(kiman.getActivities(), sys.stdout, indent=4)

    if args.getTimesheets:
        json.dump(kiman.getTimesheets(), sys.stdout, indent=4)

    if args.setTimesheets:
        if not args.kimaiUserId:
            print("You must define the kimai user id to import data with console argument",
                    file=sys.stderr)
            sys.exit(1)
        with open(args.setTimesheets, 'r') as eventsFile:
            eventsData = json.loads(eventsFile.read())
        customers = KimaiCustomers(kiman.getCustomers())
        projects = KimaiProjects(kiman.getProjects())
        activities = KimaiActivities(kiman.getActivities())
        for eventData in eventsData:
            clientName = eventData["clientName"]
            clientId = customers.getIdByName(clientName)
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
            timeSheet = kiman.addTimesheet(args.kimaiUserId, projectId, activityId,
                    eventData["begin"], eventData["end"], eventData["description"])
            print(timeSheet)
