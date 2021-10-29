from photoshop import Session
import challonge
import datetime
import re
import operator
import win32com.client
import os
import photoshop.api as ps
import configparser


configParser = configparser.RawConfigParser()   
configFilePath = r'./config.ini'
configParser.read(configFilePath)


manual_input = input("Would you like to manually input starter data? (y/n): ")
if (manual_input.lower() == "y"):
    tournament_type = input("Input tournament ID code (1 for Melee, 2 for Ult, 3 for P+, 4 for Melee (Alt)): ")
    tournament_number = input("Input tournament number (start with #): ")
    name = tournament_number
    date = input("Input date: ")
    entrants = input("Input entrants (just the number): ")
else:
    challonge.set_credentials(configParser.get('config-variables', 'username'), configParser.get('config-variables', 'token'))
    tournament_id = input("Input tournament ID code: ")
    tournament = (challonge.tournaments.show(tournament_id))

    if "Clear" in tournament_id.lower() or "Remove" in tournament_id.lower() or tournament_id == "5":
        tournament_type = "5"
    else:
        if "Melee" in tournament['game_name']:
            tournament_type = input("Melee (1 for Regular, 4 for Alternative): ")
        elif ("Ultimate" in tournament['game_name']):
            tournament_type = "2"
        elif ("Project" in tournament['game_name']):
            tournament_type = "3"
        else:
            tournament_type = input("1 for Melee, 2 for Ultimate, 3 for P+: ")

    if (tournament_type == "1" or tournament_type == "2" or tournament_type == "3" or tournament_type == "4"):
        tournament = (challonge.tournaments.show(tournament_id))
        tournament_number = (tournament["name"])
        name = tournament["name"]
        print("Grabbing tournament: " + name)
        name = name[name.find('#'):]
        if (tournament_type == "1" or tournament_type == "4"):
            name = "(" + name + ")"
        elif (tournament_type == "3"):
            name = name[ 0 : 4]
        entrants = str(tournament["participants_count"])
        date = tournament["started_at"]
        if (tournament_type == "1" or tournament_type == "4"):
            date = date.strftime("%m / %d / %Y")
        elif (tournament_type == "2"):
            date = date.strftime("%m/%d/%Y")
        else:
            date = date.strftime("%b. %d, %Y")
    print(name)
    print(entrants + " Entrants")
    print(date)
manual_name = input("Would you like to manually input names? (y/n): ")
if (manual_name.lower() == "y"):
    standings = []
    standings1 = input("1st Place Input: ")
    standings2 = input("2nd Place Input: ")
    standings3 = input("3rd Place Input: ")
    standings4 = input("4th Place Input: ")
    standings5 = input("5th Place (1) Input: ")
    standings6 = input("5th Place (2) Input: ")
    standings7 = input("7st Place (1) Input: ")
    standings8 = input("7st Place (2) Input: ")
    standings.append(standings1)
    standings.append(standings2)
    standings.append(standings3)
    standings.append(standings4)
    standings.append(standings5)
    standings.append(standings6)
    standings.append(standings7)
    standings.append(standings8)

    if (tournament_type == "1"):
        standings9 = input("9th Place (1) Input: ")
        standings10 = input("9th Place (2) Input: ")
        standings11 = input("9th Place (3) Input: ")
        standings12 = input("9th Place (4) Input: ")
        standings.append(standings9)
        standings.append(standings10)
        standings.append(standings11)
        standings.append(standings12)
else:
    def get_rank(rank):
        return rank.get('final_rank')
    def get_name(rank):
        return rank.get('name')
    unsorted = challonge.participants.index(tournament_id)
    unsorted.sort(key=get_name)
    unsorted.sort(key=get_rank)
    standings = []
    for d in unsorted:
        standings.append(d['name'])
    if (tournament_type == "1"):
        finalmax = 13
    else:
        finalmax = 9
    if (len(standings) >= finalmax):
        standings = standings[:(finalmax - 1)]
    else:
        while(len(standings) < finalmax - 1):
            standings.append("--")
if (tournament_type == "3"):
    valsLower = [item.upper() for item in standings]
    standings = []
    standings = valsLower
elif (tournament_type == "4"):
    valsUpper = [item.upper() for item in standings]
    standings = []
    standings = valsUpper

psApp = win32com.client.Dispatch("Photoshop.Application")
if tournament_type == "1":
    psApp.Open(configParser.get('config-variables', 'melee_path'))
    textlayers = ["1st Place", "2nd Place", "3rd Place", "4th Place", "5th Place(#1)", "5th Place (#2)",
    "7th Place (#1)","7th Place (#2)", "9th Place (#1)","9th Place (#2)", "9th Place (#3)","9th Place (#4)"]
elif tournament_type == "4":
    psApp.Open(configParser.get('config-variables', 'melee_alt_path'))
    textlayers = ["1st Place", "2nd Place", "3rd Place", "4th Place", "5th Place(#1)", "5th Place (#2)",
    "7th Place (#1)","7th Place (#2)"]
else:
    if tournament_type == "2":
        psApp.Open(configParser.get('config-variables', 'ult_path'))
    else:
        psApp.Open(configParser.get('config-variables', 'pplus_path'))
    textlayers = ["1st Place", "2nd Place", "3rd Place", "4th Place", "5th Place(#1)", "5th Place (#2)",
    "7th Place (#1)","7th Place (#2)"]
print("Successfully opened Photoshop document")
docRef = psApp.Application.ActiveDocument
#nLayerSets = len(list((i, x) for i, x in enumerate(docRef.layerSets))) - 1
#nArtLayers = len(
    #list((i, x) for i, x in enumerate(docRef.layerSets[nLayerSets].artLayers)),
#)

active_layer = docRef.activeLayer = docRef.layerSets[1].artLayers[0]
print(active_layer.name)

#Text Edit
i = 0
for t in textlayers:
    active_layer = docRef.activeLayer = docRef.layerSets[0].artLayers[i]
    layer_text = active_layer.TextItem
    layer_text.contents = standings[i]
    print("adding " + standings[i] + " to " + docRef.layerSets[0].artLayers[i].name)
    i = i + 1

number = tournament_number[name.find('#'):]
if tournament_type == "3":
    editable_list = [name.upper(), date.upper(), entrants + " PLAYERS"]
else:
    editable_list = [name, date, entrants + " Entrants"]

i = 0
for e in editable_list:
    active_layer = docRef.activeLayer = docRef.layerSets[1].artLayers[i]
    layer_text = active_layer.TextItem
    layer_text.contents = editable_list[i]
    print("adding " + editable_list[i] + " to " + docRef.layerSets[1].artLayers[i].name)
    i = i + 1

if (tournament_type == "1" or tournament_type == "2"):
    i = 0
    if (tournament_type == "1"):
        loop_number = len(docRef.layerSets[2].layersets)
        while i < loop_number:
            try:
                print (len(docRef.layerSets[2].layersets))
                print (docRef.layerSets[2].layersets[i].name)
                print (len(docRef.layerSets[2].layersets[i].artLayers))
                print ("Attempting to delete " + docRef.layerSets[2].layersets[i].name)
                docRef.layerSets[2].layersets[i].artLayers.removeAll()
                i = i + 1
            except Exception:
                i = i + 1
                pass
    else:
        loop_number = len(docRef.layerSets[5].layersets)
        try:
            docRef.layerSets[3].artLayers.removeAll()
        except Exception:
            pass
        i = 0
        while i < loop_number:
            try:
                print (len(docRef.layerSets[5].layersets))
                print (docRef.layerSets[5].layersets[i].name)
                print (len(docRef.layerSets[5].layersets[i].artLayers))
                print ("Attempting to delete " + docRef.layerSets[5].layersets[i].name)
                docRef.layerSets[5].layersets[i].artLayers.removeAll()
                i = i + 1
            except Exception:
                i = i + 1
                pass
