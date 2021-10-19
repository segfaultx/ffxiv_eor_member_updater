#!/usr/bin/python3

import openpyxl
import requests

BASE_URL_XIV_API_CHARACTER: str = "https://xivapi.com/character/"
GERMAN_TO_ENGLISH_CLASS_DICT: dict = {"Paladin": "Paladin",
                                      "Krieger": "Warrior",
                                      "Dunkelritter": "Dark Knight",
                                      "Revolverklinge": "Gunbreaker",
                                      "Weißmagier": "White Mage",
                                      "Gelehrter": "Scholar",
                                      "Astrologe": "Astrologian",
                                      "Mönch": "Monk",
                                      "Dragoon": "Dragoon",
                                      "Ninja": "Ninja",
                                      "Samurai": "Samurai",
                                      "Barde": "Bard",
                                      "Maschinist": "Machinist",
                                      "Tänzer": "Dancer",
                                      "Schwarzmagier": "Black Mage",
                                      "Beschwörer": "Summoner",
                                      "Rotmagier": "Red Mage",
                                      "Blaumagier": "Blue Mage",
                                      "Zimmerer": "Carpeter",
                                      "Grobschmied": "Blacksmith",
                                      "Plattner": "Armorsmith",
                                      "Goldschmied": "Goldsmith",
                                      "Gerber": ":Leatherworker",
                                      "Weber": "Weaver",
                                      "Alchemist": "Alchemist",
                                      "Mienenarbeiter": "Miner",
                                      "Gärtner": "Botanist",
                                      "Fischer": "Fisher"}


def main():
    """main method, used to process data and update the excel workbook"""

    workbook: openpyxl.Workbook = openpyxl.load_workbook("/home/amatus/Downloads/Mitgliederliste.xlsx")
    worksheet = workbook.active
    min_row: tuple = worksheet[worksheet.min_row + 1]

    for i in range(worksheet.min_row + 1, worksheet.max_row):
        current_row: tuple = worksheet[i]
        current_character_name: str = f"{current_row[0].value} {current_row[1].value}"
        current_character_info: dict = get_character_info(get_character_id(current_character_name))
        if not current_character_info:
            print(f"Cant process data for character: {current_character_name}")
            continue
        # print(current_character_info)


def do_http_get(request_url: str):
    """helper method to do http requests"""
    resp: requests.Response = requests.get(request_url)
    if resp.ok:
        return resp.json()
    else:
        raise ConnectionError


def get_character_info(character_id: str):
    """helper method to receive character info via XIV API"""
    if not character_id:
        return None
    current_request_url: str = f"{BASE_URL_XIV_API_CHARACTER}{character_id}"
    resp_json: dict = do_http_get(current_request_url)
    return resp_json


def get_character_id(character_name: str):
    """Help method to get the ID of an character via XIV API"""
    current_request_url: str = f"{BASE_URL_XIV_API_CHARACTER}search?name={character_name}&server=Moogle"
    resp_json: dict = do_http_get(current_request_url)
    return resp_json["Results"][0]["ID"] if resp_json["Results"] else None


if __name__ == '__main__':
    main()
