package main

import (
	"errors"
	"fmt"
	"github.com/karashiiro/bingode"
	"github.com/xivapi/godestone/v2"
	"log"
	"time"
)

var (
	Scraper = godestone.NewScraper(bingode.New(), godestone.EN)
)

func GetCharacterId(characterName string) (uint32, error) {
	searchOptions := godestone.CharacterOptions{
		Name:  characterName,
		World: World,
		DC:    DataCenter,
	}

	if DEBUG {
		log.Printf("Fetching Data for: %s\n", characterName)
	}

	result := Scraper.SearchCharacters(searchOptions)
	timeout := time.After(time.Second * 30)

	for {
		select {
		case character := <-result:
			{
				if character.Error != nil {
					log.Fatalln(character.Error)
				}
				return character.ID, nil
			}
		case <-timeout:
			return 0, errors.New(fmt.Sprintf("No Character with Name " + characterName + "%s found"))
		}
	}
}

func GetCharacterInfo(characterId uint32) *godestone.Character {
	result, err := Scraper.FetchCharacter(characterId)

	if err != nil {
		log.Fatalln(err)
	}

	return result
}
