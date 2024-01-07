package main

import (
	"bufio"
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
)

type Config struct {
	ClassConfig      map[string]string `json:"class_config"`
	Sub30ClassConfig map[string]string `json:"sub_30_class_config"`
	World            string            `json:"world"`
	DataCenter       string            `json:"data_center"`
}

var (
	GermanToEnglishClassMap = make(map[string]string)
	Sub30MappingMap         = make(map[string]string)
	World                   string
	DataCenter              string
	DEBUG                   = false
)

func main() {
	filename := flag.String("filename", "./Mitgliederliste.xlsx", "Path to file to process")
	config := flag.String("config", "./eor_config.json", "Path to config location")
	flag.BoolVar(&DEBUG, "debug", false, "enable debug mode")
	flag.Parse()

	loadConfig(*config)
	ProcessExcel(*filename)
	fmt.Println("Updating member data finished, press any key to close")
	reader := bufio.NewReader(os.Stdin)
	_, _, _ = reader.ReadRune()
}

func loadConfig(configLocation string) {
	var config = Config{}
	data, err := os.ReadFile(configLocation)

	if err != nil {
		log.Fatalln(err)
	}

	err = json.Unmarshal(data, &config)

	if err != nil {
		log.Fatalln(err)
	}

	if config.Sub30ClassConfig == nil || config.ClassConfig == nil || len(config.World) == 0 || len(config.DataCenter) == 0 {
		log.Fatalln("No proper config for classes, sub 30 classes, data center or world present!")
	}

	GermanToEnglishClassMap = config.ClassConfig
	Sub30MappingMap = config.Sub30ClassConfig
	World = config.World
	DataCenter = config.DataCenter
}
