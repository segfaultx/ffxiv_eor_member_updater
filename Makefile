.PHONY: windows

windows:
	GOOS=windows GOARCH=amd64 go build