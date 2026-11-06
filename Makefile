.PHONY: build build-mcp build-all test lint install install-mcp clean

BIN_EXT := $(if $(filter windows,$(shell go env GOOS)),.exe,)

build:
	go build -trimpath -o bin/outlook-web-pp-cli$(BIN_EXT) ./cmd/outlook-web-pp-cli

build-mcp:
	go build -trimpath -o bin/outlook-web-pp-mcp$(BIN_EXT) ./cmd/outlook-web-pp-mcp

build-all: build build-mcp

test:
	go test ./...

lint:
	go vet ./...

install:
	go install ./cmd/outlook-web-pp-cli

install-mcp:
	go install ./cmd/outlook-web-pp-mcp

clean:
	rm -rf bin/
