SOURCE=.
APP=excel-sample
VERSION=1.0
ARCH= $(shell uname -m)
GOBASE=$(shell pwd)
RELEASE_DIR=$(GOBASE)/bin

.DEFAULT_GOAL = build 

GO_SRC_DIRS := $(shell \
	find . -name "*.go" -not -path "./vendor/*" | \
	xargs -I {} dirname {}  | \
	uniq)
GO_TEST_DIRS := $(shell \
	find . -name "*_test.go" -not -path "./vendor/*" | \
	xargs -I {} dirname {}  | \
	uniq)	

build: 
	@go build -v -o ${APP} ${SOURCE}

lint:
	@goimports -w ${GO_SRC_DIRS}	
	@gofmt -s -w -d ${GO_SRC_DIRS}
	@golint ${GO_SRC_DIRS}
	@go vet ${GO_SRC_DIRS}

test:
	go test -v ${GO_TEST_DIRS}
	@#go test -race -count 100 ${GO_TEST_DIRS}

bench:
	go test -benchmem -bench=. ${GO_TEST_DIRS}

mod:
	go mod verify
	go mod tidy

run:
	@go run ${SOURCE}


.PHONY: build run release lint test mod