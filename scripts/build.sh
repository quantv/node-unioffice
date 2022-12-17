#!/bin/bash

set -e

go build -buildmode=c-shared -ldflags="-s -w" -o "include/libspreadsheet.so" src-go/lib.go src-go/funcs.go

#go build -buildmode=c-shared -o "include/libspreadsheet.so" src-go/lib.go
#ldflags="-s -w"
npm install