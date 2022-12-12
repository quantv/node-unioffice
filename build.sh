go build -buildmode=c-shared -ldflags="-s -w" -o "include/libspreadsheet.so" src-go/lib.go
#ldflags="-s -w"
npm install