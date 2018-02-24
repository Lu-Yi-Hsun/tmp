GOOS=windows GOARCH=amd64 go build -o window電腦.exe main.go
GOOS=darwin GOARCH=amd64 go build -o 蘋果電腦.App main.go
go build -o Linux系統.a main.go
