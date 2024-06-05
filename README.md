# sapi-go
A simple Go SAPI TTS wrapper

## Usage
After initializing your module, get the library with:
```
go get github.com/Edw590/sapi-go
```
```go
package main

import (
  	"github.com/Edw590/sapi-go"
  	"log"
)

func main() {
    tts, err := sapi.NewSapi()
    if err != nil {
        log.Fatal(err)
    }

    tts.Say("Hello world!")
}
```
