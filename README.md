# sapi-go
A simple Go Windows SAPI TTS wrapper

Main library gotten from https://github.com/DeepHorizons/tts, so thank you as it was the base for this.

Only supports SAPI.

## Usage
After initializing your module, get the library with:
```
go get github.com/Edw590/sapi-go@latest
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

    tts.SetRate(-5)
    tts.Say("This will be said slower")

    tts.SetVolume(30)
    tts.Say("This will be said on a lower volume")
}
```
