package main

import (
	"github.com/Edw590/sapi-go"
	"github.com/go-ole/go-ole"
	"time"
)

func main() {
	// Check the full SAPI documentation on https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ee125640(v=vs.85)

	ole.CoInitialize(0)
	tts, _ := sapi.NewSapi()
	tts2, _ := sapi.NewSapi()

	tts.Speak("This will be spoken synchronously and normally", sapi.SVSFDefault)
	tts.Speak("This will be spoken asynchronously and normally for the first 3 seconds and slower for the rest of the time",
		sapi.SVSFlagsAsync)
	time.Sleep(3 * time.Second)
	// Slower because SetRate is called while the speech is being spoken (asynchronous Speak() call)

	tts.SetRate(-5)
	tts.Speak("This will be spoken slower", sapi.SVSFDefault)

	tts.SetRate(0)
	tts.SetVolume(30)
	tts.Speak("This will be spoken quieter", sapi.SVSFDefault)

	tts.Speak("This will be skipped after the first second", sapi.SVSFlagsAsync)

	time.Sleep(1 * time.Second)
	tts.Skip(1)

	tts.WaitUntilDone(10000)

	tts.Speak("This will be interrupted by a higher priority speech", sapi.SVSFlagsAsync)

	tts2.SetPriority(sapi.SVPAlert)
	tts2.Speak("This will interrupt the previous speech", sapi.SVSFlagsAsync)

	time.Sleep(10 * time.Second)
}
