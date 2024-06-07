// Package sapi has its documentation on: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ee125647(v=vs.85)
package sapi

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Sapi is a structure that wraps the SAPI COM object
type Sapi struct {
	voice *ole.IDispatch
}

const (
	SVSFDefault          = 0
	SVSFlagsAsync        = 1
	SVSFPurgeBeforeSpeak = 2
	SVSFIsFilename       = 4
	SVSFIsXML            = 8
	SVSFIsNotXML         = 16
	SVSFPersistXML       = 32

	// Normalizer Flags
	SVSFNLPSpeakPunc = 64

	// Masks
	SVSFNLPMask     = 64
	SVSFVoiceMask   = 127
	SVSFUnusedFlags = -128
)

const (
	SVPNormal = 0
	SVPAlert  = 1
	SVPOver   = 2
)

// NewSapi creates the SAPI TTS object
func NewSapi() (*Sapi, error) {
	ole.CoInitialize(0)
	voice_obj, err := oleutil.CreateObject("SAPI.SpVoice")
	if err != nil {
		return nil, err
	}
	voice, err := voice_obj.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}

	return &Sapi{voice: voice}, nil
}

/////////////////////////////////////////////////////////////////
// Main methods

// Speak initiates the speaking of a text string, a text file, an XML file, or a wave file by the voice.
func (s *Sapi) Speak(message string, flags int) {
	_, _ = oleutil.CallMethod(s.voice, "Speak", message, flags)
}

// Pause pauses the voice at the nearest alert boundary and closes the output device, allowing it to be used by other voices.
func (s *Sapi) Pause() {
	_, _ = oleutil.CallMethod(s.voice, "Pause")
}

// Resume causes the voice to resume speaking when paused.
func (s *Sapi) Resume() {
	_, _ = oleutil.CallMethod(s.voice, "Resume")
}

// WaitUntilDone blocks the caller until either the voice has finished speaking or the specified time interval has elapsed.
func (s *Sapi) WaitUntilDone(ms_timeout int) {
	_, _ = oleutil.CallMethod(s.voice, "WaitUntilDone", ms_timeout)
}

// Skip skips the voice forward or backward by the specified number of "Sentence" items within the current input text stream.
func (s *Sapi) Skip(num_items int) {
	_, _ = oleutil.CallMethod(s.voice, "Skip", "Sentence", num_items)
}

/////////////////////////////////////////////////////////////////
// Setters

// SetRate sets the speaking rate of the voice.
func (s *Sapi) SetRate(rate int) {
	_, _ = oleutil.PutProperty(s.voice, "Rate", rate)
}

// SetVolume sets the base volume (loudness) level of the voice.
func (s *Sapi) SetVolume(volume int) {
	_, _ = oleutil.PutProperty(s.voice, "Volume", volume)
}

// SetPriority sets the priority level of the voice.
func (s *Sapi) SetPriority(priority int) {
	_, _ = oleutil.PutProperty(s.voice, "Priority", priority)
}

/////////////////////////////////////////////////////////////////
// Getters

// GetRate gets the speaking rate of the voice.
func (s *Sapi) GetRate() int {
	rate, _ := oleutil.GetProperty(s.voice, "Rate")

	return int(rate.Val)
}

// GetVolume gets the base volume (loudness) level of the voice.
func (s *Sapi) GetVolume() int {
	volume, _ := oleutil.GetProperty(s.voice, "Volume")

	return int(volume.Val)
}

// GetPriority gets the priority level of the voice.
func (s *Sapi) GetPriority() int {
	priority, _ := oleutil.GetProperty(s.voice, "Priority")

	return int(priority.Val)
}
