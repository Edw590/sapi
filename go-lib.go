package sapi_go

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Sapi is a structure that wraps the SAPI COM object
type Sapi struct {
	voice *ole.IDispatch
}

// NewSapi initializes the SAPI object
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

// Say speaks the given text
func (s *Sapi) Say(message string) {
	oleutil.MustCallMethod(s.voice, "Speak", message, 0)
}

// SetRate sets the rate of speech
func (s *Sapi) SetRate(rate int) {
	oleutil.PutProperty(s.voice, "Rate", rate)
}

// SetVolume sets the volume of the speech
func (s *Sapi) SetVolume(volume int) {
	oleutil.PutProperty(s.voice, "Volume", volume)
}
