package sapi

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
func (s *Sapi) Say(message string) error {
	_, err := oleutil.CallMethod(s.voice, "Speak", message, 0)

	return err
}

// SetRate sets the rate of speech
func (s *Sapi) SetRate(rate int) error {
	_, err := oleutil.PutProperty(s.voice, "Rate", rate)

	return err
}

// SetVolume sets the volume of the speech
func (s *Sapi) SetVolume(volume int) error {
	_, err := oleutil.PutProperty(s.voice, "Volume", volume)

	return err
}
