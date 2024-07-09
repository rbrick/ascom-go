package main

import (
	"reflect"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	CoverNotPresent CoverState = iota
	CoverClosed
	CoverMoving
	CoverOpen
	CoverUnknown
	CoverError
)

type CoverState int32

func (c CoverState) String() string {

	return reflect.ValueOf(c).Type().Name()
}

type CoverCalibrator interface {
	Device() AscomDevice

	// properties
	Brightness() int32
	CalibratorState() int32
	CoverState() CoverState
	MaxBrightness() int32

	// methods

	// will block until cover is open/closed
	WaitForOpenCover()
	WaitForCloseCover()

	CloseCover()
	OpenCover()
	HaltCover()

	CalibratorOn(brightness int)
	CalibratorOff()
}

type ascomCoverCalibratorImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomCoverCalibratorImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomCoverCalibratorImpl) Brightness() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "Brightness")
	return result.Value().(int32)
}

func (i *ascomCoverCalibratorImpl) MaxBrightness() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxBrightness")
	return result.Value().(int32)
}

func (i *ascomCoverCalibratorImpl) CalibratorState() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "CalibratorState")
	return result.Value().(int32)
}

func (i *ascomCoverCalibratorImpl) CoverState() CoverState {
	result := oleutil.MustGetProperty(i.dispatch, "CoverState")
	return CoverState(result.Value().(int32))
}

func (i *ascomCoverCalibratorImpl) WaitForOpenCover() {
	state := i.CoverState()

	if state != CoverOpen {
		i.OpenCover()

		tick := time.NewTicker(500 * time.Millisecond)

		for range tick.C {
			if i.CoverState() == CoverOpen {
				return
			}
		}
	}
}

func (i *ascomCoverCalibratorImpl) WaitForCloseCover() {
	state := i.CoverState()

	if state != CoverClosed {
		i.CloseCover()

		tick := time.NewTicker(500 * time.Millisecond)

		for range tick.C {
			if i.CoverState() == CoverClosed {
				return
			}
		}
	}
}

func (i *ascomCoverCalibratorImpl) OpenCover() {
	oleutil.MustCallMethod(i.dispatch, "OpenCover")

}
func (i *ascomCoverCalibratorImpl) CloseCover() {
	oleutil.MustCallMethod(i.dispatch, "CloseCover")
}
func (i *ascomCoverCalibratorImpl) HaltCover() {
	oleutil.MustCallMethod(i.dispatch, "HaltCover")
}
func (i *ascomCoverCalibratorImpl) CalibratorOn(brightness int) {
	oleutil.MustCallMethod(i.dispatch, "CalibratorOn", brightness)
}
func (i *ascomCoverCalibratorImpl) CalibratorOff() {
	oleutil.MustCallMethod(i.dispatch, "CalibratorOff")
}

func NewCoverCalibrator(progId string) (CoverCalibrator, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomCoverCalibratorImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil

}
