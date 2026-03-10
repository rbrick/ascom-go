package main

import (
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Focuser interface {
	Device() AscomDevice

	Absolute() bool
	Connecting() bool
	InterfaceVersion() int32
	IsMoving() bool
	Link() bool
	MaxIncrement() int32
	MaxStep() int32
	Position() int32
	StepSize() float64
	TempComp() bool
	SetTempComp(bool)
	TempCompAvailable() bool
	Temperature() float64

	Halt()
	Move(position int)
	WaitForMoveComplete()
}

type ascomFocuserImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomFocuserImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomFocuserImpl) Absolute() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Absolute")
	return result.Value().(bool)
}

func (i *ascomFocuserImpl) Connecting() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connecting")
	return result.Value().(bool)
}

func (i *ascomFocuserImpl) InterfaceVersion() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(int32)
}

func (i *ascomFocuserImpl) IsMoving() bool {
	result := oleutil.MustGetProperty(i.dispatch, "IsMoving")
	return result.Value().(bool)
}

func (i *ascomFocuserImpl) Link() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Link")
	return result.Value().(bool)
}

func (i *ascomFocuserImpl) MaxIncrement() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxIncrement")
	return result.Value().(int32)
}

func (i *ascomFocuserImpl) MaxStep() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxStep")
	return result.Value().(int32)
}

func (i *ascomFocuserImpl) Position() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "Position")
	return result.Value().(int32)
}

func (i *ascomFocuserImpl) StepSize() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "StepSize")
	return result.Value().(float64)
}

func (i *ascomFocuserImpl) TempComp() bool {
	result := oleutil.MustGetProperty(i.dispatch, "TempComp")
	return result.Value().(bool)
}

func (i *ascomFocuserImpl) SetTempComp(enabled bool) {
	oleutil.MustPutProperty(i.dispatch, "TempComp", enabled)
}

func (i *ascomFocuserImpl) TempCompAvailable() bool {
	result := oleutil.MustGetProperty(i.dispatch, "TempCompAvailable")
	return result.Value().(bool)
}

func (i *ascomFocuserImpl) Temperature() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "Temperature")
	return result.Value().(float64)
}

func (i *ascomFocuserImpl) Halt() {
	oleutil.MustCallMethod(i.dispatch, "Halt")
}

func (i *ascomFocuserImpl) Move(position int) {
	oleutil.MustCallMethod(i.dispatch, "Move", position)
}

func (i *ascomFocuserImpl) WaitForMoveComplete() {
	ticker := time.NewTicker(250 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if !i.IsMoving() {
			return
		}
	}
}

func NewFocuser(progId string) (Focuser, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomFocuserImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil
}
