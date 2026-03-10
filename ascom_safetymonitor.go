package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type SafetyMonitor interface {
	Device() AscomDevice

	Connecting() bool
	InterfaceVersion() int32
	IsSafe() bool
}

type ascomSafetyMonitorImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomSafetyMonitorImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomSafetyMonitorImpl) Connecting() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connecting")
	return result.Value().(bool)
}

func (i *ascomSafetyMonitorImpl) InterfaceVersion() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(int32)
}

func (i *ascomSafetyMonitorImpl) IsSafe() bool {
	result := oleutil.MustGetProperty(i.dispatch, "IsSafe")
	return result.Value().(bool)
}

func NewSafetyMonitor(progId string) (SafetyMonitor, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomSafetyMonitorImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil
}
