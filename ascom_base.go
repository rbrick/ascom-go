package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type COMObject interface {
	Destroy()
}

type AscomDevice interface {
	Driver() AscomDriver
	Choose(device string) string
}

type ascomDeviceImpl struct {
	driver   AscomDriver
	dispatch *ole.IDispatch
}

func (i *ascomDeviceImpl) Choose(device string) string {
	result := oleutil.MustCallMethod(i.dispatch, "Choose", device)
	return result.Value().(string)
}

func (i *ascomDeviceImpl) Driver() AscomDriver {
	return i.driver
}

func NewAscomDevice(dispatch *ole.IDispatch) AscomDevice {
	return &ascomDeviceImpl{
		driver:   NewAscomDriver(dispatch),
		dispatch: dispatch,
	}
}
