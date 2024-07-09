package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type AscomDriver interface {
	Connect()
	Disconnect()

	Connected() bool
	Description() string
	DriverInfo() string
	DriverVersion() string
	InterfaceVersion() string
	Name() string
	SupportedActions() string
}

type ascomDriverImpl struct {
	dispatch *ole.IDispatch
}

func (i *ascomDriverImpl) Connect() {
	oleutil.MustPutProperty(i.dispatch, "Connected", true)
}

func (i *ascomDriverImpl) Disconnect() {
	oleutil.MustPutProperty(i.dispatch, "Connected", false)
}

func (i *ascomDriverImpl) Connected() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connected")
	return result.Value().(bool)
}

func (i *ascomDriverImpl) Description() string {
	result := oleutil.MustGetProperty(i.dispatch, "Description")
	return result.Value().(string)
}

func (i *ascomDriverImpl) DriverInfo() string {
	result := oleutil.MustGetProperty(i.dispatch, "DriverInfo")
	return result.Value().(string)
}

func (i *ascomDriverImpl) DriverVersion() string {
	result := oleutil.MustGetProperty(i.dispatch, "DriverVersion")
	return result.Value().(string)
}

func (i *ascomDriverImpl) InterfaceVersion() string {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(string)
}

func (i *ascomDriverImpl) Name() string {
	result := oleutil.MustGetProperty(i.dispatch, "Name")
	return result.Value().(string)
}

func (i *ascomDriverImpl) SupportedActions() string {
	result := oleutil.MustGetProperty(i.dispatch, "SupportedActions")
	return result.Value().(string)
}

func NewAscomDriver(dispatch *ole.IDispatch) AscomDriver {
	return &ascomDriverImpl{
		dispatch,
	}
}
