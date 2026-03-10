package main

import (
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Switch interface {
	Device() AscomDevice

	CanAsync(id int) bool
	CanWrite(id int) bool
	Connecting() bool
	GetSwitch(id int) bool
	GetSwitchDescription(id int) string
	GetSwitchName(id int) string
	GetSwitchValue(id int) float64
	InterfaceVersion() int32
	MaxSwitch() int32
	MaxSwitchValue(id int) float64
	MinSwitchValue(id int) float64
	StateChangeComplete(id int) bool
	SwitchStep(id int) float64

	CancelAsync(id int)
	SetAsync(id int, state bool)
	SetAsyncValue(id int, value float64)
	SetSwitch(id int, state bool)
	SetSwitchName(id int, name string)
	SetSwitchValue(id int, value float64)

	WaitForStateChange(id int)
}

type ascomSwitchImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomSwitchImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomSwitchImpl) CanAsync(id int) bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanAsync", id)
	return result.Value().(bool)
}

func (i *ascomSwitchImpl) CanWrite(id int) bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanWrite", id)
	return result.Value().(bool)
}

func (i *ascomSwitchImpl) Connecting() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connecting")
	return result.Value().(bool)
}

func (i *ascomSwitchImpl) GetSwitch(id int) bool {
	result := oleutil.MustGetProperty(i.dispatch, "GetSwitch", id)
	return result.Value().(bool)
}

func (i *ascomSwitchImpl) GetSwitchDescription(id int) string {
	result := oleutil.MustGetProperty(i.dispatch, "GetSwitchDescription", id)
	return result.Value().(string)
}

func (i *ascomSwitchImpl) GetSwitchName(id int) string {
	result := oleutil.MustGetProperty(i.dispatch, "GetSwitchName", id)
	return result.Value().(string)
}

func (i *ascomSwitchImpl) GetSwitchValue(id int) float64 {
	result := oleutil.MustGetProperty(i.dispatch, "GetSwitchValue", id)
	return result.Value().(float64)
}

func (i *ascomSwitchImpl) InterfaceVersion() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(int32)
}

func (i *ascomSwitchImpl) MaxSwitch() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxSwitch")
	return result.Value().(int32)
}

func (i *ascomSwitchImpl) MaxSwitchValue(id int) float64 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxSwitchValue", id)
	return result.Value().(float64)
}

func (i *ascomSwitchImpl) MinSwitchValue(id int) float64 {
	result := oleutil.MustGetProperty(i.dispatch, "MinSwitchValue", id)
	return result.Value().(float64)
}

func (i *ascomSwitchImpl) StateChangeComplete(id int) bool {
	result := oleutil.MustGetProperty(i.dispatch, "StateChangeComplete", id)
	return result.Value().(bool)
}

func (i *ascomSwitchImpl) SwitchStep(id int) float64 {
	result := oleutil.MustGetProperty(i.dispatch, "SwitchStep", id)
	return result.Value().(float64)
}

func (i *ascomSwitchImpl) CancelAsync(id int) {
	oleutil.MustCallMethod(i.dispatch, "CancelAsync", id)
}

func (i *ascomSwitchImpl) SetAsync(id int, state bool) {
	oleutil.MustCallMethod(i.dispatch, "SetAsync", id, state)
}

func (i *ascomSwitchImpl) SetAsyncValue(id int, value float64) {
	oleutil.MustCallMethod(i.dispatch, "SetAsyncValue", id, value)
}

func (i *ascomSwitchImpl) SetSwitch(id int, state bool) {
	oleutil.MustCallMethod(i.dispatch, "SetSwitch", id, state)
}

func (i *ascomSwitchImpl) SetSwitchName(id int, name string) {
	oleutil.MustCallMethod(i.dispatch, "SetSwitchName", id, name)
}

func (i *ascomSwitchImpl) SetSwitchValue(id int, value float64) {
	oleutil.MustCallMethod(i.dispatch, "SetSwitchValue", id, value)
}

func (i *ascomSwitchImpl) WaitForStateChange(id int) {
	ticker := time.NewTicker(250 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if i.StateChangeComplete(id) {
			return
		}
	}
}

func NewSwitch(progId string) (Switch, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomSwitchImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil
}
