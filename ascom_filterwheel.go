package main

import (
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type FilterWheel interface {
	Device() AscomDevice

	Connected() bool
	SetConnected(bool)
	Connecting() bool
	FocusOffsets() []int32
	InterfaceVersion() int32
	Names() []string
	Position() int32
	SetPosition(int)
	SupportedActions() []string

	Action(actionName string, actionParameters string) string
	CommandBlind(command string, raw bool)
	CommandBool(command string, raw bool) bool
	CommandString(command string, raw bool) string
	Connect()
	Disconnect()
	SetupDialog()

	WaitForConnectingComplete()
	WaitForPositionComplete()
}

type ascomFilterWheelImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomFilterWheelImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomFilterWheelImpl) Connected() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connected")
	return result.Value().(bool)
}

func (i *ascomFilterWheelImpl) SetConnected(value bool) {
	oleutil.MustPutProperty(i.dispatch, "Connected", value)
}

func (i *ascomFilterWheelImpl) Connecting() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connecting")
	return result.Value().(bool)
}

func (i *ascomFilterWheelImpl) FocusOffsets() []int32 {
	result := oleutil.MustGetProperty(i.dispatch, "FocusOffsets")
	array := result.ToArray()
	defer array.Release()

	values := array.ToValueArray()
	offsets := make([]int32, len(values))

	for index, value := range values {
		offsets[index] = value.(int32)
	}

	return offsets
}

func (i *ascomFilterWheelImpl) InterfaceVersion() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(int32)
}

func (i *ascomFilterWheelImpl) Names() []string {
	result := oleutil.MustGetProperty(i.dispatch, "Names")
	array := result.ToArray()
	defer array.Release()
	return array.ToStringArray()
}

func (i *ascomFilterWheelImpl) Position() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "Position")
	return result.Value().(int32)
}

func (i *ascomFilterWheelImpl) SetPosition(value int) {
	oleutil.MustPutProperty(i.dispatch, "Position", value)
}

func (i *ascomFilterWheelImpl) SupportedActions() []string {
	result := oleutil.MustGetProperty(i.dispatch, "SupportedActions")
	array := result.ToArray()
	defer array.Release()
	return array.ToStringArray()
}

func (i *ascomFilterWheelImpl) Action(actionName string, actionParameters string) string {
	result := oleutil.MustCallMethod(i.dispatch, "Action", actionName, actionParameters)
	return result.Value().(string)
}

func (i *ascomFilterWheelImpl) CommandBlind(command string, raw bool) {
	oleutil.MustCallMethod(i.dispatch, "CommandBlind", command, raw)
}

func (i *ascomFilterWheelImpl) CommandBool(command string, raw bool) bool {
	result := oleutil.MustCallMethod(i.dispatch, "CommandBool", command, raw)
	return result.Value().(bool)
}

func (i *ascomFilterWheelImpl) CommandString(command string, raw bool) string {
	result := oleutil.MustCallMethod(i.dispatch, "CommandString", command, raw)
	return result.Value().(string)
}

func (i *ascomFilterWheelImpl) Connect() {
	oleutil.MustCallMethod(i.dispatch, "Connect")
}

func (i *ascomFilterWheelImpl) Disconnect() {
	oleutil.MustCallMethod(i.dispatch, "Disconnect")
}

func (i *ascomFilterWheelImpl) SetupDialog() {
	oleutil.MustCallMethod(i.dispatch, "SetupDialog")
}

func (i *ascomFilterWheelImpl) WaitForConnectingComplete() {
	ticker := time.NewTicker(500 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if !i.Connecting() {
			return
		}
	}
}

func (i *ascomFilterWheelImpl) WaitForPositionComplete() {
	ticker := time.NewTicker(250 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if i.Position() != -1 {
			return
		}
	}
}

func NewFilterWheel(progId string) (FilterWheel, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomFilterWheelImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil
}
