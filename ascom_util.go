package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type DeviceType string

const (
	TelescopeDevice           DeviceType = "Telescope"
	FocuserDevice             DeviceType = "Focuser"
	CoverCalibratorDevice     DeviceType = "CoverCalibrator"
	DomeDevice                DeviceType = "Dome"
	CameraDevice              DeviceType = "Camera"
	RotatorDevice             DeviceType = "Rotator"
	FilterWheelDevice         DeviceType = "FilterWheel"
	SafetyMonitorDevice       DeviceType = "SafetyMonitor"
	VideoDevice               DeviceType = "Video"
	SwitchDevice              DeviceType = "Switch"
	ObservingConditionsDevice DeviceType = "ObservingConditions"
)

type Chooser interface {
	COMObject
	Choose(devId string) (string, error)
	SetDeviceType(DeviceType)
}

type ascomChooser struct {
	comObject *ole.IUnknown
	dispatch  *ole.IDispatch
}

func (i *ascomChooser) SetDeviceType(deviceType DeviceType) {
	oleutil.MustPutProperty(i.dispatch, "DeviceType", string(deviceType))
}

func (i *ascomChooser) Choose(devId string) (string, error) {
	result, err := oleutil.CallMethod(i.dispatch, "Choose", devId)

	if err != nil {
		return "", err
	}

	return result.Value().(string), nil
}

func (i *ascomChooser) Destroy() {
	i.dispatch.Release()
	i.comObject.Release()
}

func NewChooser() (Chooser, error) {
	// new Chooser();
	chooserObject, err := oleutil.CreateObject("ASCOM.Utilities.Chooser")

	if err != nil {
		return nil, err
	}

	dispatch, err := chooserObject.QueryInterface(ole.IID_IDispatch)

	if err != nil {
		return nil, err
	}

	return &ascomChooser{
		comObject: chooserObject,
		dispatch:  dispatch,
	}, nil
}

type Util interface {
}

// implementation
type ascomUtil struct{}
