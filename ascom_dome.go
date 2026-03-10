package main

import (
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type ShutterState int32

const (
	ShutterOpen ShutterState = iota
	ShutterClosed
	ShutterOpening
	ShutterClosing
	ShutterError
)

type Dome interface {
	Device() AscomDevice

	Altitude() float64
	AtHome() bool
	AtPark() bool
	Azimuth() float64
	CanFindHome() bool
	CanPark() bool
	CanSetAltitude() bool
	CanSetAzimuth() bool
	CanSetPark() bool
	CanSetShutter() bool
	CanSlave() bool
	CanSyncAzimuth() bool
	Connecting() bool
	InterfaceVersion() int32
	ShutterStatus() ShutterState
	Slaved() bool
	SetSlaved(bool)
	Slewing() bool

	AbortSlew()
	CloseShutter()
	FindHome()
	OpenShutter()
	Park()
	SetPark()
	SlewToAltitude(altitude float64)
	SlewToAzimuth(azimuth float64)
	SyncToAzimuth(azimuth float64)

	WaitForSlewComplete()
	WaitForShutterComplete()
}

type ascomDomeImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomDomeImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomDomeImpl) Altitude() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "Altitude")
	return result.Value().(float64)
}

func (i *ascomDomeImpl) AtHome() bool {
	result := oleutil.MustGetProperty(i.dispatch, "AtHome")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) AtPark() bool {
	result := oleutil.MustGetProperty(i.dispatch, "AtPark")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) Azimuth() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "Azimuth")
	return result.Value().(float64)
}

func (i *ascomDomeImpl) CanFindHome() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanFindHome")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) CanPark() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanPark")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) CanSetAltitude() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetAltitude")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) CanSetAzimuth() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetAzimuth")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) CanSetPark() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetPark")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) CanSetShutter() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetShutter")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) CanSlave() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSlave")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) CanSyncAzimuth() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSyncAzimuth")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) Connecting() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connecting")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) InterfaceVersion() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(int32)
}

func (i *ascomDomeImpl) ShutterStatus() ShutterState {
	result := oleutil.MustGetProperty(i.dispatch, "ShutterStatus")
	return ShutterState(result.Value().(int32))
}

func (i *ascomDomeImpl) Slaved() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Slaved")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) SetSlaved(slaved bool) {
	oleutil.MustPutProperty(i.dispatch, "Slaved", slaved)
}

func (i *ascomDomeImpl) Slewing() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Slewing")
	return result.Value().(bool)
}

func (i *ascomDomeImpl) AbortSlew() {
	oleutil.MustCallMethod(i.dispatch, "AbortSlew")
}

func (i *ascomDomeImpl) CloseShutter() {
	oleutil.MustCallMethod(i.dispatch, "CloseShutter")
}

func (i *ascomDomeImpl) FindHome() {
	oleutil.MustCallMethod(i.dispatch, "FindHome")
}

func (i *ascomDomeImpl) OpenShutter() {
	oleutil.MustCallMethod(i.dispatch, "OpenShutter")
}

func (i *ascomDomeImpl) Park() {
	oleutil.MustCallMethod(i.dispatch, "Park")
}

func (i *ascomDomeImpl) SetPark() {
	oleutil.MustCallMethod(i.dispatch, "SetPark")
}

func (i *ascomDomeImpl) SlewToAltitude(altitude float64) {
	oleutil.MustCallMethod(i.dispatch, "SlewToAltitude", altitude)
}

func (i *ascomDomeImpl) SlewToAzimuth(azimuth float64) {
	oleutil.MustCallMethod(i.dispatch, "SlewToAzimuth", azimuth)
}

func (i *ascomDomeImpl) SyncToAzimuth(azimuth float64) {
	oleutil.MustCallMethod(i.dispatch, "SyncToAzimuth", azimuth)
}

func (i *ascomDomeImpl) WaitForSlewComplete() {
	ticker := time.NewTicker(500 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if !i.Slewing() {
			return
		}
	}
}

func (i *ascomDomeImpl) WaitForShutterComplete() {
	ticker := time.NewTicker(500 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		switch i.ShutterStatus() {
		case ShutterOpen, ShutterClosed, ShutterError:
			return
		}
	}
}

func NewDome(progId string) (Dome, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomDomeImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil
}
