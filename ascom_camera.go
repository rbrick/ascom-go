package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type CameraState int32

const (
	CameraIdle CameraState = iota
	CameraWaiting
	CameraExposing
	CameraReading
	CameraDownloading
	CameraError
)

type GuideDirection int32

const (
	GuideNorth GuideDirection = iota
	GuideSouth
	GuideEast
	GuideWest
)

type SensorType int32

const (
	MonochromeSensor SensorType = iota
	ColorSensor
	RGGBSensor
	CMYGSensor
	CMYG2Sensor
	LRGBSensor
)

type Camera interface {
	Device() AscomDevice

	BayerOffsetX() int32
	BayerOffsetY() int32
	BinX() int32
	SetBinX(int)
	BinY() int32
	SetBinY(int)
	CameraState() CameraState
	CameraXSize() int32
	CameraYSize() int32
	CanAbortExposure() bool
	CanAsymmetricBin() bool
	CanFastReadout() bool
	CanGetCoolerPower() bool
	CanPulseGuide() bool
	CanSetCCDTemperature() bool
	CanStopExposure() bool
	CCDTemperature() float64
	Connecting() bool
	CoolerOn() bool
	SetCoolerOn(bool)
	CoolerPower() float64
	ElectronsPerADU() float64
	ExposureMax() float64
	ExposureMin() float64
	ExposureResolution() float64
	FastReadout() bool
	SetFastReadout(bool)
	FullWellCapacity() float64
	Gain() int32
	SetGain(int)
	GainMax() int32
	GainMin() int32
	Gains() []string
	HasShutter() bool
	HeatSinkTemperature() float64
	ImageArray() []interface{}
	ImageArrayVariant() []interface{}
	ImageReady() bool
	InterfaceVersion() int32
	IsPulseGuiding() bool
	LastExposureDuration() float64
	LastExposureStartTime() string
	MaxADU() int32
	MaxBinX() int32
	MaxBinY() int32
	NumX() int32
	SetNumX(int)
	NumY() int32
	SetNumY(int)
	Offset() int32
	SetOffset(int)
	OffsetMax() int32
	OffsetMin() int32
	Offsets() []string
	PercentCompleted() int32
	PixelSizeX() float64
	PixelSizeY() float64
	ReadoutMode() int32
	SetReadoutMode(int)
	ReadoutModes() []string
	SensorName() string
	SensorType() SensorType
	SetCCDTemperature() float64
	SetSetCCDTemperature(float64)
	StartX() int32
	SetStartX(int)
	StartY() int32
	SetStartY(int)
	SubExposureDuration() float64
	SetSubExposureDuration(float64)

	AbortExposure()
	PulseGuide(direction GuideDirection, duration int)
	StartExposure(duration float64, light bool)
	StopExposure()
}

type ascomCameraImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomCameraImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomCameraImpl) BayerOffsetX() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "BayerOffsetX")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) BayerOffsetY() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "BayerOffsetY")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) BinX() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "BinX")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetBinX(value int) {
	oleutil.MustPutProperty(i.dispatch, "BinX", value)
}

func (i *ascomCameraImpl) BinY() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "BinY")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetBinY(value int) {
	oleutil.MustPutProperty(i.dispatch, "BinY", value)
}

func (i *ascomCameraImpl) CameraState() CameraState {
	result := oleutil.MustGetProperty(i.dispatch, "CameraState")
	return CameraState(result.Value().(int32))
}

func (i *ascomCameraImpl) CameraXSize() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "CameraXSize")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) CameraYSize() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "CameraYSize")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) CanAbortExposure() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanAbortExposure")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CanAsymmetricBin() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanAsymmetricBin")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CanFastReadout() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanFastReadout")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CanGetCoolerPower() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanGetCoolerPower")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CanPulseGuide() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanPulseGuide")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CanSetCCDTemperature() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetCCDTemperature")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CanStopExposure() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanStopExposure")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CCDTemperature() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "CCDTemperature")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) Connecting() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connecting")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) CoolerOn() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CoolerOn")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) SetCoolerOn(value bool) {
	oleutil.MustPutProperty(i.dispatch, "CoolerOn", value)
}

func (i *ascomCameraImpl) CoolerPower() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "CoolerPower")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) ElectronsPerADU() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "ElectronsPerADU")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) ExposureMax() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "ExposureMax")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) ExposureMin() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "ExposureMin")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) ExposureResolution() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "ExposureResolution")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) FastReadout() bool {
	result := oleutil.MustGetProperty(i.dispatch, "FastReadout")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) SetFastReadout(value bool) {
	oleutil.MustPutProperty(i.dispatch, "FastReadout", value)
}

func (i *ascomCameraImpl) FullWellCapacity() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "FullWellCapacity")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) Gain() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "Gain")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetGain(value int) {
	oleutil.MustPutProperty(i.dispatch, "Gain", value)
}

func (i *ascomCameraImpl) GainMax() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "GainMax")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) GainMin() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "GainMin")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) Gains() []string {
	result := oleutil.MustGetProperty(i.dispatch, "Gains")
	array := result.ToArray()
	defer array.Release()
	return array.ToStringArray()
}

func (i *ascomCameraImpl) HasShutter() bool {
	result := oleutil.MustGetProperty(i.dispatch, "HasShutter")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) HeatSinkTemperature() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "HeatSinkTemperature")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) ImageArray() []interface{} {
	result := oleutil.MustGetProperty(i.dispatch, "ImageArray")
	array := result.ToArray()
	defer array.Release()
	return array.ToValueArray()
}

func (i *ascomCameraImpl) ImageArrayVariant() []interface{} {
	result := oleutil.MustGetProperty(i.dispatch, "ImageArrayVariant")
	array := result.ToArray()
	defer array.Release()
	return array.ToValueArray()
}

func (i *ascomCameraImpl) ImageReady() bool {
	result := oleutil.MustGetProperty(i.dispatch, "ImageReady")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) InterfaceVersion() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) IsPulseGuiding() bool {
	result := oleutil.MustGetProperty(i.dispatch, "IsPulseGuiding")
	return result.Value().(bool)
}

func (i *ascomCameraImpl) LastExposureDuration() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "LastExposureDuration")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) LastExposureStartTime() string {
	result := oleutil.MustGetProperty(i.dispatch, "LastExposureStartTime")
	return result.Value().(string)
}

func (i *ascomCameraImpl) MaxADU() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxADU")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) MaxBinX() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxBinX")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) MaxBinY() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "MaxBinY")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) NumX() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "NumX")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetNumX(value int) {
	oleutil.MustPutProperty(i.dispatch, "NumX", value)
}

func (i *ascomCameraImpl) NumY() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "NumY")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetNumY(value int) {
	oleutil.MustPutProperty(i.dispatch, "NumY", value)
}

func (i *ascomCameraImpl) Offset() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "Offset")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetOffset(value int) {
	oleutil.MustPutProperty(i.dispatch, "Offset", value)
}

func (i *ascomCameraImpl) OffsetMax() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "OffsetMax")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) OffsetMin() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "OffsetMin")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) Offsets() []string {
	result := oleutil.MustGetProperty(i.dispatch, "Offsets")
	array := result.ToArray()
	defer array.Release()
	return array.ToStringArray()
}

func (i *ascomCameraImpl) PercentCompleted() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "PercentCompleted")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) PixelSizeX() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "PixelSizeX")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) PixelSizeY() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "PixelSizeY")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) ReadoutMode() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "ReadoutMode")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetReadoutMode(value int) {
	oleutil.MustPutProperty(i.dispatch, "ReadoutMode", value)
}

func (i *ascomCameraImpl) ReadoutModes() []string {
	result := oleutil.MustGetProperty(i.dispatch, "ReadoutModes")
	array := result.ToArray()
	defer array.Release()
	return array.ToStringArray()
}

func (i *ascomCameraImpl) SensorName() string {
	result := oleutil.MustGetProperty(i.dispatch, "SensorName")
	return result.Value().(string)
}

func (i *ascomCameraImpl) SensorType() SensorType {
	result := oleutil.MustGetProperty(i.dispatch, "SensorType")
	return SensorType(result.Value().(int32))
}

func (i *ascomCameraImpl) SetCCDTemperature() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "SetCCDTemperature")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) SetSetCCDTemperature(value float64) {
	oleutil.MustPutProperty(i.dispatch, "SetCCDTemperature", value)
}

func (i *ascomCameraImpl) StartX() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "StartX")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetStartX(value int) {
	oleutil.MustPutProperty(i.dispatch, "StartX", value)
}

func (i *ascomCameraImpl) StartY() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "StartY")
	return result.Value().(int32)
}

func (i *ascomCameraImpl) SetStartY(value int) {
	oleutil.MustPutProperty(i.dispatch, "StartY", value)
}

func (i *ascomCameraImpl) SubExposureDuration() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "SubExposureDuration")
	return result.Value().(float64)
}

func (i *ascomCameraImpl) SetSubExposureDuration(value float64) {
	oleutil.MustPutProperty(i.dispatch, "SubExposureDuration", value)
}

func (i *ascomCameraImpl) AbortExposure() {
	oleutil.MustCallMethod(i.dispatch, "AbortExposure")
}

func (i *ascomCameraImpl) PulseGuide(direction GuideDirection, duration int) {
	oleutil.MustCallMethod(i.dispatch, "PulseGuide", int(direction), duration)
}

func (i *ascomCameraImpl) StartExposure(duration float64, light bool) {
	oleutil.MustCallMethod(i.dispatch, "StartExposure", duration, light)
}

func (i *ascomCameraImpl) StopExposure() {
	oleutil.MustCallMethod(i.dispatch, "StopExposure")
}

func NewCamera(progId string) (Camera, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomCameraImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil
}
