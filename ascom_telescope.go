package main

import (
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type AlignmentMode int32

const (
	AltAzAlignment AlignmentMode = iota
	PolarAlignment
	GermanPolarAlignment
)

type DriveRate int32

const (
	DriveSidereal DriveRate = iota
	DriveLunar
	DriveSolar
	DriveKing
)

type EquatorialCoordinateType int32

const (
	EquatorialOther EquatorialCoordinateType = iota
	EquatorialTopocentric
	EquatorialJ2000
	EquatorialJ2050
	EquatorialB1950
)

type PierSide int32

const (
	PierUnknown PierSide = -1
	PierEast    PierSide = 0
	PierWest    PierSide = 1
)

type TelescopeAxis int32

const (
	PrimaryAxis TelescopeAxis = iota
	SecondaryAxis
	TertiaryAxis
)

type Telescope interface {
	Device() AscomDevice

	AlignmentMode() AlignmentMode
	Altitude() float64
	ApertureArea() float64
	ApertureDiameter() float64
	AtHome() bool
	AtPark() bool
	Azimuth() float64
	CanFindHome() bool
	CanMoveAxis(TelescopeAxis) bool
	CanPark() bool
	CanPulseGuide() bool
	CanSetDeclinationRate() bool
	CanSetGuideRates() bool
	CanSetPark() bool
	CanSetPierSide() bool
	CanSetRightAscensionRate() bool
	CanSetTracking() bool
	CanSlew() bool
	CanSlewAltAz() bool
	CanSlewAltAzAsync() bool
	CanSlewAsync() bool
	CanSync() bool
	CanSyncAltAz() bool
	CanUnpark() bool
	Connecting() bool
	Declination() float64
	DeclinationRate() float64
	SetDeclinationRate(float64)
	DoesRefraction() bool
	SetDoesRefraction(bool)
	EquatorialSystem() EquatorialCoordinateType
	FocalLength() float64
	GuideRateDeclination() float64
	SetGuideRateDeclination(float64)
	GuideRateRightAscension() float64
	SetGuideRateRightAscension(float64)
	InterfaceVersion() int32
	IsPulseGuiding() bool
	RightAscension() float64
	RightAscensionRate() float64
	SetRightAscensionRate(float64)
	SideOfPier() PierSide
	SetSideOfPier(PierSide)
	SiderealTime() float64
	SiteElevation() float64
	SetSiteElevation(float64)
	SiteLatitude() float64
	SetSiteLatitude(float64)
	SiteLongitude() float64
	SetSiteLongitude(float64)
	SlewSettleTime() int32
	SetSlewSettleTime(int)
	Slewing() bool
	TargetDeclination() float64
	SetTargetDeclination(float64)
	TargetRightAscension() float64
	SetTargetRightAscension(float64)
	Tracking() bool
	SetTracking(bool)
	TrackingRate() DriveRate
	SetTrackingRate(DriveRate)

	AbortSlew()
	Action(actionName string, actionParameters string) string
	CommandBlind(command string, raw bool)
	CommandBool(command string, raw bool) bool
	CommandString(command string, raw bool) string
	Connect()
	DestinationSideOfPier(rightAscension float64, declination float64) PierSide
	Disconnect()
	FindHome()
	MoveAxis(axis TelescopeAxis, rate float64)
	Park()
	PulseGuide(direction GuideDirection, duration int)
	SetPark()
	SetupDialog()
	SlewToAltAz(azimuth float64, altitude float64)
	SlewToAltAzAsync(azimuth float64, altitude float64)
	SlewToCoordinates(rightAscension float64, declination float64)
	SlewToCoordinatesAsync(rightAscension float64, declination float64)
	SlewToTarget()
	SlewToTargetAsync()
	SyncToAltAz(azimuth float64, altitude float64)
	SyncToCoordinates(rightAscension float64, declination float64)
	SyncToTarget()
	Unpark()

	WaitForConnectingComplete()
	WaitForPulseGuideComplete()
	WaitForSlewComplete()
}

type ascomTelescopeImpl struct {
	device   AscomDevice
	dispatch *ole.IDispatch
}

func (i *ascomTelescopeImpl) Device() AscomDevice {
	return i.device
}

func (i *ascomTelescopeImpl) AlignmentMode() AlignmentMode {
	result := oleutil.MustGetProperty(i.dispatch, "AlignmentMode")
	return AlignmentMode(result.Value().(int32))
}

func (i *ascomTelescopeImpl) Altitude() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "Altitude")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) ApertureArea() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "ApertureArea")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) ApertureDiameter() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "ApertureDiameter")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) AtHome() bool {
	result := oleutil.MustGetProperty(i.dispatch, "AtHome")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) AtPark() bool {
	result := oleutil.MustGetProperty(i.dispatch, "AtPark")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) Azimuth() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "Azimuth")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) CanFindHome() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanFindHome")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanMoveAxis(axis TelescopeAxis) bool {
	result := oleutil.MustCallMethod(i.dispatch, "CanMoveAxis", int(axis))
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanPark() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanPark")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanPulseGuide() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanPulseGuide")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSetDeclinationRate() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetDeclinationRate")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSetGuideRates() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetGuideRates")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSetPark() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetPark")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSetPierSide() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetPierSide")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSetRightAscensionRate() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetRightAscensionRate")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSetTracking() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSetTracking")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSlew() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSlew")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSlewAltAz() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSlewAltAz")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSlewAltAzAsync() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSlewAltAzAsync")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSlewAsync() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSlewAsync")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSync() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSync")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanSyncAltAz() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanSyncAltAz")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CanUnpark() bool {
	result := oleutil.MustGetProperty(i.dispatch, "CanUnpark")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) Connecting() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Connecting")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) Declination() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "Declination")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) DeclinationRate() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "DeclinationRate")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetDeclinationRate(value float64) {
	oleutil.MustPutProperty(i.dispatch, "DeclinationRate", value)
}

func (i *ascomTelescopeImpl) DoesRefraction() bool {
	result := oleutil.MustGetProperty(i.dispatch, "DoesRefraction")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) SetDoesRefraction(value bool) {
	oleutil.MustPutProperty(i.dispatch, "DoesRefraction", value)
}

func (i *ascomTelescopeImpl) EquatorialSystem() EquatorialCoordinateType {
	result := oleutil.MustGetProperty(i.dispatch, "EquatorialSystem")
	return EquatorialCoordinateType(result.Value().(int32))
}

func (i *ascomTelescopeImpl) FocalLength() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "FocalLength")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) GuideRateDeclination() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "GuideRateDeclination")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetGuideRateDeclination(value float64) {
	oleutil.MustPutProperty(i.dispatch, "GuideRateDeclination", value)
}

func (i *ascomTelescopeImpl) GuideRateRightAscension() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "GuideRateRightAscension")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetGuideRateRightAscension(value float64) {
	oleutil.MustPutProperty(i.dispatch, "GuideRateRightAscension", value)
}

func (i *ascomTelescopeImpl) InterfaceVersion() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "InterfaceVersion")
	return result.Value().(int32)
}

func (i *ascomTelescopeImpl) IsPulseGuiding() bool {
	result := oleutil.MustGetProperty(i.dispatch, "IsPulseGuiding")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) RightAscension() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "RightAscension")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) RightAscensionRate() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "RightAscensionRate")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetRightAscensionRate(value float64) {
	oleutil.MustPutProperty(i.dispatch, "RightAscensionRate", value)
}

func (i *ascomTelescopeImpl) SideOfPier() PierSide {
	result := oleutil.MustGetProperty(i.dispatch, "SideOfPier")
	return PierSide(result.Value().(int32))
}

func (i *ascomTelescopeImpl) SetSideOfPier(value PierSide) {
	oleutil.MustPutProperty(i.dispatch, "SideOfPier", int(value))
}

func (i *ascomTelescopeImpl) SiderealTime() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "SiderealTime")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SiteElevation() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "SiteElevation")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetSiteElevation(value float64) {
	oleutil.MustPutProperty(i.dispatch, "SiteElevation", value)
}

func (i *ascomTelescopeImpl) SiteLatitude() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "SiteLatitude")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetSiteLatitude(value float64) {
	oleutil.MustPutProperty(i.dispatch, "SiteLatitude", value)
}

func (i *ascomTelescopeImpl) SiteLongitude() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "SiteLongitude")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetSiteLongitude(value float64) {
	oleutil.MustPutProperty(i.dispatch, "SiteLongitude", value)
}

func (i *ascomTelescopeImpl) SlewSettleTime() int32 {
	result := oleutil.MustGetProperty(i.dispatch, "SlewSettleTime")
	return result.Value().(int32)
}

func (i *ascomTelescopeImpl) SetSlewSettleTime(value int) {
	oleutil.MustPutProperty(i.dispatch, "SlewSettleTime", value)
}

func (i *ascomTelescopeImpl) Slewing() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Slewing")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) TargetDeclination() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "TargetDeclination")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetTargetDeclination(value float64) {
	oleutil.MustPutProperty(i.dispatch, "TargetDeclination", value)
}

func (i *ascomTelescopeImpl) TargetRightAscension() float64 {
	result := oleutil.MustGetProperty(i.dispatch, "TargetRightAscension")
	return result.Value().(float64)
}

func (i *ascomTelescopeImpl) SetTargetRightAscension(value float64) {
	oleutil.MustPutProperty(i.dispatch, "TargetRightAscension", value)
}

func (i *ascomTelescopeImpl) Tracking() bool {
	result := oleutil.MustGetProperty(i.dispatch, "Tracking")
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) SetTracking(value bool) {
	oleutil.MustPutProperty(i.dispatch, "Tracking", value)
}

func (i *ascomTelescopeImpl) TrackingRate() DriveRate {
	result := oleutil.MustGetProperty(i.dispatch, "TrackingRate")
	return DriveRate(result.Value().(int32))
}

func (i *ascomTelescopeImpl) SetTrackingRate(value DriveRate) {
	oleutil.MustPutProperty(i.dispatch, "TrackingRate", int(value))
}

func (i *ascomTelescopeImpl) AbortSlew() {
	oleutil.MustCallMethod(i.dispatch, "AbortSlew")
}

func (i *ascomTelescopeImpl) Action(actionName string, actionParameters string) string {
	result := oleutil.MustCallMethod(i.dispatch, "Action", actionName, actionParameters)
	return result.Value().(string)
}

func (i *ascomTelescopeImpl) CommandBlind(command string, raw bool) {
	oleutil.MustCallMethod(i.dispatch, "CommandBlind", command, raw)
}

func (i *ascomTelescopeImpl) CommandBool(command string, raw bool) bool {
	result := oleutil.MustCallMethod(i.dispatch, "CommandBool", command, raw)
	return result.Value().(bool)
}

func (i *ascomTelescopeImpl) CommandString(command string, raw bool) string {
	result := oleutil.MustCallMethod(i.dispatch, "CommandString", command, raw)
	return result.Value().(string)
}

func (i *ascomTelescopeImpl) Connect() {
	oleutil.MustCallMethod(i.dispatch, "Connect")
}

func (i *ascomTelescopeImpl) DestinationSideOfPier(rightAscension float64, declination float64) PierSide {
	result := oleutil.MustCallMethod(i.dispatch, "DestinationSideOfPier", rightAscension, declination)
	return PierSide(result.Value().(int32))
}

func (i *ascomTelescopeImpl) Disconnect() {
	oleutil.MustCallMethod(i.dispatch, "Disconnect")
}

func (i *ascomTelescopeImpl) FindHome() {
	oleutil.MustCallMethod(i.dispatch, "FindHome")
}

func (i *ascomTelescopeImpl) MoveAxis(axis TelescopeAxis, rate float64) {
	oleutil.MustCallMethod(i.dispatch, "MoveAxis", int(axis), rate)
}

func (i *ascomTelescopeImpl) Park() {
	oleutil.MustCallMethod(i.dispatch, "Park")
}

func (i *ascomTelescopeImpl) PulseGuide(direction GuideDirection, duration int) {
	oleutil.MustCallMethod(i.dispatch, "PulseGuide", int(direction), duration)
}

func (i *ascomTelescopeImpl) SetPark() {
	oleutil.MustCallMethod(i.dispatch, "SetPark")
}

func (i *ascomTelescopeImpl) SetupDialog() {
	oleutil.MustCallMethod(i.dispatch, "SetupDialog")
}

func (i *ascomTelescopeImpl) SlewToAltAz(azimuth float64, altitude float64) {
	oleutil.MustCallMethod(i.dispatch, "SlewToAltAz", azimuth, altitude)
}

func (i *ascomTelescopeImpl) SlewToAltAzAsync(azimuth float64, altitude float64) {
	oleutil.MustCallMethod(i.dispatch, "SlewToAltAzAsync", azimuth, altitude)
}

func (i *ascomTelescopeImpl) SlewToCoordinates(rightAscension float64, declination float64) {
	oleutil.MustCallMethod(i.dispatch, "SlewToCoordinates", rightAscension, declination)
}

func (i *ascomTelescopeImpl) SlewToCoordinatesAsync(rightAscension float64, declination float64) {
	oleutil.MustCallMethod(i.dispatch, "SlewToCoordinatesAsync", rightAscension, declination)
}

func (i *ascomTelescopeImpl) SlewToTarget() {
	oleutil.MustCallMethod(i.dispatch, "SlewToTarget")
}

func (i *ascomTelescopeImpl) SlewToTargetAsync() {
	oleutil.MustCallMethod(i.dispatch, "SlewToTargetAsync")
}

func (i *ascomTelescopeImpl) SyncToAltAz(azimuth float64, altitude float64) {
	oleutil.MustCallMethod(i.dispatch, "SyncToAltAz", azimuth, altitude)
}

func (i *ascomTelescopeImpl) SyncToCoordinates(rightAscension float64, declination float64) {
	oleutil.MustCallMethod(i.dispatch, "SyncToCoordinates", rightAscension, declination)
}

func (i *ascomTelescopeImpl) SyncToTarget() {
	oleutil.MustCallMethod(i.dispatch, "SyncToTarget")
}

func (i *ascomTelescopeImpl) Unpark() {
	oleutil.MustCallMethod(i.dispatch, "Unpark")
}

func (i *ascomTelescopeImpl) WaitForConnectingComplete() {
	ticker := time.NewTicker(500 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if !i.Connecting() {
			return
		}
	}
}

func (i *ascomTelescopeImpl) WaitForPulseGuideComplete() {
	ticker := time.NewTicker(100 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if !i.IsPulseGuiding() {
			return
		}
	}
}

func (i *ascomTelescopeImpl) WaitForSlewComplete() {
	ticker := time.NewTicker(500 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if !i.Slewing() {
			return
		}
	}
}

func NewTelescope(progId string) (Telescope, error) {
	object, err := oleutil.CreateObject(progId)

	if err != nil {
		return nil, err
	}

	dispatch := object.MustQueryInterface(ole.IID_IDispatch)

	return &ascomTelescopeImpl{
		device:   NewAscomDevice(dispatch),
		dispatch: dispatch,
	}, nil
}
