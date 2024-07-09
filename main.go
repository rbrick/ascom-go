package main

import (
	"log"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func init() {
	ole.CoInitialize(0)
}

func main() {
	chooser, err := NewChooser()

	if err != nil {
		log.Fatalln(err)
	}

	chooser.SetDeviceType(TelescopeDevice)
	chosenDevice, err := chooser.Choose("")

	if err != nil {
		log.Fatalln(err)
	}

	telescopeObject, err := oleutil.CreateObject(chosenDevice)

	if err != nil {
		log.Fatalln(err)
	}

	dispatch := telescopeObject.MustQueryInterface(ole.IID_IDispatch)

	driver := NewAscomDriver(dispatch)

	driver.Connect()

	log.Println(
		oleutil.MustGetProperty(dispatch, "AtHome").Value().(bool))

	driver.Disconnect()
}
