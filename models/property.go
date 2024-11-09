package models

// MessageEntryProperty holds information about a property of a message entry in the msg file
type MessageEntryProperty struct {
	Class string
	Mapi  int64
	Data  interface{}
}
