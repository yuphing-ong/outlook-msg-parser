package msgparser

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io/ioutil"
	"log"
	"math"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/richardlehane/mscfb"
	"golang.org/x/net/html/charset"

	"outlook-msg-parser/models"
)

const PropsKey = "__properties_version1.0"

// PropertyStreamPrefix is the prefix used for a property stream in the msg binary
const PropertyStreamPrefix = "__substg1.0_"

// ReplyToRegExp is a regex to extract the reply to header
const ReplyToRegExp = "^Reply-To:\\s*(?:<?(?<nameOrAddress>.*?)>?)?\\s*(?:<(?<address>.*?)>)?$"

// ParseMsgFileWithDebug parses the msg file with debug information
func ParseMsgFileWithDebug(file string) (res *models.Message, err error) {
	return parseMsgFile(file, true)
}

// ParseMsgFile parses the msg file and sets the properties
func ParseMsgFile(file string) (res *models.Message, err error) {
	return parseMsgFile(file, false)
}

// parseMsgFile is the internal function that parses the msg file and sets the properties
func parseMsgFile(file string, debug bool) (res *models.Message, err error) {
	res = &models.Message{}
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	doc, err := mscfb.New(f)
	if err != nil {
		return nil, err
	}
	err = processEntries(doc, res, debug)
	if err != nil {
		return nil, err
	}
	return res, nil
}

// processEntries iterates through the entries in the mscfb.Reader and processes each entry
func processEntries(doc *mscfb.Reader, res *models.Message, debug bool) error {
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		if strings.HasPrefix(entry.Name, PropertyStreamPrefix) {
			msg := extractMessageProperty(entry)
			res.SetProperties(msg)
			if debug {
				fmt.Println("Entry:", entry, "Class:", msg.Class, "Mapi:", msg.Mapi, "Data:", msg.Data)
			}
		}
	}
	return nil
}

// extractMessageProperty processes an entry and returns a MessageEntryProperty
func extractMessageProperty(entry *mscfb.File) models.MessageEntryProperty {
	analysis := parseEntryName(entry)
	data := extractData(entry, analysis)

	messageProperty := models.MessageEntryProperty{
		Class: analysis.Class,
		Mapi:  analysis.Mapi,
		Data:  data,
	}
	return messageProperty
}

// extractData extracts the data from the entry based on the analysis result
func extractData(entry *mscfb.File, info models.MessageEntryProperty) interface{} {
	if info.Class == "" {
		return "class null"
	}
	mapi := info.Mapi
	switch mapi {
	case -1:
		return "-1"
	case 0x1e:
		// PT_STRING8: A null-terminated 8-bit character string
		bytes2 := make([]byte, entry.Size)
		entry.Read(bytes2)
		read, _ := charset.NewReader(bytes.NewReader(bytes2), "ISO-8859-1")
		if read != nil {
			resu, _ := ioutil.ReadAll(read)
			return string(resu)
		}
	case 0x1f:
		// PT_UNICODE: A null-terminated Unicode string
		bytes2 := make([]byte, entry.Size)
		entry.Read(bytes2)
		runes := make([]rune, len(bytes2)/2)
		c := 0
		for i := 0; i < len(bytes2)-1; i += 2 {
			ch := (int)(bytes2[i+1])
			cl := (int)(bytes2[i]) & 0xff
			runes[c] = (rune)((ch << 8) + cl)
			c++
		}
		return string(runes)
	case 0x102:
		// PT_BINARY: A binary value
		bytes2 := make([]byte, entry.Size)
		entry.Read(bytes2)
		return bytes2
	case 0x40:
		// PT_SYSTIME: A 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601
		bytes := make([]byte, entry.Size)
		entry.Read(bytes)
		if len(bytes) > 0 {
			buf := bytes[:8]
			a := binary.LittleEndian.Uint64(buf)
			a /= 10000
			a -= 11644473600000
			return time.Unix(0, int64(a)*int64(time.Millisecond)).String()
		}
	case 0x0002:
		// PT_I2: A 16-bit integer
		bytes := make([]byte, 2)
		entry.Read(bytes)
		return int16(binary.LittleEndian.Uint16(bytes))
	case 0x0003:
		// PT_LONG: A 32-bit integer
		bytes := make([]byte, 4)
		entry.Read(bytes)
		return int32(binary.LittleEndian.Uint32(bytes))
	case 0x0004:
		// PT_R4: A 4-byte floating point number
		bytes := make([]byte, 4)
		entry.Read(bytes)
		return math.Float32frombits(binary.LittleEndian.Uint32(bytes))
	case 0x0005:
		// PT_DOUBLE: An 8-byte floating point number
		bytes := make([]byte, 8)
		entry.Read(bytes)
		return math.Float64frombits(binary.LittleEndian.Uint64(bytes))
	case 0x0006:
		// PT_CURRENCY: A 64-bit integer representing a currency value
		bytes := make([]byte, 8)
		entry.Read(bytes)
		return int64(binary.LittleEndian.Uint64(bytes))
	case 0x0007:
		// PT_APPTIME: A double representing the number of days since December 30, 1899
		bytes := make([]byte, 8)
		entry.Read(bytes)
		return math.Float64frombits(binary.LittleEndian.Uint64(bytes))
	case 0x000B:
		// PT_BOOLEAN: A Boolean value
		bytes := make([]byte, 2)
		entry.Read(bytes)
		return binary.LittleEndian.Uint16(bytes) != 0
	case 0x0014:
		// PT_I8: A 64-bit integer
		bytes := make([]byte, 8)
		entry.Read(bytes)
		return int64(binary.LittleEndian.Uint64(bytes))
	case 0x0048:
		// PT_CLSID: A GUID
		bytes := make([]byte, 16)
		entry.Read(bytes)
		return fmt.Sprintf("%08x-%04x-%04x-%04x-%012x", binary.LittleEndian.Uint32(bytes[0:4]), binary.LittleEndian.Uint16(bytes[4:6]), binary.LittleEndian.Uint16(bytes[6:8]), binary.BigEndian.Uint16(bytes[8:10]), bytes[10:16])
	case 0x00FB:
		// PT_SVREID: A server entry identifier
		bytes := make([]byte, entry.Size)
		entry.Read(bytes)
		return bytes
	case 0x1002:
		// PT_MV_I2: A multiple-value 16-bit integer
		count := entry.Size / 2
		values := make([]int16, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 2)
			entry.Read(bytes)
			values[i] = int16(binary.LittleEndian.Uint16(bytes))
		}
		return values
	case 0x1003:
		// PT_MV_LONG: A multiple-value 32-bit integer
		count := entry.Size / 4
		values := make([]int32, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 4)
			entry.Read(bytes)
			values[i] = int32(binary.LittleEndian.Uint32(bytes))
		}
		return values
	case 0x1004:
		// PT_MV_R4: A multiple-value 4-byte floating point number
		count := entry.Size / 4
		values := make([]float32, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 4)
			entry.Read(bytes)
			values[i] = math.Float32frombits(binary.LittleEndian.Uint32(bytes))
		}
		return values
	case 0x1005:
		// PT_MV_DOUBLE: A multiple-value 8-byte floating point number
		count := entry.Size / 8
		values := make([]float64, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 8)
			entry.Read(bytes)
			values[i] = math.Float64frombits(binary.LittleEndian.Uint64(bytes))
		}
		return values
	case 0x1006:
		// PT_MV_CURRENCY: A multiple-value currency value
		count := entry.Size / 8
		values := make([]int64, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 8)
			entry.Read(bytes)
			values[i] = int64(binary.LittleEndian.Uint64(bytes))
		}
		return values
	case 0x1007:
		// PT_MV_APPTIME: A multiple-value application time
		count := entry.Size / 8
		values := make([]float64, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 8)
			entry.Read(bytes)
			values[i] = math.Float64frombits(binary.LittleEndian.Uint64(bytes))
		}
		return values
	case 0x1040:
		// PT_MV_SYSTIME: A multiple-value system time
		count := entry.Size / 8
		values := make([]time.Time, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 8)
			entry.Read(bytes)
			a := binary.LittleEndian.Uint64(bytes)
			a /= 10000
			a -= 11644473600000
			values[i] = time.Unix(0, int64(a)*int64(time.Millisecond))
		}
		return values
	case 0x101E:
		// PT_MV_STRING8: A multiple-value null-terminated 8-bit character string
		bytes := make([]byte, entry.Size)
		entry.Read(bytes)
		strs := strings.Split(string(bytes), "\x00")
		return strs[:len(strs)-1] // Remove the last empty string after the final null terminator
	case 0x101F:
		// PT_MV_UNICODE: A multiple-value null-terminated Unicode string
		bytes := make([]byte, entry.Size)
		entry.Read(bytes)
		runes := make([]rune, len(bytes)/2)
		c := 0
		for i := 0; i < len(bytes)-1; i += 2 {
			ch := (int)(bytes[i+1])
			cl := (int)(bytes[i]) & 0xff
			runes[c] = (rune)((ch << 8) + cl)
			c++
		}
		strs := strings.Split(string(runes), "\x00")
		return strs[:len(strs)-1] // Remove the last empty string after the final null terminator
	case 0x1102:
		// PT_MV_BINARY: A multiple-value binary value
		bytes := make([]byte, entry.Size)
		entry.Read(bytes)
		return bytes
	case 0x1048:
		// PT_MV_CLSID: A multiple-value GUID
		count := entry.Size / 16
		values := make([]string, count)
		for i := 0; i < int(count); i++ {
			bytes := make([]byte, 16)
			entry.Read(bytes)
			values[i] = fmt.Sprintf("%08x-%04x-%04x-%04x-%012x", binary.LittleEndian.Uint32(bytes[0:4]), binary.LittleEndian.Uint16(bytes[4:6]), binary.LittleEndian.Uint16(bytes[6:8]), binary.BigEndian.Uint16(bytes[8:10]), bytes[10:16])
		}
		return values
	case 0x10FB:
		// PT_MV_SVREID: A multiple-value server entry identifier
		bytes := make([]byte, entry.Size)
		entry.Read(bytes)
		return bytes
	default:
		return fmt.Sprintf("default mapi: %x", mapi)
	}
	return ""
}

func parseEntryName(entry *mscfb.File) models.MessageEntryProperty {
	name := entry.Name
	res := models.MessageEntryProperty{}
	if strings.HasPrefix(name, PropertyStreamPrefix) {
		val := name[len(PropertyStreamPrefix):]
		if len(val) < 8 {
			log.Println("Invalid entry name length")
			return res
		}
		class := val[0:4]
		typeEntry := val[4:8]
		mapi, err := strconv.ParseInt(typeEntry, 16, 64)
		if err != nil {
			log.Printf("Error parsing MAPI type: %v", err)
			return res
		}
		res = models.MessageEntryProperty{
			Class: class,
			Mapi:  mapi,
		}
	}
	return res
}
