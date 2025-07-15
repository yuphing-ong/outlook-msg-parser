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
	"github.com/saintfish/chardet"

	"github.com/yuphing-ong/outlook-msg-parser/models"
	"golang.org/x/net/html/charset"
	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/transform"
)

const PropsKey = "__properties_version1.0"

// PropertyStreamPrefix is the prefix used for a property stream in the msg binary
const PropertyStreamPrefix = "__substg1.0_"
const RecepientStreamPrefix = "__recip_version1.0_"

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

	res.CalculateFinalBody()

	return res, nil
}

// processEntries iterates through the entries in the mscfb.Reader and processes each entry
func processEntries(doc *mscfb.Reader, res *models.Message, debug bool) error {
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		if debug {
			log.Printf("\n\n-->Processing entry: %s, size: %d, path: %s", entry.Name, entry.Size, entry.Path)
		}
		/*if strings.HasPrefix(entry.Name, "__recip_version1.0") {
			processRecipientStream(entry, res)
		} else if strings.HasPrefix(entry.Name, "__attach_version1.0_") {
			processAttachmentStream(entry, res)
		} else if strings.HasPrefix(entry.Name, "__substg1.0_") {
			processPropertyStream(entry, res, debug)
		} else if entry.Name == "__properties_version1.0" {
			processPropertiesStream(entry, res,debug)
		} else {
			// Handle other types of streams if necessary
		}*/
		if strings.HasPrefix(entry.Name, "__substg1.0_") {
			processPropertyStream(entry, res, debug)
		} else {
			if debug {
				log.Printf("Skipping entry: %s, size: %d, path: %s", entry.Name, entry.Size, entry.Path)
			}
			if entry.Size != 0 {
				entryBytes := make([]byte, entry.Size)
				_, err := entry.Read(entryBytes)
				if debug {
					if err != nil {
						log.Printf("Error reading entry bytes: %v", err)
					} else {
						log.Printf("Entry bytes: %x", entryBytes)
					}
				}
			}
		}

	}
	return nil
}

func processPropertiesStream(entry *mscfb.File, res *models.Message) {
	if entry.Size == 0 {
		//log.Printf("Properties stream %s has size 0", entry.Name)
		return
	}

	// Read the entire entry data
	data := make([]byte, entry.Size)
	_, err := entry.Read(data)
	if err != nil {
		//log.Fatalf("Failed to read properties stream: %v", err)
	}

	// Parse the properties from the data
	offset := 0
	for offset < len(data) {
		// Each property is typically stored with a property tag and value
		if offset+8 > len(data) {
			break
		}

		// Read the property tag (4 bytes for property ID and 4 bytes for property type)
		propTag := binary.LittleEndian.Uint32(data[offset : offset+4])
		propType := binary.LittleEndian.Uint32(data[offset+4 : offset+8])
		offset += 8

		// Determine the length of the property value based on the property type
		var valueLen int
		switch propType {
		case 0x0002: // PT_I2
			valueLen = 2
		case 0x0003: // PT_LONG
			valueLen = 4
		case 0x000B: // PT_BOOLEAN
			valueLen = 2
		case 0x0014: // PT_I8
			valueLen = 8
		case 0x0048: // PT_CLSID
			valueLen = 16
		case 0x00FB: // PT_SVREID
			valueLen = int(binary.LittleEndian.Uint32(data[offset : offset+4]))
			offset += 4
		case 0x1E, 0x1F: // PT_STRING8, PT_UNICODE
			valueLen = int(binary.LittleEndian.Uint32(data[offset : offset+4]))
			offset += 4
		default:
			valueLen = int(binary.LittleEndian.Uint32(data[offset : offset+4]))
			offset += 4
		}

		// Read the property value
		if offset+valueLen > len(data) {
			break
		}
		propValue := data[offset : offset+valueLen]
		offset += valueLen

		// Create a MessageEntryProperty and set the property in the message
		property := models.MessageEntryProperty{
			Class: fmt.Sprintf("%04x", propTag),
			Mapi:  int64(propType),
			Data:  extractDataFromBytes(propValue, propType),
		}

		if property.Class != "0000" {
			res.SetProperties(property)
		}
	}
}

// processSubStorageStream processes a sub-storage stream and iterates through its entries
func processSubStorageStream(entry *mscfb.File, res *models.Message, debug bool) {

	subStorage, err := mscfb.New(entry)
	if err != nil {
		//log.Fatalf("Failed to parse sub-storage: %v", err)
	}
	for subEntry, err := subStorage.Next(); err == nil; subEntry, err = subStorage.Next() {
		if debug {
			log.Printf("Processing sub-entry: %s, size: %d", subEntry.Name, subEntry.Size)
		}
		/*if strings.HasPrefix(subEntry.Name, "__recip_version1.0_") {
			processRecipientStream(subEntry, res)
		} else if strings.HasPrefix(subEntry.Name, "__attach_version1.0_") {
			processAttachmentStream(subEntry, res)
		} else {
			processPropertyStream(subEntry, res, debug)
		}*/
	}
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

func processPropertyStream(entry *mscfb.File, res *models.Message, debug bool) {

	msg := extractMessageProperty(entry)

	if debug {
		log.Printf("***** Processing Property Stream: %+v", msg)
	}

	if len(entry.Path) > 0 && strings.Contains(entry.Path[0], "__recip_version1.0_") {
		// Recipient stream
		processRecipientStream(entry, &msg, res)
	} else {
		if debug {
			log.Printf("Skipping entry path: %s, size: %d, path: %s", entry.Name, entry.Size, entry.Path)
		}
	}

	res.SetProperties(msg)

}

func processRecipientStream(entry *mscfb.File, msg *models.MessageEntryProperty, res *models.Message) {

	// Determine recipient type and email address

	//log.Printf("############# Recipient Data: %v", msg.Data)

	recipientIDStr := entry.Path[0][len("__recip_version1.0_#"):]

	/*switch recipientID {
	case "00000000":
		msg.Class = "RecipientType"
		msg.Mapi = 0x0C15
		msg.Data = "Originator"
	case "00000001":
		msg.Class = "RecipientType"
		msg.Mapi = 0x0C15
		msg.Data = "To"
	case "00000002":
		msg.Class = "RecipientType"
		msg.Mapi = 0x0C15
		msg.Data = "CC"
	case "00000003":
		msg.Class = "RecipientType"
		msg.Mapi = 0x0C15
		msg.Data = "BCC"
	}*/

	// If recipient ID is not 0, set it as TO

	recipientID, err := strconv.Atoi(recipientIDStr)
	if err != nil {
		return
	}
	if recipientID != 0 {
		res.LastRecipient = recipientID
	}
	////log.Printf("##################>Parsed Recipient: %+v", msg)
}

// processAttachmentStream processes an attachment stream and sets the attachment properties in the Message instance
func processAttachmentStream(entry *mscfb.File, msg *models.Message) {
	// Iterate through the properties in the attachment stream
	/*for {
		prop, err := entry.Next()
		if err == mscfb.EOF {
			break
		}
		if err != nil {
			log.Fatalf("Failed to read property: %v", err)
		}

		// Parse the property name and extract data
		property := parseEntryName(prop)
		data := extractData(prop, property)
		property.Data = data

		// Set the properties of the message
		msg.SetProperties(property)
	}

	// Print the parsed attachment for manual verification
	log.Printf("Parsed Attachment: %+v", msg)*/
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
		// PT_STRING8: Robust charset detection and decoding for special icons/emojis
		bytes2 := make([]byte, entry.Size)
		entry.Read(bytes2)
		// Use chardet to detect encoding
		detector := chardet.NewTextDetector()
		result, err := detector.DetectBest(bytes2)
		var decoded string
		if err == nil && result != nil {
			var enc encoding.Encoding
			switch strings.ToLower(result.Charset) {
			case "windows-1252":
				enc = charmap.Windows1252
			case "iso-8859-1":
				enc = charmap.ISO8859_1
			case "utf-8":
				enc = nil // no transform needed
			default:
				enc, _ = charset.Lookup(result.Charset)
			}
			if enc != nil {
				reader := transform.NewReader(bytes.NewReader(bytes2), enc.NewDecoder())
				resu, err := ioutil.ReadAll(reader)
				if err == nil {
					decoded = string(resu)
				}
			} else {
				// Assume UTF-8
				decoded = string(bytes2)
			}
		} else {
			// Fallback: try Windows-1252, then ISO-8859-1, then UTF-8
			read, err := charset.NewReaderLabel("windows-1252", bytes.NewReader(bytes2))
			if err != nil {
				read, err = charset.NewReaderLabel("iso-8859-1", bytes.NewReader(bytes2))
			}
			if err == nil && read != nil {
				resu, _ := ioutil.ReadAll(read)
				decoded = string(resu)
			} else {
				decoded = string(bytes2)
			}
		}
		return decoded
	case 0x1f:
		// PT_UNICODE: A null-terminated Unicode string (UTF-16LE)
		bytes2 := make([]byte, entry.Size)
		entry.Read(bytes2)
		u16s := make([]uint16, len(bytes2)/2)
		for i := 0; i < len(bytes2)-1; i += 2 {
			u16s[i/2] = binary.LittleEndian.Uint16(bytes2[i : i+2])
		}
		// Use unicode/utf16 for robust decoding
		importU16 := func(u []uint16) []rune {
			// Use utf16.Decode if available
			// fallback: convert directly
			runes := make([]rune, len(u))
			for i, v := range u {
				runes[i] = rune(v)
			}
			return runes
		}
		// Try to use utf16.Decode if available
		var runes []rune
		// This import is only for demonstration, in real code, import "unicode/utf16" at the top
		// runes = utf16.Decode(u16s)
		// For now, fallback to direct conversion
		runes = importU16(u16s)
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

func extractDataFromBytes(data []byte, propType uint32) interface{} {
	switch propType {
	case 0x0002: // PT_I2
		return int16(binary.LittleEndian.Uint16(data))
	case 0x0003: // PT_LONG
		return int32(binary.LittleEndian.Uint32(data))
	case 0x000B: // PT_BOOLEAN
		return binary.LittleEndian.Uint16(data) != 0
	case 0x0014: // PT_I8
		return int64(binary.LittleEndian.Uint64(data))
	case 0x0048: // PT_CLSID
		return fmt.Sprintf("%08x-%04x-%04x-%04x-%012x", binary.LittleEndian.Uint32(data[0:4]), binary.LittleEndian.Uint16(data[4:6]), binary.LittleEndian.Uint16(data[6:8]), binary.BigEndian.Uint16(data[8:10]), data[10:16])
	case 0x1E: // PT_STRING8
		return string(data)
	case 0x1F: // PT_UNICODE
		runes := make([]rune, len(data)/2)
		for i := 0; i < len(data)-1; i += 2 {
			ch := (int)(data[i+1])
			cl := (int)(data[i]) & 0xff
			runes[i/2] = (rune)((ch << 8) + cl)
		}
		return string(runes)
	default:
		return data
	}
}
