package models

import (
	"encoding/binary"
	"log"
	"strconv"
	"strings"
	"time"
	"unicode"

	"regexp"

	"github.com/richardlehane/mscfb"
)

// Message is a struct that holds a structered result of parsing the entry
type Message struct {
	MessageClass            string                // PR_MESSAGE_CLASS
	MessageID               string                // PR_INTERNET_MESSAGE_ID
	Subject                 string                // PR_SUBJECT
	FromEmail               string                // PR_SENDER_EMAIL_ADDRESS
	FromName                string                // PR_SENDER_NAME
	ToDisplay               string                // PR_DISPLAY_TO
	To                      string                // PR_DISPLAY_TO
	CCDisplay               string                // PR_DISPLAY_CC
	BCCDisplay              string                // PR_DISPLAY_BCC
	CC                      string                // PR_DISPLAY_CC
	BCC                     string                // PR_DISPLAY_BCC
	BodyPlainText           string                // PR_BODY
	BodyHTML                string                // PR_HTML
	ConvertedBodyHTML       string                // The body in HTML format (converted from RTF)
	Headers                 string                // Email headers (if available)
	Date                    time.Time             // PR_MESSAGE_DELIVERY_TIME
	ClientSubmitTime        time.Time             // PR_CLIENT_SUBMIT_TIME
	CreationDate            time.Time             // PR_CREATION_TIME
	LastModificationDate    time.Time             // PR_LAST_MODIFICATION_TIME
	Attachments             []Attachment          // Attachments
	Properties              map[int64]interface{} // Other properties
	TransportMessageHeaders string                // Message Headers
	Address                 []string              // Email Address
	LastRecipient           int                   // Last recipient of the message

	bodyCandidates []string
	htmlCandidates []string
}

type Attachment struct {
	Name string
	// Add other relevant fields as needed
}

const AttachmentPrefix = "__attach_"

// SetProperties sets the message properties and collects body/html candidates
func (res *Message) SetProperties(msgProps MessageEntryProperty) {
	name := msgProps.Class
	data := msgProps.Data
	if res.Properties == nil {
		res.Properties = make(map[int64]interface{}, 2)
	}
	class, err := strconv.ParseInt(name, 16, 32)
	if err != nil {
		log.Print("Parse Error")
	}

	// Check if the entry is an attachment and handle it separately
	if strings.HasPrefix(name, AttachmentPrefix) {
		//res.HandleAttachment(entry)
		return
	}

	// --- Robust body and HTML property handling ---
	// In the switch, collect all body and HTML candidates
	switch class {
	case 0x1a:
		// PR_MESSAGE_CLASS: The message class of the message
		if res.MessageClass == "" {
			res.MessageClass = data.(string)
		}

	case 0x1035:
		// PR_INTERNET_MESSAGE_ID: The Internet message ID of the message
		if res.MessageID == "" {
			res.MessageID = data.(string)
		}

	case 0x37:
		// PR_SUBJECT: The subject of the message
		if res.Subject == "" {
			res.Subject = data.(string)
		}

	case 0xe1d:
		// PR_NORMALIZED_SUBJECT: The normalized subject of the message
		if res.Subject == "" {
			res.Subject = data.(string)
		}

	case 0xc1f:
		// PR_SENDER_EMAIL_ADDRESS: The email address of the sender
		if isValidEmail(data.(string)) {
			if res.FromEmail == "" {
				res.FromEmail = data.(string)
			} else if !strings.Contains(res.FromEmail, data.(string)) {
				res.FromEmail = data.(string) + ", " + res.FromEmail
			}
		}
	case 0x65:
		// PR_SENT_REPRESENTING_EMAIL_ADDRESS: The email address of the user represented by the sender
		if res.FromEmail == "" && isValidEmail(data.(string)) {
			res.FromEmail = data.(string)
		} else if !strings.Contains(res.FromEmail, data.(string)) {
			res.FromEmail = data.(string) + ", " + res.FromEmail
		}
	case 0x3ffa:
		// PR_LAST_MODIFIER_NAME: The name of the last user to modify the message
		if res.FromName == "" {
			res.FromName = data.(string)
		}

	case 0x1000, 0x3ff9, 0x65e0, 0x65e2, 0xff9, 0x120b:
		// PR_BODY: The plain text body of the message
		if v, ok := data.([]uint8); ok {
			if cleaned, ok := CleanAndAcceptBodyCandidate(string(v), 10); ok {
				res.bodyCandidates = append(res.bodyCandidates, cleaned)
			}
		} else if v, ok := data.(string); ok {
			if cleaned, ok := CleanAndAcceptBodyCandidate(v, 10); ok {
				res.bodyCandidates = append(res.bodyCandidates, cleaned)
			}
		}
	case 0x1001, 0x1013, 0x3ffb, 0x65e1, 0x65e3, 0x5ff7, 0xc25, 0xf03:
		// PR_BODY_HTML: The HTML body of the message
		if v, ok := data.([]uint8); ok {
			if cleaned, ok := CleanAndAcceptBodyCandidate(string(v), 10); ok {
				res.htmlCandidates = append(res.htmlCandidates, cleaned)
			}
		} else if v, ok := data.(string); ok {
			if cleaned, ok := CleanAndAcceptBodyCandidate(v, 10); ok {
				res.htmlCandidates = append(res.htmlCandidates, cleaned)
			}
		}

	case 0x3007:
		// PR_CREATION_TIME: The creation time of the message
		if res.CreationDate.IsZero() {
			res.CreationDate = data.(time.Time)
		}

	case 0x3008:
		// PR_LAST_MODIFICATION_TIME: The last modification time of the message
		if res.LastModificationDate.IsZero() {
			res.LastModificationDate = data.(time.Time)
		}

	case 0xe06:
		// PR_CLIENT_SUBMIT_TIME: The client submit time of the message
		if res.ClientSubmitTime.IsZero() {
			res.ClientSubmitTime = data.(time.Time)
		}

	case 0xe0f:
		// PR_MESSAGE_DELIVERY_TIME: The delivery time of the message
		if res.Date.IsZero() {
			res.Date = data.(time.Time)
		}

	case 0x0002:
		// PR_IMPORTANCE: The importance level of the message
		if intData, ok := data.([]uint8); ok {
			res.Properties[class] = intData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x0003:
		// PR_PRIORITY: The priority level of the message
		if intData, ok := data.([]uint8); ok {
			res.Properties[class] = intData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x0004:
		// PR_PRIORITY: The priority level of the message
		if floatData, ok := data.([]uint8); ok {
			res.Properties[class] = floatData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x1002:
		// PR_REPORT_TEXT: Text of a report
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = string(byteData)
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x1008:
		// PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED: Indicates if a delivery report is requested
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = byteData[0] != 0
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x1009:
		// PR_READ_RECEIPT_REQUESTED: Indicates if a read receipt is requested
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = byteData[0] != 0
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x1014:
		// PR_RTF_SYNC_BODY_CRC: CRC of the RTF body
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = int32(binary.LittleEndian.Uint32(byteData))
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x1015:
		// PR_RTF_SYNC_BODY_COUNT: Count of the RTF body
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = int32(binary.LittleEndian.Uint32(byteData))
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x003b:
		// PR_ENTRYID: Entry identifier
		if binData, ok := data.([]byte); ok {
			res.Properties[class] = binData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x003f:
		// PR_OBJECT_TYPE: Type of the object
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = int32(binary.LittleEndian.Uint32(byteData))
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x0041:
		// PR_ICON: Icon of the message
		if binData, ok := data.([]byte); ok {
			res.Properties[class] = binData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x0051:
		// PR_ACCESS: Access level of the message
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = int32(binary.LittleEndian.Uint32(byteData))
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x0071:
		// PR_ACCESS_LEVEL: Access level of the message
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = int32(binary.LittleEndian.Uint32(byteData))
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x0c19:
		// PR_SENDER_ENTRYID: Entry identifier of the sender
		if binData, ok := data.([]byte); ok {
			res.Properties[class] = binData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x0c1d:
		// PR_SENT_REPRESENTING_ENTRYID: Entry identifier of the user represented by the sender
		if binData, ok := data.([]byte); ok {
			res.Properties[class] = binData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x300b:
		// PR_HASATTACH: Indicates if the message has attachments
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = byteData[0] != 0
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0xe04, 0x800d:
		// PR_DISPLAY_TO: The display names of the primary (To) recipients
		if byteData, ok := data.([]uint8); ok {
			if res.ToDisplay == "" {
				res.ToDisplay = string(byteData)
			}
		} else if strData, ok := data.(string); ok {
			if res.ToDisplay == "" {
				res.ToDisplay = strData
			}
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0xe03, 0x800e:
		// PR_DISPLAY_CC: The display names of the carbon copy (CC) recipients

		if byteData, ok := data.([]uint8); ok {
			if res.CCDisplay == "" {
				res.CCDisplay = string(byteData)
			}
		} else if strData, ok := data.(string); ok {
			if res.CCDisplay == "" {
				res.CCDisplay = strData
			}
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0xe02, 0x800f:
		// PR_DISPLAY_BCC: The display names of the blind carbon copy (BCC) recipients

		if byteData, ok := data.([]uint8); ok {
			if res.BCCDisplay == "" {
				res.BCCDisplay = string(byteData)
			}
		} else if strData, ok := data.(string); ok {
			if res.BCCDisplay == "" {
				res.BCCDisplay = strData
			}
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x8002:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		if strData, ok := data.([]string); ok {
			res.Properties[class] = strData
		} else if strData, ok := data.(string); ok {
			res.Properties[class] = strData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x0ff6:
		// PR_CONVERSATION_TOPIC: Conversation topic
		if res.Properties[class] == nil {
			if byteData, ok := data.([]uint8); ok {
				res.Properties[class] = string(byteData)
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x0fff:
		// PR_CONVERSATION_INDEX: Conversation index
		if binData, ok := data.([]byte); ok {
			res.Properties[class] = binData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

		// Documented but not implemented properties
	case 0x1005:
		// PR_BODY_CONTENT_LOCATION: Content location of the body
		// Not implemented

	case 0x1006:
		// PR_BODY_CONTENT_ID: Content ID of the body
		// Not implemented

	case 0x1007:
		// PR_BODY_CONTENT_TYPE: Content type of the body
		// Not implemented

	case 0x100b:
		// PR_BODY_ENCODING: Encoding of the body
		// Not implemented

	case 0x100c:
		// PR_BODY_SIZE: Size of the body
		// Not implemented

	case 0x100d:
		// PR_BODY_TAG: Tag of the body
		// Not implemented

	case 0x100f:
		// PR_BODY_TYPE: Type of the body
		// Not implemented

	case 0x1011:
		// PR_BODY_CHARSET: Charset of the body
		// Not implemented

	case 0x1016:
		// PR_BODY_LANGUAGE: Language of the body
		// Not implemented

	case 0x1017:
		// PR_BODY_SUBTYPE: Subtype of the body
		// Not implemented

	case 0x1018:
		// PR_BODY_TRANSFER_ENCODING: Transfer encoding of the body
		// Not implemented

	case 0x1019:
		// PR_BODY_DISPOSITION: Disposition of the body
		// Not implemented

	case 0x101a:
		// PR_BODY_DISPOSITION_TYPE: Disposition type of the body
		// Not implemented

	case 0x101b:
		// PR_BODY_DISPOSITION_PARAMS: Disposition parameters of the body
		// Not implemented

	case 0x101c:
		// PR_BODY_DISPOSITION_FILENAME: Disposition filename of the body
		// Not implemented

	case 0x101e:
		// PR_BODY_DISPOSITION_CREATION_DATE: Disposition creation date of the body
		// Not implemented

	case 0x43:
		// PR_BODY_DISPOSITION_MODIFICATION_DATE: Disposition modification date of the body
		// Not implemented

	case 0x52:
		// PR_BODY_DISPOSITION_READ_DATE: Disposition read date of the body
		// Not implemented

	case 0xe0b:
		// PR_BODY_CRC: CRC of the message body
		// Not implemented

	case 0xe4b:
		// PR_RTF_SYNC_BODY_CRC: CRC of the RTF body
		// Not implemented

	case 0xe4c:
		// PR_RTF_SYNC_BODY_COUNT: Count of the RTF body
		// Not implemented

	case 0xe58:
		// PR_RTF_SYNC_BODY_TAG: Tag of the RTF body
		// Not implemented

	case 0xe59:
		// PR_RTF_SYNC_BODY_TAG: Tag of the RTF body
		// Not implemented

	case 0x3013:
		// PR_CREATION_TIME: Creation time of the message
		// Not implemented

	case 0x3014:
		// PR_LAST_MODIFICATION_TIME: Last modification time of the message
		// Not implemented

	case 0x8000:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		// Not implemented

	case 0x8007:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		// Not implemented

	case 0x8008:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		// Not implemented

	case 0x800b:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		// Not implemented

	case 0x802c:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		// Not implemented

	case 0x802e:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		// Not implemented

	case 0x4099:
		// PR_MESSAGE_FLAGS: Flags indicating the status or attributes of the message
		if intData, ok := data.(int32); ok {
			res.Properties[class] = intData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}
	case 0x1003:
		// PR_IMPORTANCE: The importance level of the message
		if intData, ok := data.([]uint8); ok {
			res.Properties[class] = intData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x1004:
		// PR_PRIORITY: The priority level of the message
		if intData, ok := data.([]uint8); ok {
			res.Properties[class] = intData
		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x007D:
		// PR_TRANSPORT_MESSAGE_HEADERS: Transport message headers
		if res.TransportMessageHeaders == "" {
			if byteData, ok := data.([]uint8); ok {
				res.TransportMessageHeaders = string(byteData)
			} else if strData, ok := data.(string); ok {
				res.TransportMessageHeaders = strData
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}

	case 0x3003, 0xC025, 0x39FE:
		// PR_EMAIL_ADDRESS - PR_SMTP_ADDRES
		if byteData, ok := data.([]uint8); ok {
			address := string(byteData)
			if isValidEmail(address) {
				res.Address = append(res.Address, string(byteData))

				if res.LastRecipient == 0 {
					// Add the new address to TO
					res.To = res.To + address + "; "
				} else if res.LastRecipient == 1 {
					// Add the new address to CC
					res.CC = res.CC + address + "; "
				} else if res.LastRecipient == 2 {
					// Add the new address to BCC
					res.BCC = res.BCC + address + "; "
				}
			}
		} else if strData, ok := data.(string); ok {
			address := strData
			if isValidEmail(address) {
				res.Address = append(res.Address, strData)

				// Recipient ID  seems to not be present so we will copy all of them a CC

				if !strings.Contains(res.To, strData) {
					res.To = res.To + strData + "; "
				}

				/*if res.LastRecipient == 0 {
					// Add the new address to TO
					res.To = res.To + strData + "; "
				} else if res.LastRecipient == 1 {
					// Add the new address to CC
					res.CC = res.CC + strData + "; "
				} else if res.LastRecipient == 2 {
					// Add the new address to BCC
					res.BCC = res.BCC + "; "
				}*/
			}

		} else {
			log.Printf("Unexpected type for property %x: %T", class, data)
		}

	case 0x0C24:
		// PR_SENT_REPRESENTING_ADDRTYPE

	// Handle attachment file name properties (PR_ATTACH_FILENAME, PR_ATTACH_LONG_FILENAME)
	case 0x3701, 0x3702, 0x371d:
		if byteData, ok := data.([]uint8); ok {
			res.Properties[class] = string(byteData)
		} else if strData, ok := data.(string); ok {
			res.Properties[class] = strData
		}

	// Handle PR_ATTACH_MIME_TAG (MIME tag for attachment)
	case 0x8004:
		if strSlice, ok := data.([]string); ok {
			res.Properties[class] = strings.Join(strSlice, ", ")
		} else if strData, ok := data.(string); ok {
			res.Properties[class] = strData
		}

	// Handle string properties that may come as []uint8
	case 0x1010, 0x1012, 0x101d, 0x3019, 0x301b, 0x8021:
		if byteData, ok := data.([]uint8); ok {
			res.Properties[class] = string(byteData)
		} else if strData, ok := data.(string); ok {
			res.Properties[class] = strData
		}

	// Handle additional properties based on log
	case 0x100a:
		// PR_BODY_HTML_ALT or similar: treat as string if possible
		if byteData, ok := data.([]uint8); ok {
			res.Properties[class] = string(byteData)
		} else if strData, ok := data.(string); ok {
			res.Properties[class] = strData
		}
	case 0x8005:
		// Possibly PR_ATTACH_CONTENT_ID or similar, can be []string or string
		if strSlice, ok := data.([]string); ok {
			res.Properties[class] = strings.Join(strSlice, ", ")
		} else if strData, ok := data.(string); ok {
			res.Properties[class] = strData
		}
	case 0x8011, 0x8025, 0x802d:
		// Unknown, but treat []uint8 as string
		if byteData, ok := data.([]uint8); ok {
			res.Properties[class] = string(byteData)
		} else if strData, ok := data.(string); ok {
			res.Properties[class] = strData
		}

	default:
		// Store other properties in the Properties map
		if class == 0 {
			return
		}
		if _, exists := res.Properties[class]; !exists {
			if strData, ok := data.(string); ok {
				res.Properties[class] = strData
			} else {
				log.Printf("Unexpected type for property %x: %T", class, data)
			}
		}
	}
}

// CleanAndAcceptBodyCandidate cleans the input and returns it if it is a valid body candidate.
func CleanAndAcceptBodyCandidate(input string, minLen int) (string, bool) {
	cleaned := strings.TrimSpace(input)
	if len(cleaned) < minLen {
		return "", false // too short
	}
	// Filter Exchange/X.500 address patterns
	if strings.Contains(cleaned, "/O=") && strings.Contains(cleaned, "/CN=") {
		return "", false
	}
	// Filter if more than 40% of runes are not letters, digits, space, or common punctuation
	nonLetter := 0
	total := 0
	var output strings.Builder
	for _, r := range cleaned {
		if unicode.IsPrint(r) || r == '\n' || r == '\r' || r == '\t' {
			output.WriteRune(r)
			total++
			if !(unicode.IsLetter(r) || unicode.IsDigit(r) || unicode.IsSpace(r) || strings.ContainsRune(",.;:!?()[]{}-_'\"/@#%&$*", r)) {
				nonLetter++
			}
		}
	}
	final := output.String()
	if total == 0 || float64(nonLetter)/float64(total) > 0.4 {
		return "", false
	}
	// Filter if contains too many replacement chars (�) relative to length
	replacementCount := strings.Count(final, "�")
	length := len(final)
	if length > 0 {
		// Allow up to 20% for very short strings, 10% for medium, 5% for long
		var maxAllowed int
		switch {
		case length < 50:
			maxAllowed = length / 5 // 20%
		case length < 500:
			maxAllowed = length / 10 // 10%
		default:
			maxAllowed = length / 20 // 5%
		}
		if replacementCount > maxAllowed && replacementCount > 2 {
			return "", false
		}
	}
	return final, true
}

func isValidEmail(email string) bool {
	re := regexp.MustCompile(`^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z]{2,}$`)
	if (len(email) > 0) && (len(email) > 60) {
		return false
	}
	return re.MatchString(email)
}

// ValidateEmailList validates a comma-separated list of email addresses
func ValidateEmailList(emailList string) bool {
	emails := strings.Split(emailList, ",")
	for _, email := range emails {
		email = strings.TrimSpace(email)
		if !isValidEmail(email) {
			return false
		}
	}
	return true
}

// HandleAttachment processes and stores attachment information
func (res *Message) HandleAttachment(entry *mscfb.File) {
	// Implement attachment handling logic here
	// For example, store the attachment in a separate list or map
	attachment := Attachment{
		Name: entry.Name,
		// Add other relevant fields and processing as needed
	}
	res.Attachments = append(res.Attachments, attachment)
}

// CalculateFinalBody selects the best HTML or plain text body from the candidate arrays.
//
// Reasoning:
// - If any HTML candidates are present, the longest one is chosen as BodyHTML (HTML is preferred for fidelity).
// - If no HTML is present but plain text candidates exist, the longest one is chosen as BodyPlainText.
// - If neither is present, both fields remain empty.
// This approach ensures the richest available content is used, and avoids short/empty/partial bodies.
func (res *Message) CalculateFinalBody() {
	if len(res.htmlCandidates) > 0 {
		// Pick the longest valid HTML
		best := res.htmlCandidates[0]
		for _, h := range res.htmlCandidates[1:] {
			if len(h) > len(best) {
				best = h
			}
		}
		res.BodyHTML = best
	}

	if len(res.bodyCandidates) > 0 {
		// Pick the longest valid plain text
		best := res.bodyCandidates[0]
		for _, b := range res.bodyCandidates[1:] {
			if len(b) > len(best) {
				best = b
			}
		}
		res.BodyPlainText = best
	}

	if len(res.BodyPlainText) > 0 && len(res.BodyHTML) == 0 {
		// If we have plain text but no HTML, set HTML to the plain text
		res.BodyHTML = res.BodyPlainText
	}
	if len(res.BodyHTML) > 0 && len(res.BodyPlainText) == 0 {
		// If we have HTML but no plain text, set plain text to the HTML
		res.BodyPlainText = res.BodyHTML
	}
	if len(res.BodyPlainText) == 0 && len(res.BodyHTML) == 0 {
		// If both are empty, set them to a default value or leave them empty
		res.BodyPlainText = "No content available"
		res.BodyHTML = "No content available"
	}
}
