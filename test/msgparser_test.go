package main

import (
	"testing"

	msgparser "github.com/yuphing-ong/outlook-msg-parser"
)

func TestParseMsgFile(t *testing.T) {
	// Parse the test.msg file
	msg, err := msgparser.ParseMsgFile("test.msg")
	if err != nil {
		t.Fatalf("Failed to parse file: %v", err)
	}

	// Perform assertions to verify the parsed data
	if msg.Subject == "" {
		t.Error("Subject is empty")
	}
	if msg.FromEmail == "" {
		t.Error("FromEmail is empty")
	}
	if msg.ToDisplay == "" {
		t.Error("To is empty")
	}
	if msg.CCDisplay == "" {
		t.Error("CC is empty")
	}
	if msg.BodyPlainText == "" && msg.BodyHTML == "" {
		t.Error("Both BodyPlainText and BodyHTML are empty")
	}
	if msg.Date.IsZero() {
		t.Error("Date is not set")
	}
	if msg.ClientSubmitTime.IsZero() {
		t.Error("ClientSubmitTime is not set")
	}
	if msg.CreationDate.IsZero() {
		t.Error("CreationDate is not set")
	}
	if msg.LastModificationDate.IsZero() {
		t.Error("LastModificationDate is not set")
	}

	// Print the parsed message for manual verification
	t.Logf("Parsed Message: %+v", msg)
}
