# OutlookMessageParser-Go (Orinally from https://github.com/oucema001/OutlookMessageParser-Go )

A Go library to parse `.msg` files from Microsoft Outlook.

This library is designed to handle various MAPI properties and extract relevant information from `.msg` files. It supports different versions and implementations of the MAPI standard, ensuring compatibility with a wide range of `.msg` files.

## Features

- Parse plain text and HTML bodies of messages
- Extract email headers, subject, sender, recipients (To, CC, BCC)
- Handle attachments separately
- Support for multiple MAPI property tags
- Compatibility with different versions of Microsoft Outlook

## Installation

To install the library, use `go get`:

```sh
go get github.com/willthrom/outlook-msg-parser


##  Usage

```sh
Here is an example of how to use the library to parse a .msg file:

package main

import (
    "log"
    "os"

    "github.com/willthrom/outlook-msg-parser"
)

func main() {
    // Open the .msg file
    file, err := os.Open("path/to/your/file.msg")
    if err != nil {
        log.Fatalf("Failed to open file: %v", err)
    }
    defer file.Close()

    // Parse the .msg file
    msg, err := OutlookMessageParser.ParseMsgFile(file)
    if err != nil {
        log.Fatalf("Failed to parse file: %v", err)
    }

    // Print the parsed message
    log.Printf("Parsed Message: %+v", msg)
}

```
