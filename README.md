# outlook-msg-parser

A Go library to parse `.msg` files from Microsoft Outlook.

This library is designed to handle various MAPI properties and extract relevant information from `.msg` files. It supports different versions and implementations of the MAPI standard, ensuring compatibility with a wide range of `.msg` files.

## Features

- Parse plain text and HTML bodies of messages
- Extract email headers, subject, sender, recipients (To, CC, BCC)
- Handle attachments separately
- Support for multiple MAPI property tags
- Compatibility with different versions of Microsoft Outlook

## Mentions to the original developer 

Forked from willthrom/outlook-msg-parser, added some additional properties (see below), and clear up the example, since it didn't work as expected (e.g. ParseMsgFile expects a string pointing to the file, not a *os.File)

## Installation

To install the library, use `go get`:

```sh
go get github.com/yuphing-ong/outlook-msg-parser


##  Usage

```sh
Here is an example of how to use the library to parse a .msg file:

package main

import (
    "log"
    "os"

    "github.com/yuphing-ong/outlook-msg-parser"
)

func main() {

    file = "test.msg"
    // Parse the .msg file
    msg, err := OutlookMessageParser.ParseMsgFile(file)
    if err != nil {
        log.Fatalf("Failed to parse file: %v", err)
    }

    // Print the parsed message
    log.Printf("Parsed Message: %+v", msg)
}

```

## Properties added

| Hex| Descriptor | Type |
| --- | --- | --- |
| 0x3702 | PidTagAttachEncoding | PT_Binary|
