package main

import (
    "log"

    "github.com/go-ole/go-ole"
    "github.com/go-ole/go-ole/oleutil"
)

func Doc2XML(source, destination string) error {
    _ = ole.CoInitialize(0)
    defer ole.CoUninitialize()

    iunk, err := oleutil.CreateObject("Word.Application")
    if err != nil {
        log.Fatalf("error creating Word object: %s", err)
        return err
    }

    word := iunk.MustQueryInterface(ole.IID_IDispatch)
    defer word.Release()

    // opening then saving works due to the call to doc.Settings.SetUpdateFieldsOnOpen(true) above

    docs := oleutil.MustGetProperty(word, "Documents").ToIDispatch()
    wordDoc := oleutil.MustCallMethod(docs, "Open", source).ToIDispatch()

    // file format constant comes from https://msdn.microsoft.com/en-us/vba/word-vba/articles/wdsaveformat-enumeration-word
    const wdFormatPDF = 11
    oleutil.MustCallMethod(wordDoc, "SaveAs2", destination, wdFormatPDF)
    oleutil.MustCallMethod(wordDoc, "Close")
    oleutil.MustCallMethod(word, "Quit")
    return nil
}
