; kontaktSXinsert.ahk
; Part of kontaktSupporter.ahk

;----------------------------- kontaktSTBinsert -----------------------------
; TB => "TextBausteine"

kontaktSTBinsert(s, onceOnly := 0){
  global 
  local foundPos, fileName, extractedFileName
  
  theText := s
  
  foundPos := 1
  insertRegEx := "(?<!\\)°([.\S]+)"
  foundPos := RegExMatch(theText, insertRegEx, &extractedFileName, foundPos)
  
  while (foundPos){
    fileName := (extractedFileName[1]) ; 1. Gruppe
    fileName := textModulsFilePathInUse "\" fileName

    fileName := autoAppendTxtExt(fileName)
    
    if (FileExist(fileName)){
      toReplace := FileRead(fileName)
      theText := RegExReplace(theText, insertRegEx, toReplace, &OutputVarCount, 1, foundPos)
    } else {
      msgbox("Datei nicht gefunden (Quelle: `"" . extractedFileName[1] . "`"): `"" fileName "`" !", "Datei-Fehler aufgetreten, kontaktSupporter wird beendet!", "Icon!")
      exit()
    }
    if (onceOnly){
      foundPos := 0
    } else {
      ; next one
      foundPos := RegExMatch(theText, insertRegEx, &extractedFileName, foundPos)
    }
  }
  theText := StrReplace(theText, "\\°", "°")
  
  ; Macro Expansion
  foundPos := 1
  macroRegEx := "(?<!(?<!\\)\\)#(.{1,3})#\s"
  
  foundPos := RegExMatch(theText, macroRegEx, &extractedFileName, foundPos)
  while (foundPos){
    macroName := (extractedFileName[1]) ; 1. Gruppe
    toReplace := ""
    
    Switch macroName {
      Case "d":
        toReplace := formattime(,"dd.MM.yyyy")
        theText := RegExReplace(theText, macroRegEx, toReplace, &OutputVarCount, 1, foundPos)
        ; replace "\\" -> "\"
        theText := RegExReplace(theText, "\\\\" escapeRegExp(toReplace), "\" toReplace, &OutputVarCount, 1, 1)
        
      Case "dl":
        toReplace := formattime(,"LongDate")
        theText := RegExReplace(theText, macroRegEx, toReplace, &OutputVarCount, 1, foundPos)
        theText := RegExReplace(theText, "\\\\" escapeRegExp(toReplace), "\" toReplace, &OutputVarCount, 1, 1)
        
      Case "u":
        toReplace := formattime(,"HH:mm:ss")
        theText := RegExReplace(theText, macroRegEx, toReplace, &OutputVarCount, 1, foundPos)
        theText := RegExReplace(theText, "\\\\" escapeRegExp(toReplace), "\" toReplace, &OutputVarCount, 1, 1)
        
      Default:
        theText := RegExReplace(theText, macroRegEx, "", &OutputVarCount, 1, foundPos)
    }
    ; next one
    foundPos := RegExMatch(theText, macroRegEx, &extractedFileName, foundPos)
  }
  ; Keep deactivated macros, but without the backslash!
  theText := RegExReplace(theText,"(?<!\\)\\(#.{1,3}#\s)","$1")
  
  return theText
}
;------------------------------- escapeRegExp -------------------------------
escapeRegExp(s) {
  local index, value
  
  regexSpecialCharacters := ["\", ".", "+", "*", "?","[", "^", "]", "$", "(",")", "{", "}", "=", "!","<", ">", "|", ":", "-"]
  r := ""
  loop Parse s {
    v1 := A_LoopField
     for index, value in regexSpecialCharacters {
        if (v1 = value)
          v1 := "\" v1
    }
    r .= v1
  }

  return r
}
;------------------------------ unescapeRegExp ------------------------------
/*
unescapeRegExp(s) {
  r := RegExReplace(s,"\\(?!\\)","")

  return r
}
 */