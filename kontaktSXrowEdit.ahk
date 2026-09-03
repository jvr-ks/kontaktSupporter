; kontaktSXrowEdit.ahk
; Part of kontaktSupporter.ahk


;---------------------------- attachmentMenuBuild ----------------------------
attachmentMenuBuild(theMenu){
  global
  
  theMenu.Add("Ordner (Anhang) öffnen", openAttachmentFolder)
  
  theMenu.Add()
  theMenu.Add("Datei auswählen", selectAttachmentFile)
  theMenu.Add("Pfad auswählen (bei mehreren Dateien)", selectAttachmentPath)
  
  theMenu.Add()
  
  loop externalApps.length {
    theExternalName := externalApps[A_Index][1]
    if (theExternalName == ""){
      theMenu.Add()
    } else {
      theMenu.Add("Ordner/Datei öffnen mit " theExternalName, openExternal.Bind(externalApps[A_Index][1], externalApps[A_Index][2]))
    }
  }
  
  theMenu.Add("Externe Apps Liste `"" externalAppsFile "`" bearbeiten", editExternalApps)
  
  return theMenu
}

;--------------------------- openAttachmentFolder ---------------------------
openAttachmentFolder(*){
  global
  local toRun
  
  ; attachment can be a folder, or a file
  if (extractPath(attachment) == baseDirectory)
    toRun := attachmentsPathAbsolute
  else
    toRun := extractPath(attachment)
  
  openFolder(toRun)
}
;---------------------------------- rowEdit ----------------------------------
rowEdit(*){
 global
  ; local msg, width, height, opt, checked9, theFile
  local msg, width, height, opt, theFile
  
  AttachmentMenu := Menu()
  AttachmentMenu := attachmentMenuBuild(AttachmentMenu)
  
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    if (!rowEditIsVisible){
      guiRowEdit := Gui("+OwnDialogs ToolWindow MaximizeBox +Resize +parent" guiMain.Hwnd, "Bearbeiten - kontaktSupporter Tabellen-Eintrag bearbeiten")
      guiRowEdit.Opt("+Owner" guiMain.Hwnd)
          
      guiRowEdit.SetFont("S" . fontsize, fontName)
      
      ContentMenu := Menu()
      ContentMenu.Add("Datei öffnen (bearbeiten)", editBodyTextFile)
      ContentMenu.Add("Neue Datei aus Vorlage erstellen", createBodyTextFileFromProto)
      ContentMenu.Add("Vorhandene Datei auswählen", selectBodyTextFile)
      ContentMenu.Add("Neue Datei erstellen", createBodyTextFile)
      ContentMenu.Add("Ordner (Inhalt) öffnen", openBodyFileFolder)
      ContentMenu.Add("Ordner (Vorlagen) öffnen", openFolder.Bind(protoFilePathInUse))
      ContentMenu.Add("Ordner (Text Bausteine) öffnen", openFolder.Bind(textModulsFilePathInUse))
      ContentMenu.Add("♒ Vorschau", rowPreview)
      
      rowEditMenu := MenuBar()
      rowEditMenu.Add("☰" "Inhalt", ContentMenu)
      rowEditMenu.Add("🔗" "Anhang", AttachmentMenu)
      rowEditMenu.Add("📖→" "Bearbeiten schließen", guiRowEditClose)
      rowEditMenu.Add("🗙" "kontaktSupporter beenden", exit)

      
      guiRowEdit.MenuBar := rowEditMenu
    
      guiRowEdit.Add("Edit", "x" clientTopX " y" clientTopY " r0 h0 w100", "dummy1") ; Focus dummy and size dummy
      guiRowEdit.Add("Edit", "section x+m yp+0 w400 h0", "dummy2") ; section dummy
      
      guiRowEdit.Add("Text", "x" clientTopX, "Name: ")
      guiRowEdit_Edit1 := guiRowEdit.Add("Edit", "r1 w400 vCol1 xs yp+0" , nameRaw)
      guiRowEdit_Edit1.OnEvent("LoseFocus", rowEditSave , 1)
      
      guiRowEdit.Add("Text", "x" clientTopX, "Adresse: ")
      guiRowEdit_Edit2 := guiRowEdit.Add("Edit", "r1 w400 vCol2 xs yp+0" , adrTextRaw)
      guiRowEdit_Edit2.OnEvent("LoseFocus", rowEditSave , 1)
      
      guiRowEdit.Add("Text", "x" clientTopX, "Cc: ")
      guiRowEdit_Edit3 := guiRowEdit.Add("Edit", "r1 w400 vCol3 xs yp+0" , ccTextRaw)
      guiRowEdit_Edit3.OnEvent("LoseFocus", rowEditSave , 1)
      
      guiRowEdit.Add("Text", "x" clientTopX, "Bcc: ")
      guiRowEdit_Edit4 := guiRowEdit.Add("Edit", "r1 w400 vCol4 xs yp+0" , bccTextRaw)
      guiRowEdit_Edit4.OnEvent("LoseFocus", rowEditSave , 1)
      
      guiRowEdit.Add("Text", "x" clientTopX, "Subject: ")
      guiRowEdit_Edit5 := guiRowEdit.Add("Edit", "r1 w400 r2 vCol5 xs yp+0" , subjectTextRaw)
      guiRowEdit_Edit5.OnEvent("LoseFocus", rowEditSave , 1)
      
      guiRowEdit.Add("Text", "x" clientTopX)
      
      guiRowEdit.Add("Text", " x" clientTopX, "☰ Inhalt:`n(Dateiname)")
      guiRowEdit_Edit6 := guiRowEdit.Add("Edit", "readonly r2 w400 vCol6 xs yp" , bodyFileNameRaw)
      
      guiRowEdit.Add("Text", "x" clientTopX, "🔗 Anhang: ")
      guiRowEdit_Edit7 := guiRowEdit.Add("Edit", "r2 w400 vCol7 xs yp+0" , attachmentRaw)
      guiRowEdit_Edit7.OnEvent("LoseFocus", rowEditSave , 1)
      
      guiRowEdit_Edit7a := guiRowEdit.Add("Edit", "readonly r2 w400 xs yp+0", attachment)
      
      guiRowEdit.Add("Text", "x" clientTopX)
      
      guiRowEdit.Add("Text", "x" clientTopX, "Memo: ")
      guiRowEdit_Edit8 := guiRowEdit.Add("Edit", "r4 w400 vCol9 xs yp+0" , decodeNotAllowed(memoFieldRaw))
      guiRowEdit_Edit8.OnEvent("LoseFocus", rowEditSave , 1)
      
      guiRowEdit.Add("Text", "x" clientTopX, "Erledigt (Ok): ")
      guiRowEdit_checkBox1_checked := (okFieldRaw) ? "Checked" : ""
      guiRowEdit_checkBox1 := guiRowEdit.Add("Checkbox", "vCol8 xs yp+0 " guiRowEdit_checkBox1_checked)
      guiRowEdit_checkBox1.OnEvent("Click", rowEditSave , 1)
      
      guiRowEdit_buttonDismiss := guiRowEdit.Add("Button", "w400 xs" clientTopX, "Schließen")
      guiRowEdit_buttonDismiss.OnEvent("Click", guiRowEditClose)
      
      guiRowEdit.show("Autosize Center")
      
      guiRowEdit.OnEvent("Close", guiRowEdit_Close)
    } else {
      guiRowEdit.destroy
    }
    
    rowEditIsVisible := !rowEditIsVisible
   }
}
;-------------------------------- rowEditSave --------------------------------
rowEditSave(*){
  global 
  local Name, value, s , nPlus1
  
  Saved  := guiRowEdit.Submit(0)
  
  guiRowEdit_Edit7a.Text := modifyAttachmentsPath(guiRowEdit_Edit7.Text)
  guiRowEdit_Edit7.Text := pathToRelativ(guiRowEdit_Edit7.Text, extractPath(attachment))
  
  ; first column is the Index now
  For name , value in Saved.OwnProps() {
    nPlus1 := "Col" (A_Index + 1)
    LV1.Modify(currentLineNumber, nPlus1, value)
    mTb[currentLineNumber][A_Index] := encodeNotAllowed(value)
  }
  
  LV1Update()
  LV1Save()
}
;----------------------------- guiRowEditClose -----------------------------
guiRowEditClose(*){
  global 

  rowEditIsVisible := 0
  guiRowEdit.destroy
}
;----------------------------- guiRowEdit_Close -----------------------------
guiRowEdit_Close(*){
  global 

  rowEditIsVisible := 0
}
;----------------------------- editBodyTextFile -----------------------------
editBodyTextFile(*){
  global 
  local theFile, thePath, toRun, fileExtension

  toRun := ""
  theFile := autoAppendTxtExt(bodyFileName)
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    if (InStr(theFile, ":")){ ; filename includes directory?
      toRun := theFile
    } else {
      toRun := bodyFilePathInUse "\" theFile
    }
    if (FileExist(toRun)){
      fileExtension := extractExtension(toRun)
      if (fileExtension = ".bat" || fileExtension = ".cmd"){
        msgbox("Die Datei `"" toRun "`" ist eine ausführbare Datei, bitte mit einem externen Editor öffnen!", "Fehler aufgetreten!", "Icon!")
      } else {
        run toRun
      }
    } else {
      msgbox("Datei `"" toRun "`" nicht gefunden (editBodyTextFile)!", "Fehler aufgetreten!", "Icon!")
    }
  }
}
;---------------------------- selectBodyTextFile ----------------------------
selectBodyTextFile(*){
  global
  local selectedFile, effectivFP
  
  effectivFP := ""
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    selectedFile := FileSelect("1", bodyFilePathInUse, "Bitte eine Datei auswählen!")
    
    selectedFile := pathToRelativ(selectedFile, bodyFilePathInUse)
    
    if (selectedFile != ""){
      guiRowEdit_Edit6.Value := selectedFile
      rowEditSave()
    } else {
      showHintColored("Abgebrochen!")
    }
  }
}
;--------------------------- selectAttachmentFile ---------------------------
selectAttachmentFile(*){
  global
  local selectedFile, effectivFP
  
  effectivFP := ""
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    selectedFile := FileSelect("1", baseDirectory "\" attachmentsPath, "Bitte eine Datei auswählen! (bei mehreren Dateien: Pfad auswählen verwenden oder den Pfad manuell im Feld `"Anhang:`" eintragen!)")

    if (selectedFile != ""){
      effectivFP := selectedFile
      selectedFile := pathToRelativ(selectedFile, baseDirectory "\" attachmentsPath)
      
      guiRowEdit_Edit7.Value := selectedFile
      guiRowEdit_Edit7a.Value := modifyAttachmentsPath(effectivFP)
      
      rowEditSave()
    }
  }
}
;--------------------------- selectAttachmentPath ---------------------------
selectAttachmentPath(*){
  global
  local selectedPath, effectivFP
  
  effectivFP := ""
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    selectedPath := FileSelect("D", baseDirectory "\" attachmentsPath, "Bitte den Pfad auswählen! (bei mehreren Dateien): oder den Pfad manuell im Feld `"Anhang:`" eintragen!)")

    if (selectedPath != ""){
      selectedPath := pathToRelativ(selectedPath, baseDirectory "\" attachmentsPath)
      guiRowEdit_Edit7.Value := selectedPath
      guiRowEdit_Edit7a.Value := selectedPath
      
      rowEditSave()
    }
  }
}
;---------------------------- createBodyTextFile ----------------------------
createBodyTextFile(*){
  global
  local msg
  
  if (bodyFileName != "") {
    showHintColored("Der bestehende Dateinamen-Feldeintrag wird überschrieben,`n`naber die Datei `"" bodyFileName "`" wird nicht gelöscht!")
  }
  input := InputBox("Bitte geben Sie einen neue Dateinamen ein!", "Dateiname", "w500 h100")
  if input.Result != "Cancel" {
    if (input.Value != ""){
      newFile := bodyFilePathInUse "\" autoAppendTxtExt(input.Value)
      if (!FileExist(newFile)){
        FileAppend "", newFile, "`n"
        guiRowEdit_Edit6.Value := input.Value
        rowEditSave()
      } else {
        msgbox("Datei ist schon vorhanden!", "Bitte beachten!", "Icon!")
        guiRowEdit_Edit6.Value := input.Value
        rowEditSave()
      }
    } else {
      showHintColored("Dateiname ist leer!")
    }
  } else {
    showHintColored("Abgebrochen!")
  }
}
;------------------------ createBodyTextFileFromProto ------------------------
createBodyTextFileFromProto(*){
  global
  local theNewFile
  
  theNewFile := ""
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    if (bodyFileName != "") {
      showHintColored("Der bestehende Dateinamen-Feldeintrag wird überschrieben,`n`naber die Datei `"" bodyFileName "`" wird nicht gelöscht!")
    }
    if (FileExist(protoFilePathInUse)){
      theNewFile := selectPrototype()
      if (FileExist(bodyFilePathInUse "\" theNewFile)){
        guiRowEdit_Edit6.Value := theNewFile
        mTb[currentLineNumber][6] := theNewFile
        rowEditSave()
      } else {
        msgbox("Problem beim Verwenden einer Vorlage aus: `"" protoFilePathInUse "`" !", "Fehler aufgetreten!", "Icon!")
      }
    } else {
      msgbox("Problem beim Öffnen des Ordners: `"" protoFilePathInUse "`" !", "Fehler aufgetreten!", "Icon!")
      return
    }
  }
}
;------------------------------ selectPrototype ------------------------------
selectPrototype(){
  global 
  local e, newFileName, newFile,newPath
  
  newFile := ""
  newPath := ""
  newFileName := ""
  
  prototypeFile := FileSelect("1", protoFilePathInUse, "Vorlage wählen")
  
  if (prototypeFile != ""){
    input := InputBox("Bitte geben Sie einen neue Dateinamen ein!`n(Die Dateiendung wird automatisch ergänzt)", "Dateiname", "w500 h100")
    if input.Result != "Cancel" {
      if (input.Value != ""){
        newFileName := extractFileName(input.Value)
        prototypeFileExtension := extractExtension(autoAppendTxtExt(prototypeFile))
        newFile := newFileName prototypeFileExtension
        newPath := bodyFilePathInUse "\" newFile
        if (!FileExist(newPath)){
          try {
            FileCopy(prototypeFile, newPath, 0)
          } catch as e {
            msgbox("Unerwarteter Fehler beim Erstellen der Kopie!`n`nwhat: " e.what "`nfile: " e.file 
        . "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra,, 16) 
        
            newFile := ""
          }
        } else {
          msgbox("Datei `"" newPath "`" ist schon vorhanden!", "Fehler aufgetreten!", "Icon!")
        }
      }
    }
  }
  return newFile
}
;---------------------------- openBodyFileFolder ----------------------------
openBodyFileFolder(*){
  global
  local toRun
  
  toRun := bodyFilePathInUse
  
  if (InStr(bodyFileName, ":")){ ; absolute path?
    toRun := extractPath(bodyFileName)
  } else {
    if (InStr(bodyFileName, "\")){
      found := RegExMatch(bodyFileName,"(.*\\).*", &match) ; extract sub-directory
      if (found){
        toRun := bodyFilePathInUse "\" match.1
      } else {
        toRun := bodyFilePathInUse
        msgbox("Fehler beim bestimmen des Pfades für: `"" bodyFileName "`" !", "Fehler aufgetreten!", "Icon!")
      }
    } else {
      toRun := bodyFilePathInUse
    }
  }
  
  run toRun
}
;-------------------------------- openFolder --------------------------------
openFolder(folder, *){
  global
  
  run A_ComSpec " /c start " folder
}

;----------------------------------------------------------------------------
