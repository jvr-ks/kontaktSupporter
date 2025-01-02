; kontaktSXguiPreview.ahk
; Part of kontaktSupporter.ahk

;-------------------------------- previewGui --------------------------------
previewGui(){
  global

  paddingLeft := 0
  paddingRight := 10
  paddingTop := 60
  paddingBottom := 60
  sciRAWwidth := 800
  sciRAWheight := 600
  
  guiPreview := Gui("+OwnDialogs MaximizeBox +Resize + AlwaysOnTop", "Vorschau (keine Änderungen in der Vorschau möglich!)")
  
  guiPreviewMenu := MenuBar()
  
  guiPreviewMenu.Add("🖼" "Anhang Vorschau", openImagePreview)
  guiPreviewMenu.Add("♒→" "Vorschau beenden", guiPreview_Close)
  guiPreviewMenu.Add("🗙" "kontaktSupporter beenden ", exit)
  
  guiPreview.MenuBar := guiPreviewMenu
  
  guiPreview.SetFont("S" . fontsizePreview, fontNamePreview)
  
  guiPreview.Add("Edit", "x" clientTopX " y" clientTopY " r0 h0 w600", "dummy1") ; Focus dummy and size dummy
  
  ;guiPreview.Add("Text", "w1024 x" clientTopX, "eMail Parameter: ")
  previewParam := guiPreview.Add("Text", "w1024 x" clientTopX)
  
  ;guiPreview.Add("Text", "w1024 x" clientTopX, "eMail Parameter (effektiv): ")
  previewParamExpanded := guiPreview.Add("Text", "w1024 x" clientTopX)
  
  SB := guiPreview.Add("StatusBar")
  SB.SetParts(300, 300)
  
  if (!disableScintilla){
    sciStartX := paddingLeft
    sciStarty := paddingTop

    sciWidth := sciRAWwidth - paddingLeft - paddingRight
    sciHeight := sciRAWheight - paddingTop - paddingBottom
    
    previewText := guiPreview.AddScintilla("x" sciStartX " y" sciStarty " w" sciWidth " h" sciHeight "  DefaultOpt")
    setLighTheme(previewText)
    ;setKSLexer(previewText)
    
    previewText.callback := sci_Change
    
    previewText.Doc.ptr := previewText.Doc.Create(500000+100)

    previewText.Tab.Use := false
    previewText.Tab.Width := 2
    previewText.Margin.Width := 50
    previewText.Margin.Type := 0
    
    previewText.AutoSizeNumberMargin := true
  } else {
    previewText := guiPreview.Add("Edit", "w800 h600 x" clientTopX, "")
  }
  
  guiPreview.show(Format("Hide x{1} y{2} w{3} h{4}", guiPreviewPosX, guiPreviewPosY, guiPreviewClientWidth, guiPreviewClientHeight))
  
  guiPreview.OnEvent("Size", guiPreview_Size , 1)
  guiPreview.OnEvent("Close", guiPreview_Close)
  
  guiPreviewHwnd := guiPreview.Hwnd 
  
  previewIsVisible := 0
}
;----------------------------- openImagePreview -----------------------------
openImagePreview(*){
  global
  
  if (imagePreviewIsVisible){
    guiImagePreview_Close()
  } else {
    imagePreviewIsVisible := 1
    imagePreview()
  }
}
;------------------------------- guiPreview_Size -------------------------------
guiPreview_Size(theGui, theMinMax, clientWidth, clientHeight,*) {
  global 
  local previewTextWidth, previewTextHeight
  
  if (theMinMax = -1)
      return
  
  guiPreviewClientWidth := clientWidth
  guiPreviewClientHeight := clientHeight
  
  IniWrite clientWidth, configFile, "gui", "guiPreviewClientWidth"
  IniWrite clientHeight, configFile, "gui", "guiPreviewClientHeight"
  
  if (disableScintilla){  
    previewTextWidth := (guiPreviewClientWidth - marginLeft - marginRight)
    previewTextHeight := (guiPreviewClientHeight - previewTextMarginHeight)
  
    previewText.Move(, , previewTextWidth, previewTextHeight)
  } else {
    sciRAWwidth := guiPreviewClientWidth
    sciRAWheight := guiPreviewClientHeight

    previewTextWidth := sciRAWwidth - paddingLeft - paddingRight
    previewTextHeight := sciRAWheight - paddingTop - paddingBottom

    previewText.Move(, , previewTextWidth, previewTextHeight)   
  }
}
;------------------------------ guiPreviewMove ------------------------------
guiPreviewMove(){
  global
  local posX, posY
  
  if (IsSet(guiPreview)){
    guiPreview.GetPos(&posX, &posY)
  
    if (posX != 0 && posY != 0){  
      guiPreviewPosX := limitToViewableXPos(posX, guiPreviewClientWidth)
      guiPreviewPosY := limitToViewableYPos(posY, guiPreviewClientHeight)
      
      IniWrite guiPreviewPosX, configFile, "gui", "guiPreviewPosX"
      IniWrite guiPreviewPosY, configFile, "gui", "guiPreviewPosY"
    }
  }
}
;----------------------------- guiPreview_Close -----------------------------
guiPreview_Close(*){
  global
  
  if (previewIsVisible){
    previewIsVisible := 0
    displayPreview := 0
    if (IsSet(guiPreview))
      guiPreview.Hide()
  }
}
;-------------------------------- sci_Change --------------------------------
sci_Change(*){
  global 
  local guiCtrlObj, CurrentCol, CurrentLine, oSaved

  guiCtrlObj := guiPreview.FocusedCtrl
  if (IsObject(guiCtrlObj)){
    CurrentCol := EditGetCurrentCol(guiCtrlObj)
    CurrentLine := EditGetCurrentLine(guiCtrlObj)
    ;oSaved := guiPreview.Submit(0)
    ;currentLineContent := guiPreview.Text
    ;tooltip currentLineContent
    SB.SetText("Line: " CurrentLine " Column: " CurrentCol , 2, 1)
  }
}
;-------------------------------- rowPreview --------------------------------
rowPreview(*){
  global

  displayPreview := !displayPreview
  
  if (displayPreview)
    sendSelected(1)
  else
    guiPreview_Close()
    
}
;------------------------------ rowPreviewShow ------------------------------
rowPreviewShow(){
  global
  
  if (!selectedIsValid(currentLineNumber)){
    showHintColored("Bitte zunächst einen Eintrag auswählen!")
  } else {
    if (!previewIsVisible){
      previewIsVisible := 1
      guiPreview.show()
    } else {
      msgbox "Fehler! Dieser Codeteil sollte nicht erreicht werden!"
    }
  }
}
;----------------------------- rowPreviewSetText -----------------------------
rowPreviewSetText(s){
  global
  
  if (!disableScintilla){
    previewText.Text := s
  } else {
    previewText.Value := s
  }
}
;--------------------------- generatePreviewParam ---------------------------
generatePreviewParam(){
  global
  
  param := format("{:03}", currentLineNumber) ": "
  param .= (nameRaw != "") ? nameRaw " / " : ""
  param .= (adrTextRaw != "") ? adrTextRaw " / " : ""
  param .= (ccTextRaw != "") ? ccTextRaw " / " : ""
  param .= (bccTextRaw != "") ? bccTextRaw " / " : ""
  param .= (subjectTextRaw != "") ? subjectTextRaw " / " : ""
  param .= (attachmentRaw != "") ? attachmentRaw " / " : ""
  
  previewParam.Value := param
  
  ; paramExpanded := format("{:03}", currentLineNumber) ": "
  ; paramExpanded .= (name != "") ? name " / " : ""
  ; paramExpanded .= (adrText != "") ? adrText " / " : ""
  ; paramExpanded .= (ccText != "") ? ccText " / " : ""
  ; paramExpanded .= (bccText != "") ? bccText " / " : "---"
  paramExpanded := (subjectText != "") ? StrReplace(subjectText, "%20", " ") " / " : ""
  paramExpanded .= (attachment != "") ? attachment " / " : ""
  
  previewParamExpanded.Value := paramExpanded
}
;------------------------------ imagePreviewGui ------------------------------
imagePreviewGui(){
  global
  
  guiImagePreview := Gui("+AlwaysOnTop +resize", "Anhang Vorschau")
  
  guiImagePreviewEdit1 := guiImagePreview.Add("Edit","readonly x2 r3 w" guiImagePreviewClientWidth - 2)

  previewImage := guiImagePreview.Add("Picture", "x2")
  
  guiImagePreview.show(Format("Hide x{1} y{2} w{3} h{4}", guiImagePreviewPosX, guiImagePreviewPosY, guiImagePreviewClientWidth, guiImagePreviewClientHeight))
  guiImagePreviewHWND := guiImagePreview.hwnd
  
  guiImagePreview.OnEvent("Size", guiImagePreview_Size , 1)
  guiImagePreview.OnEvent("Close", guiImagePreview_Close)
}
;------------------------------- imagePreview -------------------------------
imagePreview(){
  global
  local ext
  
  if (attachment != ""){
    ext := extractExtension(extractFileName(attachment))
    if (ext = ".png" || ext = ".jpg" || ext = ".gif"){
      if (!pToken) {
        msgbox("GDI funktioniert leider nicht", "Fehler aufgetreten!", "Icon!")
      } else {
        theBitmap := Gdip_CreateBitmapFromFile(attachment)
        
        if (!theBitmap){
          msgbox("Datei " attachment " konnte nicht geladen werden!", " Fehler aufgetreten!", "Icon!")
          return
        }
        guiImagePreview.Show()
        imageWidth := Gdip_GetImageWidth(theBitmap)
        imageHeight := Gdip_GetImageHeight(theBitmap)
        
        resizeGuiImagePreview()
      }
    } else {
      ; attachment is not an image
    
      showHintColored("Keine interne Vorschau für dieses Format vorhanden.`nÖffne externe App,`nbitte diese dann manuell schließen ...", -10000,,,,, "y2 xcenter")
      
      guiPreview_Close
      run A_ComSpec " /c start " attachment 
    }
  } else {
    msgbox("Fehler, Anhang ist leer oder wurde nicht gefunden!")
    guiImagePreview.Hide()
  }
}
;--------------------------- resizeGuiImagePreview ---------------------------
resizeGuiImagePreview(){
  global
  local resizeFactor, newWidth, newHeight, hCRBitmap
  local gPointer, hbm, hdc, obm
  
  if (!theBitmap){
    return
  }
  
  border := 10
  resizeFactor := dpiCorrect * min(guiImagePreviewClientWidth/imageWidth, guiImagePreviewClientHeight/imageHeight)
  newWidth :=  Round(imageWidth * resizeFactor)
  newHeight := Round(imageHeight * resizeFactor)
  
  newBitMap := Gdip_CreateBitmap(newWidth, newHeight)
  gPointer := Gdip_GraphicsFromImage(newBitMap) ; a pointer
  Gdip_SetInterpolationMode(gPointer, 7)
  Gdip_DrawImage(gPointer, theBitmap, 0, 0, newWidth, newHeight, 0, 0, imageWidth, imageHeight)

  previewImage.Value := ""
  hCRBitmap := Gdip_CreateHBITMAPFromBitmap(newBitMap)

  previewImage.Value := "HBITMAP:*" hCRBitmap

  Gdip_DeleteGraphics(gPointer)
  Gdip_DisposeImage(newBitMap)
  DeleteObject(hCRBitmap)
  
  msg := "Pfad: " modifyAttachmentsPath(attachment)
  msg .= "`nDatei: " extractFileName(attachment)
  msg .= "`nBildbreite: " imageWidth ", Bildhöhe: " imageHeight
    
  guiImagePreviewEdit1.Value := msg
}
;--------------------------- guiImagePreview_Size ---------------------------
guiImagePreview_Size(theGui, theMinMax, clientWidth, clientHeight,*) {
  global 
  
  if (theMinMax = -1)
      return
  
  guiImagePreviewClientWidth := clientWidth
  guiImagePreviewClientHeight := clientHeight
  
  IniWrite clientWidth, configFile, "gui", "guiImagePreviewClientWidth"
  IniWrite clientHeight, configFile, "gui", "guiImagePreviewClientHeight"
  
  previewImage.Move(,, guiImagePreviewClientWidth - 10, guiImagePreviewClientHeight - 10)
  
  resizeGuiImagePreview()
  AutoXYWH2("w", guiImagePreviewEdit1)
}
;---------------------------- guiImagePreviewMove ----------------------------
guiImagePreviewMove(){
  global
  local posX, posY
  
  if (IsSet(guiImagePreview)){
    guiImagePreview.GetPos(&posX, &posY)
  
    if (posX != 0 && posY != 0){  
      guiImagePreviewPosX := limitToViewableXPos(posX, guiImagePreviewClientWidth)
      guiImagePreviewPosY := limitToViewableYPos(posY, guiImagePreviewClientHeight)

      IniWrite guiImagePreviewPosX, configFile, "gui", "guiImagePreviewPosX"
      IniWrite guiImagePreviewPosY, configFile, "gui", "guiImagePreviewPosY"
    }
  }
}
;--------------------------- guiImagePreview_Close ---------------------------
guiImagePreview_Close(*){
  global
  if (IsSet(theBitmap))
    Gdip_DisposeImage(theBitmap)
    
  guiImagePreview.Hide()
  imagePreviewIsVisible := 0
}
;----------------------------------------------------------------------------
































