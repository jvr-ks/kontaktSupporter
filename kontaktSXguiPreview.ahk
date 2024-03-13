; kontaktSXguiPreview.ahk
; Part of kontaktSupporter.ahk

;-------------------------------- previewGui --------------------------------
previewGui(){
  global

  paddingLeft := 0
  paddingRight := 10
  paddingTop := 120
  paddingBottom := 60
  sciRAWwidth := 800
  sciRAWheight := 600
  
  guiPreview := Gui("+OwnDialogs MaximizeBox +Resize + AlwaysOnTop", "Vorschau (keine Änderungen in der Vorschau möglich!)")
  
  guiPreviewMenu := MenuBar()
  guiPreviewMenu.Add("♒→" "Vorschau beenden", guiPreview_Close)
  guiPreviewMenu.Add("🗙" "kontaktSupporter beenden ", exit)
  
  guiPreview.MenuBar := guiPreviewMenu
  
  guiPreview.SetFont("S" . fontsizePreview, fontNamePreview)
  
  guiPreview.Add("Edit", "x" clientTopX " y" clientTopY " r0 h0 w600", "dummy1") ; Focus dummy and size dummy
  
  guiPreview.Add("Text", "w600 x" clientTopX, "eMail Parameter: ")
  previewParam := guiPreview.Add("Text", "w600 x" clientTopX)
  
  guiPreview.Add("Text", "w600 x" clientTopX, "eMail Parameter (expanded): ")
  previewParamExpanded := guiPreview.Add("Text", "w600 x" clientTopX)
  
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
  
  guiPreview.show("Hide x" guiPreviewPosX " y" guiPreviewPosY " w" guiPreviewClientWidth " h" guiPreviewClientHeight)
  
  ;WinSetAlwaysOnTop -1, "A"
  
  guiPreview.OnEvent("Size", guiPreview_Size , 1)
  guiPreview.OnEvent("Close", guiPreview_Close)
  
  guiPreviewHwnd := guiPreview.Hwnd 
  ; WinMinimize "A" ; trigger size-event
  ; WinRestore "ahk_id " guiPreviewHwnd
  
  previewIsVisible := 0

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
      guiPreviewPosX := max(coordsAppToScreen(posX), minPosLeft)
      guiPreviewPosY := max(coordsAppToScreen(posY), minPosTop)

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
  
  paramExpanded := format("{:03}", currentLineNumber) ": "
  paramExpanded .= (name != "") ? name " / " : ""
  paramExpanded .= (adrText != "") ? adrText " / " : ""
  paramExpanded .= (ccText != "") ? ccText " / " : ""
  paramExpanded .= (bccText != "") ? bccText " / " : "---"
  paramExpanded .= (subjectText != "") ? subjectText " / " : ""
  paramExpanded .= (attachment != "") ? attachment " / " : ""
  
  previewParamExpanded.Value := paramExpanded
}
;------------------------------ imagePreviewGui ------------------------------
imagePreviewGui(){
  global
  
  guiImagePreview := Gui("+toolwindow +AlwaysOnTop", attachment)
  guiImagePreview.Show("Hide center")
  guiImagePreview.OnEvent("Close", guiImagePreview_Close)
}
;------------------------------- imagePreview -------------------------------
imagePreview(){
  global
  local ext, deltaX, deltaY, pBitmap, originalWidth, originalHeight, ratio, newWidth, newHeight
  local G, hbm, hdc, obm
  
  if (attachment != ""){
    ext := extractExtension(extractFileName(attachment))
    if (ext = ".png" || ext = ".jpg" || ext = ".gif"){

      If (!pToken) {
        guiImagePreview.Add("ActiveX", Format("w{1} h{2}", A_ScreenWidth - 30, A_ScreenHeight - 30), "mshtml:<img src='" attachment "' />")
      } else {
        ; GDI
        pBitmap := Gdip_CreateBitmapFromFile(attachment)
        If (!pBitmap){
          msgbox("Datei " attachment " konnte nicht geladen werden!", " Fehler aufgetreten!", "Icon!")
          return
        }
        previewImage := guiImagePreview.Add("Picture")
                
        pBitmap := Gdip_CreateBitmapFromFile(attachment)
        newWidth := Gdip_GetImageWidth(pBitmap)
        newHeight := Gdip_GetImageHeight(pBitmap)

        sizeW := A_ScreenWidth
        sizeH := A_ScreenHeight
        if (newWidth > sizeW || newHeight > sizeH){
          resizeFactor := Min(sizeW/newWidth, sizeH/newHeight) 
                  
          PBitmapResized := Gdip_CreateBitmap(Round(newWidth * resizeFactor), Round(newHeight * resizeFactor))
          G := Gdip_GraphicsFromImage(pBitmapResized)
          Gdip_DrawImage(G, pBitmap, 0, 0, Round(newWidth * resizeFactor), Round(newHeight * resizeFactor), 0, 0, newWidth, newHeight)

          hCRBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmapResized)

          previewImage.Value := "HBITMAP:*" hCRBitmap

          Gdip_DeleteGraphics(G)
          Gdip_DisposeImage(PBitmapResized)
          Gdip_DisposeImage(pBitmap)
          DeleteObject(hCRBitmap)
        } else {
          hCRBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap)
          previewImage.Value := "HBITMAP:*" hCRBitmap
          Gdip_DisposeImage(pBitmap)
          DeleteObject(hCRBitmap)
        }
      }
    guiImagePreview.Show("autosize")
    }
  } else {
    guiImagePreview.Hide()
  }
}
;--------------------------- guiImagePreview_Close ---------------------------
guiImagePreview_Close(*){
  global
  
    guiImagePreview.Hide()
}

;----------------------------------------------------------------------------
































