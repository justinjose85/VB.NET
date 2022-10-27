Sub Main()
Dim stopWatch As New Stopwatch()
stopWatch.Start()

dateiname = ThisDoc.FileName(True)
Dim oProgressBar As Inventor.ProgressBar
oMessage = "Bearbeitung läuft auf " & dateiname
iStepCount = 6
oProgressBar = ThisApplication.CreateProgressBar(False, iStepCount, oMessage)
k = 1 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

If ThisApplication.ActiveDocument.DocumentType = kDrawingDocumentObject Then '1
Dim oDrawDoc As DrawingDocument = ThisApplication.ActiveDocument
Docloc = oDrawDoc.FullFilename
oRevision = oDrawDoc.PropertySets.Item("Inventor User Defined Properties").Item("AIMD_REVISION").Value
		vorlage = "*pfad*"
Raender =	"Standard"
		Schriftfelder = "*feldername*"

Dim Sheet1 As Sheet = oDrawDoc.Sheets.Item(1)
Dim oTitleBlock As TitleBlock = Sheet1.TitleBlock
Dim oTitBlkName As String =  oTitleBlock.Name
If Left(oTitBlkName,10) ="Ersatzteil" Then 
oProgressBar.Close
Exit Sub
End If

On Error Resume Next
Dim oSheettemp  As Sheet
oSheettemp = oDrawDoc.Sheets.Item(1)
Dim oTB1  As TitleBlock
oTB1 = oSheettemp.TitleBlock
Dim titleDef As TitleBlockDefinition
titleDef = oTB1.Definition

For Each defText In titleDef.Sketch.TextBoxes
Dim oValue As String
oValue = oTB1.GetResultText(defText)

    If defText.Text = "<Änderungsberichtnr. für A eintragen>" And oValue <> String.Empty Then
		    oPrompt = defText
	Else If defText.Text = "<Änderungsberichtnr. für B eintragen>" And oValue <> String.Empty Then
		    oPrompt = defText
	Else If defText.Text = "<Änderungsberichtnr. für C eintragen>" And oValue <> String.Empty Then
			oPrompt = defText
	Else If defText.Text = "<Änderungsberichtnr. für D eintragen>" And oValue <> String.Empty Then
	        oPrompt = defText
	Else If defText.Text = "<Änderungsberichtnr. für E eintragen>" And oValue <> String.Empty Then
	        oPrompt = defText
	Else If defText.Text = "<Änderungsberichtnr. für F eintragen>" And oValue <> String.Empty Then
	        oPrompt = defText
	Else If defText.Text = "<Änderungsberichtnr. für G eintragen>" And oValue <> String.Empty Then
	        oPrompt = defText
	Else If defText.Text = "<Änderungsberichtnr. für H eintragen>" And oValue <> String.Empty Then
	        oPrompt = defText
		Exit For
    End If
Next    

Dim Mat As String
Mat=oTB1.GetResultText(oPrompt)
If Mat ="" Then '1
Else
'If (ThisDrawing.ModelDocument Is Nothing) Then Return
Dim modelName As String = IO.Path.GetFileName(ThisDoc.ModelDocument.FullFileName)
oModelaNr = iProperties.Value(modelName, "Custom", "AVV_Aenderungsnummer")

If oModelaNr <> "" Then '2
aNrOld = Mid(Split(Mat,"-")(0),2) & Mid(Split(Mat ,"-")(1),1)
aNrNew = Mid(Split(oModelaNr,"-")(0),2) & Mid(Split(oModelaNr ,"-")(1),1)
If aNrNew > aNrOld Then Mat = oModelaNr 'No End
End If '2

iProperties.Value(modelName, "Custom", "AVV_Aenderungsnummer") = Mat
End If '1

Dim oNewDoc As inventor.DrawingDocument = ThisApplication.Documents.Add(kDrawingDocumentObject, vorlage, True)
Dim oSheets As Sheets = oDrawDoc.Sheets
Dim oSheet As Sheet

k = 2 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

oDrawDoc.Save

Dim oSheetCount As Integer
oSheetCount = oDrawDoc.Sheets.Count 

For Each oSheet In oSheets
oSheet.Activate()
oSheet.Border.Delete
oSheet.TitleBlock.Delete
oSheet.CopyTo(oNewDoc)
Next
k = 3 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

oDrawDoc.Close(True)
oNewDoc.Sheets.Item(1).Delete 

k = 4 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

Dim oSheetNew As Sheet
Dim oSheetNewCount As Integer
oSheetNewCount = oNewDoc.Sheets.Count 

If oSheetNewCount <> oSheetCount Then '3
oNewDoc.Close(True)
oProgressBar.Close
ThisApplication.Documents.Open(Docloc, True) 
MessageBox.Show("Nicht alle Blätter in "& dateiname &" kopiert werden." & vbLf & vbLf & "Hinweis : Getrennte Zeichnungsansicht(en).", "Fehler")
Exit Sub
Else '3

For Each oSheetNew In oNewDoc.Sheets 
	oSheetNew.Activate		
	oSheetNew.AddBorder(Raender)
	oSheetNew.AddTitleBlock(Schriftfelder)	
	
	Dim oPartsList As PartsList
If oSheetNew.PartsLists.Count > 0 Then '2
	oPartsList = oNewDoc.ActiveSheet.PartsLists.Item(1)
    oPartsList.Delete               
    Dim oDrawingView As DrawingView
    oDrawingView = oNewDoc.ActiveSheet.DrawingViews(1)
    Dim oBorder As Border
    oBorder = oNewDoc.ActiveSheet.Border
    
    Dim oPlacementPoint As Point2d
	xrev = oBorder.RangeBox.MaxPoint.X
    yrev = oBorder.RangeBox.MinPoint.Y
    oPlacementPoint = ThisApplication.TransientGeometry.CreatePoint2d(xrev, yrev)
	Dim oPartsList1 As PartsList
    oPartsList1 = oNewDoc.ActiveSheet.PartsLists.Add(oDrawingView, oPlacementPoint)
	oPartsList1 = oNewDoc.ActiveSheet.PartsLists.Item(1)
	PartHight=oPartsList1.RangeBox.MaxPoint.Y-oPartsList1.RangeBox.MinPoint.Y
	TitleY=oNewDoc.ActiveSheet.TitleBlock.RangeBox.MaxPoint.Y
	PointX=oBorder.Rangebox.Maxpoint.x
	PointY=TitleY+PartHight
	Dim oPlacementPoint1 As Point2d
	oPlacementPoint1 = ThisApplication.TransientGeometry.CreatePoint2d(PointX, PointY)
	oPartslist1.position = oPlacementPoint1
'oPartsList1.Sort("Artikelnummer", 1)
oPartsList1.Sort("Pos.", 1)
oPartsList1.Renumber
oPartsList1.Style.UpdateFromGlobal 
oPartsList1.Style = oNewDoc.StylesManager.PartsListStyles.Item("Bauteilliste (DIN)")
oPartsList1.SaveItemOverridesToBOM
End If'2
Next

k = 5 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

oNewDoc.Sheets.Item(1).Activate  

Dim oControlDef as ControlDefinition
oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd")
oControlDef.Execute2(True)

k = 6 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

Dim oTopNode As BrowserNode 
oTopNode = oNewDoc.BrowserPanes.ActivePane.TopNode 
Dim oNode As BrowserNode 
For Each oNode In oTopNode.BrowserNodes 
oNode.Expanded = False 
Next
oNewDoc.PropertySets.Item("Inventor User Defined Properties").Item("*REVISION*").Value = oRevision
oNewDoc.PropertySets.Item("Inventor Summary Information").Item("Revision Number").Value = oRevision

InventorVb.DocumentUpdate()
ThisApplication.ActiveEditDocument.SaveAs(Docloc, False)
End If '3

Else'1
End If '1

oProgressBar.UpdateProgress
Dim Time As DateTime = DateTime.Now
Dim Format2 As String = "HH:mm:ss"
stopWatch.Stop()
Dim ts As TimeSpan = stopWatch.Elapsed
Dim elapsedTime As String = String.Format("{0:00}:{1:00}:{2:00}:{3:000}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds)
'oProgressBar.Message = " Erledigt um " & Time.ToString(Format2) & "  (Laufzeit : " & elapsedTime & ")"
oProgressBar.Close
ThisApplication.StatusBarText= " Erledigt um " & Time.ToString(Format2) & "  (Laufzeit : " & elapsedTime & ")"
System.Threading.Thread.CurrentThread.Sleep(500)
End Sub
