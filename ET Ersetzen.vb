Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Win32
Imports Inventor.ViewOrientationTypeEnum
Imports Inventor.DrawingViewStyleEnum

Sub Main()
Dim stopWatch As New Stopwatch()
 stopWatch.Start()
 
Dim mats As String = New String(){"ET.idw","ET-Inhaltsverzeichnis.idw"}', "ET-Deckblatt.idw"}
temp = InputListBox("Wähle die Vorlage aus", mats, "ET.idw", Title := "ET-Vorlage austauschen", ListName := "ET-Vorlagen")
If temp = "" Then Exit Sub 

Dim dateiname As String = ThisDoc.FileName(True)
Dim oPartDoc As Object= ThisDoc.ModelDocument
Dim oProgressBar As Inventor.ProgressBar
oMessage = "Bearbeitung läuft auf " & dateiname'Split(dateiname, ".")(0) 
iStepCount = 6
oProgressBar = ThisApplication.CreateProgressBar(False, iStepCount, oMessage)
k = 1 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

If ThisApplication.ActiveDocument.DocumentType = kDrawingDocumentObject Then '1
Dim oDrawDoc As DrawingDocument = ThisApplication.ActiveDocument
Docloc = oDrawDoc.FullFilename
oRevision = oDrawDoc.PropertySets.Item("Inventor User Defined Properties").Item("AIMD_REVISION").Value
If temp = "ET.idw" Then
vorlage = "A:\VorlagenStile2021\Vorlagen\ET.idw"
Schriftfelder1 = "Ersatzteil A4 DE,GB"
Schriftfelder2 = "Ersatzteil Stückliste"
Stuekliste = "Stückliste(DIN)_Quer_ET"
PointX = 28.7000
PointY = 15.1000
Orientierung = 10242
RowMax = 28
PrtHigtMax = 14.1
ElseIf temp = "ET-Inhaltsverzeichnis.idw" Then
vorlage = "A:\VorlagenStile2021\Vorlagen\ET-Inhaltsverzeichnis.idw"
Schriftfelder1 = "Ersatzteil A4 DE,GB Inhaltsverzeichnis"
Schriftfelder2 = "Ersatzteil Inhaltsverzeichnis Stückliste"
Stuekliste = "Stückliste(DIN)_HOCH_ET"
PointX = 20.0000
PointY = 24.1000
Orientierung = 10243
RowMax = 45
PrtHigtMax = 23.1
'ElseIf temp = "ET-Deckblatt.idw" Then
'vorlage = "A:\VorlagenStile2021\Vorlagen\ET-Deckblatt.idw"
'Schriftfelder1 = "Ersatzteil Deckblatt"
'Schriftfelder2 = "Ersatzteil Deckblatt"
End If

oDrawDoc.Save
Dim oNewDoc As inventor.DrawingDocument = ThisApplication.Documents.Add(kDrawingDocumentObject, vorlage, True)
k = 2 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

Dim oSheets As Sheets = oDrawDoc.Sheets
Dim oSheet As Sheet

On Error Resume Next

Dim oSheetCount As Integer = oDrawDoc.Sheets.Count 

k = 3 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

For Each oSheet In oSheets
'oSheet.Activate()
oSheet.Border.Delete
oSheet.TitleBlock.Delete
oSheet.CopyTo(oNewDoc)
Next

oDrawDoc.Close(True)
oNewDoc.Sheets.Item(2).Delete
oNewDoc.Sheets.Item(1).Delete

k = 4 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

Dim oSheetNew As Sheet
Dim oSheetNewCount As Integer = oNewDoc.Sheets.Count 

If oSheetNewCount <> oSheetCount Then '3
oNewDoc.Close(True)
oProgressBar.Close
ThisApplication.Documents.Open(Docloc, True)
MessageBox.Show("Nicht alle Blätter in "& dateiname &" kopiert werden." & vbLf & vbLf & "Hinweis : Getrennte Zeichnungsansicht(en).", "Fehler")
Exit Sub
Else '3

Dim oSheetBase As Sheet = oNewDoc.Sheets.Item(2)
If oSheetBase.DrawingViews.Count < 1 Then
Dim ViewScale As Double= 0.05
Dim oViewLoc As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(-15,10)
oDrawingView = oSheetBase.DrawingViews.AddBaseView(oPartDoc,oViewLoc, ViewScale,kDefaultViewOrientation, kHiddenLineRemovedDrawingViewStyle,,,) 'oBaseViewOptions)
Else
oDrawingView = oSheetBase.DrawingViews(1)
End If

	Dim oPartsList4 As PartsList
	If oNewDoc.Sheets.Item(2).PartsLists.Count > 0 Then '2
	oPartsList4 = oNewDoc.Sheets.Item(2).PartsLists.Item(1)
    oPartsList4.Delete
	End If

	Dim oPlacementPoint As Point2d
	oPlacementPoint = ThisApplication.TransientGeometry.CreatePoint2d(PointX, PointY)
	Dim oPartsList1 As PartsList
	oPartsList1 = oNewDoc.Sheets.Item(2).PartsLists.Add(oDrawingView, oPlacementPoint)
	oPartsList1 = oNewDoc.Sheets.Item(2).PartsLists.Item(1)
	oPartsList1.Sort("Art.-Nr.", 1)
	oPartsList1.Renumber
	oPartsList1.Style.UpdateFromGlobal 
	oPartsList1.Style = oNewDoc.StylesManager.PartsListStyles.Item(Stuekliste)
	'oPartsList1.SaveItemOverridesToBOM

PartHight=oPartsList1.RangeBox.MaxPoint.Y-oPartsList1.RangeBox.MinPoint.Y	

If PartHight > PrtHigtMax Then '4
oSheetcountreq = Ceil(PartHight/PrtHigtMax)+1
'oSheetcountreq = Ceil(Round(oPartsList1.PartsListRows.Count/RowMax,1))+1

Dim j As Long = 0
Dim q As Long = 1
Dim r As Long = RowMax

Dim l As Long
For l = 3 To oSheetcountreq

If oSheetcountreq > oNewDoc.Sheets.Count Then
oNewDoc.Sheets.Add()
oNewDoc.ActiveSheet.Size = DrawingSheetSizeEnum.kA4DrawingSheetSize
'ActiveSheet.ChangeSize("A4", MoveBorderItems := True)
'oNewDoc.Sheets.Item(l).AddTitleBlock(Schriftfelder2)
'oNewDoc.Sheets.Item(l).Orientation  = Orientierung
'ActiveSheet.TitleBlock = Schriftfelder
'ThisApplication.ActiveDocument.ActiveSheet.Orientation=Orientierung
End If

If oNewDoc.Sheets.Item(l).DrawingViews.Count < 1 Then
oDrawingView.CopyTo(oNewDoc.Sheets.Item(l))
End If

Next

Dim oPartsList3 As PartsList
Dim oPartsList2 As PartsList

Dim m As Long
For m = 2 To oSheetcountreq



	If oNewDoc.Sheets.Item(m).PartsLists.Count > 0 Then '2
	oPartsList3 = oNewDoc.Sheets.Item(m).PartsLists.Item(1)
    oPartsList3.Delete
	End If
	
	oPartsList2 = oNewDoc.Sheets.Item(m).PartsLists.Add(oDrawingView, oPlacementPoint)
	oPartsList2 = oNewDoc.Sheets.Item(m).PartsLists.Item(1)
	oPartsList2.Sort("Art.-Nr.", 1)
	oPartsList2.Renumber
	oPartsList2.Style.UpdateFromGlobal 
	oPartsList2.Style = oNewDoc.StylesManager.PartsListStyles.Item(Stuekliste)
	'oPartsList(m).SaveItemOverridesToBOM

Dim i As Long
For i = 0 To j
	oPartsList2.PartsListRows.Item(i).Visible = False
Next
j +=RowMax

Dim p As Long
For p = RowMax+q To oPartsList2.PartsListRows.Count
	oPartsList2.PartsListRows.Item(p).Visible = False
Next
q +=RowMax

'ThisApplication.ActiveView.Fit
oNewDoc.Sheets.Item(m).Activate
PartHight=oPartsList2.RangeBox.MaxPoint.Y-oPartsList2.RangeBox.MinPoint.Y

If PartHight > PrtHigtMax Then
While PartHight > PrtHigtMax
oPartsList2.PartsListRows.Item(r).Visible = False
PartHight=oPartsList2.RangeBox.MaxPoint.Y-oPartsList2.RangeBox.MinPoint.Y
r -= 1
j -= 1
q -= 1
End While

End If
r+=RowMax

ThisApplication.CommandManager.ControlDefinitions.Item("AppZoomallCmd").Execute	
Next

End If'4	

k = 5 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

Dim oControlDef as ControlDefinition
oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd")
oControlDef.Execute2(True)

oNewDoc.Sheets.Item(1).Activate
ThisApplication.CommandManager.ControlDefinitions.Item("AppZoomallCmd").Execute
oNewDoc.Sheets.Item(1).AddTitleBlock(Schriftfelder1)
Dim n As Long
For n = 2 To oNewDoc.Sheets.Count
oNewDoc.Sheets.Item(n).Activate
oNewDoc.Sheets.Item(n).AddTitleBlock(Schriftfelder2)
oNewDoc.Sheets.Item(n).Orientation  = Orientierung
ThisApplication.CommandManager.ControlDefinitions.Item("AppZoomallCmd").Execute	
Next

'oSheet.Activate

k = 6 	
oProgressBar.Message = ("Schritt " & k & " von " & iStepCount)
oProgressBar.UpdateProgress

Dim oTopNode As BrowserNode 
oTopNode = oNewDoc.BrowserPanes.ActivePane.TopNode 
Dim oNode As BrowserNode 
For Each oNode In oTopNode.BrowserNodes 
oNode.Expanded = False 
Next

oNewDoc.PropertySets.Item("Inventor User Defined Properties").Item("AIMD_REVISION").Value = oRevision
oNewDoc.PropertySets.Item("Inventor Summary Information").Item("Revision Number").Value = oRevision
InventorVb.DocumentUpdate()
ThisApplication.ActiveEditDocument.SaveAs(Docloc, False)
End If '3

Else'1
Exit Sub
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
End sub