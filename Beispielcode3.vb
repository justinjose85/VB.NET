Imports System.Windows.Forms
AddReference "Autodesk.Connectivity.WebServices.dll"
Imports ACW = Autodesk.Connectivity.WebServices
AddReference "Autodesk.DataManagement.Client.Framework.Vault.dll"
AddReference "Autodesk.DataManagement.Client.Framework.dll"
Imports VDF = Autodesk.DataManagement.Client.Framework
AddReference "Connectivity.Application.VaultBase.dll"
Imports VB = Connectivity.Application.VaultBase

Sub Main()

Dim stopWatch As New Stopwatch()
stopWatch.Start()

Dim curDoc	=	ThisApplication.ActiveEditDocument
If curDoc.DocumentType = kAssemblyDocumentObject Then
Dim oAssDoc As AssemblyDocument = ThisApplication.ActiveEditDocument 
Dim oAssDef As ComponentDefinition = oAssDoc.ComponentDefinition


Dim oOrdner As String = "C:\Datenexport\" 
Dim oFolderName As String = iProperties.Value("Custom", "AIMD_PARTNO")
If iProperties.Value("Custom", "AIMD_PARTNO") = "" Then oFolderName= iProperties.Value("Project", "Part Number")
oAssemOrd = oOrdner & oFolderName &" ("& iProperties.Value("Custom", "AIMD_REVISION") &")" & "\"

Dim oErrors As New ArrayList

Call TraverseAssembly(oAssDef.Occurrences, 1, oAssemOrd, oErrors) 

Shell("explorer.exe /Select," & oAssemOrd, vbNormalFocus)

stopWatch.Stop()
Dim ts As TimeSpan = stopWatch.Elapsed
Dim elapsedTime As String = String.Format("{0:00}:{1:00}:{2:00}:{3:000}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds)

Dim oValue As String
oValue = InputListBox("Laufzeit : "& elapsedTime, oErrors, "", "PDF nicht erstellt", "Datei(en) fehlgeschlagen.")

End If
End Sub
  
  
  Sub  TraverseAssembly(Occurrences As ComponentOccurrences, Level As Integer, oAssemOrd As String, oErrors As ArrayList)  
	On Error Resume Next
	
	ThisApplication.SilentOperation = True
	
	
	Dim oOcc As ComponentOccurrence
	Dim DxfVer As String = "FLAT PATTERN DXF?AcadVersion=2004&OuterProfileLayer=Outer"
	
	
	For Each oOcc In Occurrences 
			oPartDoc = oOcc.Definition.Document
			InvSumInfo = oPartDoc.PropertySets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
			DesTraProp = oPartDoc.PropertySets("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
			InvDocSum = oPartDoc.PropertySets("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
			oTitle = InvSumInfo.ItemByPropId(2).Value'Title
			oRevisn = InvDocSum.Item("AIMD_REVISION").Value 'Revision
			If oRevisn = "" Then oRevisn = "0"
   			oArtkNr = InvDocSum.Item("AIMD_PARTNO").Value'Artiklenummer
			
	If Not System.IO.Directory.Exists(oAssemOrd) Then System.IO.Directory.CreateDirectory(oAssemOrd)		
	
	
	oNameRt = Right(System.IO.Path.GetFileNameWithoutExtension(oOcc.Definition.Document.FullFileName),3)
	oName = Left(System.IO.Path.GetFileNameWithoutExtension(oOcc.Definition.Document.FullFileName),3)
	If oName = "KSS" Or oName = "BNM" Or oName ="ISO" Or oName = "DIN" Or oName = "Bli" Or oName = "EN8" Or oName = "EN1" Or oName = "HN3" Or oName = "DKO" Or oName = "FAS" Or oNameRt = "-W3" Or oNameRt = "-PP" Or oNameRt = "-BK" Then Continue For
		
	
	Dim PDFAddIn As TranslatorAddIn
    Dim oContext As TranslationContext
    Dim oOptions As NameValueMap
    Dim oDataMedium As DataMedium
    
    Call ConfigurePDFAddinSettings(PDFAddIn, oContext, oOptions, oDataMedium)
  
    Dim oDrawDoc As DrawingDocument
    

	
Dim docfullfilename As String = oPartDoc.FullFileName
Dim docfilename As String = RPointToBackSlash(oPartDoc.FullFileName)

'Alle Zeichnungen aus dem Vault abrufen
'Auf Vault-Connection zugreifen und ggf. rausgehen
Dim mVltCon As VDF.Vault.Currency.Connections.Connection = VB.ConnectionManager.Instance.Connection
If mVltCon Is Nothing Then Exit Sub
'Auf ACW-PropertyDefininition Status zugreifen
Dim filePropDefs As ACW.PropDef() = mVltCon.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
Dim ACWNamePropDef As ACW.PropDef
For Each def As ACW.PropDef In filePropDefs
	'MessageBox.Show(def.DispName)
    If def.DispName = "Name" Then
        ACWNamePropDef = def
		Exit For
    End If
Next def  
'Suchoptionen festlegen
Dim namesucheoptionen As New ACW.SrchCond() With { _
	.PropDefId = ACWNamePropDef.Id, _
	.PropTyp = ACW.PropertySearchType.SingleProperty, _
	.SrchOper = 1, _
	.SrchRule = ACW.SearchRuleType.Must, _
	.SrchTxt = docfilename & " idw" _
}

Dim bookmark As String = String.Empty
Dim status As ACW.SrchStatus = Nothing
Dim results As ACW.File() = mVltCon.WebServiceManager.DocumentService.FindFilesBySearchConditions(New ACW.SrchCond() {namesucheoptionen }, Nothing, Nothing, False, True, bookmark, status)

Dim settings As New VDF.Vault.Settings.AcquireFilesSettings(mVltCon)
If results Is Nothing Then

If Not Left(oArtkNr, 1) = "9" And Not oErrors.Contains(oArtkNr)Then 
oErrors.Add(oArtkNr)
End If
Else
	'Dim vaultFile As iLogicGetFromVault.GetFile

	For Each res In results
		If Right(res.Name, 4) = ".idw" Then
		Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration = New VDF.Vault.Currency.Entities.FileIteration(mVltCon, res)
		settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeRelatedDocumentation = True
		settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest
		settings.AddFileToAcquire(oFileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download)
		End If 
	Next
End If

'Fehler will immer auschecken deswegen kurz eine temporäre Datei öffnen
vorlage = "A:\VorlagenStile2021\Vorlagen\Norm.ipt"
'Try
Dim oNewDoc As Document
'Try
oNewDoc = ThisApplication.Documents.Add(kPartDocumentObject, vorlage, False)
'Catch
'End Try	
Dim aquiresults As VDF.Vault.Results.AcquireFilesResults 
'ThisApplication.SilentOperation = True
aquiresults = mVltCon.FileManager.AcquireFiles(settings)
'ThisApplication.SilentOperation = False
'Alle heruntergeladenen idw's in Liste
Dim idwList As New ArrayList

For Each aquiresult As VDF.Vault.Results.FileAcquisitionResult In aquiresults.FileResults
	Dim aquiresultpath As String = aquiresult.LocalPath.FullPath
If UCase(aquiresultpath).Contains(".IDW") Then
		idwList.Add(aquiresultpath)
	End If
Next

oNewDoc.Close(True)
'idw's öffnen
Dim oAdd As Int32 = 0

For Each idw As String In idwList

				oNewDoc  = ThisApplication.Documents.Open(idw, False)
				If oAdd = 0 Then
				opdfOrdNew = oAssemOrd & oArtkNr & " (" & oRevisn &")"
				Else
				opdfOrdNew = oAssemOrd & oArtkNr & " (" & oRevisn & ")-" & oAdd
				End If
				oDataMedium.FileName = opdfOrdNew & ".pdf"

				Call PDFAddIn.SaveCopyAs(oNewDoc, oContext, oOptions, oDataMedium)
            	oNewDoc.Close
				oAdd+=1



Next
		
		
		If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then 
			If Not System.IO.Directory.Exists(oAssemOrd) Then System.IO.Directory.CreateDirectory(oAssemOrd)
			oAssemOrdNew = oAssemOrd & oArtkNr & " (" & oRevisn & ")\"
			Call TraverseAssembly(oOcc.SubOccurrences, Level + 1, oAssemOrdNew, oErrors) 
		Else 
			If oOcc.Definition.Type = kSheetMetalComponentDefinitionObject Then 
			oAssemOrdNew = oAssemOrd & oArtkNr & " (" & oRevisn &")"
			Dim oDataIO As DataIO = oOcc.Definition.DataIO
			Dim oComDef As SheetMetalComponentDefinition = oOcc.Definition
			oDataIO.WriteDataToFile(DxfVer, oAssemOrdNew  & ".dxf")
			If oComDef.Bends.Count > 0 Then oDataIO.WriteDataToFile("ACIS SAT", oAssemOrdNew & ".sat")
									
			End If
			
			ThisApplication.StatusBarText = oOcc.Name
		End If

	ThisApplication.SilentOperation = False
		
	Next
	
  End Sub
  
  Sub ConfigurePDFAddinSettings(ByRef PDFAddIn As TranslatorAddIn, ByRef oContext As TranslationContext, ByRef oOptions As NameValueMap, ByRef oDataMedium As DataMedium)

    PDFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
    oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        
    oOptions = ThisApplication.TransientObjects.CreateNameValueMap
    oOptions.Value("All_Color_AS_Black") = 1
    oOptions.Value("Remove_Line_Weights") = 0
    oOptions.Value("Vector_Resolution") = 400
    oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
    oOptions.Value("Custom_Begin_Sheet") = 1
    oOptions.Value("Custom_End_Sheet") = 1

    oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
End Sub

Function RPointToBackSlash(ByVal strText As String) As String
    strText = Left(strText, InStrRev(strText, ".") - 1)
    RPointToBackSlash = Right(strText, Len(strText) - InStrRev(strText, "\"))
End Function

