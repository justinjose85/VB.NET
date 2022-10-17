Sub Main()
Dim stopWatch As New Stopwatch()
stopWatch.Start()

    Dim oInvApp As Inventor.Application = ThisApplication
    Dim oDoc As AssemblyDocument = ThisApplication.ActiveEditDocument
    Dim oAsmDef As AssemblyComponentDefinition = oDoc.ComponentDefinition 
	    oOccs = oAsmDef.Occurrences
    Dim oOcc As ComponentOccurrence 
    Dim oName As String
	Dim oFileType As Boolean = True
	
	On Error Resume Next
	oCurPos = oAsmDef.RepresentationsManager.ActivePositionalRepresentation.Name
	oAsmDef.RepresentationsManager.PositionalRepresentations.Item("Hauptansicht").Activate

	Call TraverseAssembly(oAsmDef.Occurrences, 1) 
	ThisApplication.CommandManager.ControlDefinitions.Item("AssemblyBonusTools_AlphaSortComponentsCmd").Execute
	
	For Each oOcc In oOccs
	'oOcc.Definition.Document.DisplayName=""

	'If oOcc.DefinitionDocumentType = kPartDocumentObject Then 
	oNameRt = Right(System.IO.Path.GetFileNameWithoutExtension(oOcc.Definition.Document.FullFileName),3)
	oName = Left(System.IO.Path.GetFileNameWithoutExtension(oOcc.Definition.Document.FullFileName),3)
	oArtNr = oOcc.Definition.Document.PropertySets("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}").Item("AIMD_PARTNO").Value
	
	If oName = "KSS" Or oName = "BNM" Or oName ="ISO" Or oName = "DIN" Or oName = "Bli" Or oName = "EN8" Or oName = "EN1" Or oName = "HN3" Or oName = "DKO" Or oName = "FAS" Then
	oFileType = False
	Else If oNameRt = "-W3" Or oNameRt = "-PP" Or oNameRt = "-BK" Then
	oFileType = False
	Else
	oFileType = True
	End If
	
	If oFileType = True Then
	ThisApplication.StatusBarText=System.IO.Path.GetFileNameWithoutExtension(oOcc.Definition.Document.FullFileName) & "   " & oOcc.Name 
	j = GoExcel.FindRow("A:\VorlagenStile2021\Design Data\iLogic\Inhaltscenter Liste\Inhaltscenter Liste.xlsx", "Tabelle1", "Artikelnummer", "=", oArtNr)
	oPfad = GoExcel.CurrentRowValue("Pfad")
	
	If Not j < 1 Then 'vbNullString Then 
	'Else
	oFamName = Split(oPfad,"*")(0)
	oRawNr = Split(oPfad,"*")(1)
	Dim oContentCenter As ContentCenter
	oContentCenter = oInvApp.ContentCenter
	Call TraverseNode (oContentCenter.TreeViewTopNode, oOcc, oFamName, oRawNr)

	End If
	End If 
   	'End If
	Next
	
	'Kontrolle
	Dim oValueList As New ArrayList
	Dim oFilenewType As Boolean = True
	For Each oOcc In oOccs
	'If oOcc.DefinitionDocumentType = kPartDocumentObject Then 
	oName = Left(System.IO.Path.GetFileNameWithoutExtension(oOcc.Definition.Document.FullFileName),3)
	oNameRt = Right(System.IO.Path.GetFileNameWithoutExtension(oOcc.Definition.Document.FullFileName),3)
	oArtNr = oOcc.Definition.Document.PropertySets("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}").Item("AIMD_PARTNO").Value
	If oName = "KSS" Or oName = "BNM" Or oName ="ISO" Or oName = "DIN" Or oName = "Bli" Or oName = "EN8" Or oName = "EN1" Or oName = "HN3" Or oName = "DKO" Or oName = "FAS" Then
	oFilenewType = False
	Else If oNameRt = "-W3" Or oNameRt = "-PP" Or oNameRt = "-BK" Then
	oFilenewType = False
	Else
	oFilenewType = True
	End If
	
	If oFilenewType = True Then
	k = GoExcel.FindRow("A:\VorlagenStile2021\Design Data\iLogic\Inhaltscenter Liste\Inhaltscenter Liste.xlsx", "Tabelle1", "Artikelnummer", "=", oArtNr)
	
	
	If Not k < 1 Then
	oValueList.Add(oOcc.Name & "  " & k)
	End If
	
	End If
	
	'End If
	Next

Call TraverseAssembly(oAsmDef.Occurrences, 1)

'oAsmDef.RepresentationsManager.ActivePositionalRepresentation.Name = oCurPos 
oAsmDef.RepresentationsManager.PositionalRepresentations.Item(oCurPos).Activate
stopWatch.Stop()
Dim ts As TimeSpan = stopWatch.Elapsed
Dim elapsedTime As String = String.Format("{0:00}:{1:00}:{2:00}:{3:000}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds)

Dim oValue As String
oValue = InputListBox("Laufzeit : " & elapsedTime, oValueList, "", "Normteile ersetzt.", "Datei(en) fehlgeschlagen.")'& vbCrLf & vbCrLf & "Bitte Kontrollieren")
	
ThisApplication.CommandManager.ControlDefinitions.Item("AssemblyBonusTools_AlphaSortComponentsCmd").Execute
End Sub
  
  Sub  TraverseNode(ByVal Node As ContentTreeViewNode, oOcc As ComponentOccurrence, oFamName As String, oRawNr As Integer) 
On Error Resume Next
	Dim oFamily As ContentFamily	
	For Each oNode As ContentTreeViewNode In Node.ChildNodes
            If oNode.Families.Count > 0 Then
                For Each oFamily In oNode.Families 
				If oFamily.DisplayName = oFamName Then
					Dim oError As MemberManagerErrorsEnum
    				Dim strContentPartFileName As String
    				Dim strErrorMessage As String
    				strContentPartFileName = oFamily.CreateMember(oRawNr, oError, strErrorMessage)
					oOcc.Replace(strContentPartFileName, True)
					'Exit Sub
					'Exit For
        		End If	
                Next
            End If

            If oNode.ChildNodes.Count > 0 Then
                Call TraverseNode (oNode, oOcc, oFamName, oRawNr)
            End If

        Next
	'Exit Sub
  End Sub
  
  Sub  TraverseAssembly(Occurrences As ComponentOccurrences, Level As Integer)  
	On Error Resume Next
	Dim oOcc As ComponentOccurrence
	Dim Assemblynum As Integer = 1
	Dim Partnum As Integer = 1
	For Each oOcc In Occurrences 
			oPartDoc = oOcc.Definition.Document
			InvSumInfo = oPartDoc.PropertySets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
			DesTraProp = oPartDoc.PropertySets("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
			InvDocSum = oPartDoc.PropertySets("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
			'oOcc.Definition.Document.DisplayName=""
		If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then 
			oTitle = InvSumInfo.ItemByPropId(2).Value'Title
			oRevisn = InvDocSum.Item("AIMD_REVISION").Value 'Revision
			If oRevisn = "" Then oRevisn = "0"
   			oArtkNr = InvDocSum.Item("AIMD_PARTNO").Value'Artiklenummer
			If oArtkNr <> "" Then
			oOcc.Name = oArtkNr & "  " & oTitle & " ("& oRevisn & ") :"  & Assemblynum
			Else
			oBautNr = DesTraProp.ItemByPropId(5).Value
			oOcc.Name = oBautNr & "  " & oTitle & " ("& oRevisn & ") :"  & Assemblynum
			End If
			Assemblynum += 1 
			'Call TraverseAssembly(oOcc.SubOccurrences, Level + 1) 
		Else
		    
			oTitle = InvSumInfo.ItemByPropId(2).Value'Title
			oRevisn = InvDocSum.Item("AIMD_REVISION").Value 'Revision
			If oRevisn = "" Then oRevisn = "0"
			oAbmesng = InvDocSum.Item("AVV_ABMESSUNG").Value'Abmessung
			oMaterl = DesTraProp.ItemByPropId(20).Value'Material
			oArtkNr = InvDocSum.Item("AIMD_PARTNO").Value'Artiklenummer
			If oArtkNr <> "" Then
			oOcc.Name =oArtkNr & "  " & oTitle & " ("& oRevisn & ") (" & oAbmesng & ") ("& oMaterl & ") :"  & Partnum
			Else
			oBautNr = DesTraProp.ItemByPropId(5).Value
		    oOcc.Name =oBautNr & "  " & oTitle & " ("& oRevisn & ") (" & oAbmesng & ") ("& oMaterl & ") :"  & Partnum
      		End If
			Partnum += 1	 
		End If 
	Next
  End Sub
