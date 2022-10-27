Private Sub Main()

		Dim oInvApp As Inventor.Application = ThisApplication
	    Dim oContentCenter As ContentCenter
        oContentCenter = oInvApp.ContentCenter
        Families = New List(Of Family)
        ReadContentCenter (oContentCenter.TreeViewTopNode)
        ListToExcel()

    End Sub

	Private Families As List(Of Family)

    Private Class FamilyMember
        Public Property Artikelnummer
      	Public Property Artikelnummer1
		Public Property Artikelnummer2
		Public Property Artikelnummer3
		Public Property Abmessung
		Public Property Dateiname
		Public Property ArtikelnummerDis
		Public Property oTableRow
		Public Property AIMD_PARTNO
		Public Property Artikel
    End Class

    Private Class Family
        Public Property Name
		
        Public Property Members As List(Of FamilyMember)

        Public Sub New()
            Members = New List(Of FamilyMember)
        End Sub
    End Class

    Private Sub ReadContentCenter(ByVal Node As ContentTreeViewNode)
        For Each oNode As ContentTreeViewNode In Node.ChildNodes
            If oNode.Families.Count > 0 Then
                For Each oFamily As ContentFamily In oNode.Families
				
					'FullNodeName = oNode.FullTreeViewPath
					Dim Family As New Family
                    Family.Name = oFamily.DisplayName
					
                    Dim oColumnNumber(8) As Integer
                    oColumnNumber = GetColumnNumbers(oFamily)
						
						Dim oTableRow1 As Integer
						oTableRow1 = 1

                    For Each oMember As ContentTableRow In oFamily.TableRows
                        						
						Dim FamilyMember As New FamilyMember
                        
						Try
                            FamilyMember.Artikelnummer = oMember.GetCellValue(oColumnNumber(0))
						Catch ex As Exception
                            FamilyMember.Artikelnummer = "NA" 				
                        End Try
                       				
						Try
                            FamilyMember.Artikelnummer1 = oMember.GetCellValue(oColumnNumber(1))
						Catch ex As Exception
                            FamilyMember.Artikelnummer1 = "NA" 				
                        End Try
						
						Try
                            FamilyMember.Artikelnummer2 = oMember.GetCellValue(oColumnNumber(2))
						Catch ex As Exception
                            FamilyMember.Artikelnummer2 = "NA" 				
                        End Try
						
						Try
                            FamilyMember.Artikelnummer3 = oMember.GetCellValue(oColumnNumber(3))
						Catch ex As Exception
                            FamilyMember.Artikelnummer3 = "NA" 				
                        End Try
						
						Try
                            FamilyMember.Abmessung = oMember.GetCellValue(oColumnNumber(4))
						Catch ex As Exception
                            FamilyMember.Abmessung = "NA" 				
                        End Try
						
						Try
                            FamilyMember.Dateiname = oMember.GetCellValue(oColumnNumber(5))
						Catch ex As Exception
                            FamilyMember.Dateiname = "NA" 				
                        End Try
						
						Try
                            FamilyMember.ArtikelnummerDis = oMember.GetCellValue(oColumnNumber(6))
						Catch ex As Exception
                            FamilyMember.ArtikelnummerDis = "NA" 				
                        End Try
						
						Try
						FamilyMember.AIMD_PARTNO = oMember.GetCellValue(oColumnNumber(7))
						Catch ex As Exception
                            FamilyMember.AIMD_PARTNO= "NA" 				
                        End Try
						
						Try
						FamilyMember.Artikel = oMember.GetCellValue(oColumnNumber(8))
						Catch ex As Exception
                            FamilyMember.Artikel= "NA" 				
                        End Try
						
						
						'01 Artikel
						
						FamilyMember.oTableRow = oTableRow1 
						oTableRow1 = oTableRow1+1
                        Family.Members.Add (FamilyMember)
                    Next

                    Families.Add (Family)
                Next
            End If

            If oNode.ChildNodes.Count > 0 Then
                ReadContentCenter (oNode)
            End If

        Next

    End Sub

    Private Function GetColumnNumbers(ByVal oFamily As ContentFamily) As Integer()

        Dim oColumnNumbers(8) As Integer
        Dim oFlag(8) As Boolean
        oFlag(0) = False
        oFlag(1) = False
        oFlag(2) = False
	oFlag(3) = False
	oFlag(4) = False
	oFlag(5) = False
	oFlag(6) = False
	oFlag(7) = False
	oFlag(8) = False

        For i As Integer = 1 To oFamily.TableColumns.Count
            If oFamily.TableColumns.Item(i).InternalName = "Artikelnummer" Then
                oColumnNumbers(0) = i
                oFlag(0) = True
           
			
			ElseIf oFamily.TableColumns.Item(i).InternalName = "Artikelnummer " Then
                oColumnNumbers(1) = i
                oFlag(1) = True

			ElseIf oFamily.TableColumns.Item(i).InternalName = "Artieklnummer" Then
                oColumnNumbers(2) = i
                oFlag(2) = True	
				
			ElseIf oFamily.TableColumns.Item(i).InternalName = "AIMD_PARTNO" Then
            	oColumnNumbers(3) = i
            	oFlag(3) = True	
			
			ElseIf oFamily.TableColumns.Item(i).DisplayHeading = "Abmessung" Then
            	oColumnNumbers(4) = i
            	oFlag(4) = True	
				
			ElseIf oFamily.TableColumns.Item(i).DisplayHeading = "Dateiname" Then
            	oColumnNumbers(5) = i
            	oFlag(5) = True	
							
			ElseIf oFamily.TableColumns.Item(i).DisplayHeading = "Artikelnummer" Then
            	oColumnNumbers(6) = i
            	oFlag(6) = True	
			
			ElseIf oFamily.TableColumns.Item(i).DisplayHeading = "AIMD_PARTNO [Custom]" Then
            	oColumnNumbers(7) = i
            	oFlag(7) = True	
				
			ElseIf oFamily.TableColumns.Item(i).DisplayHeading = "01 Artikel" Then
            	oColumnNumbers(8) = i
            	oFlag(8) = True	
				
				
		    End If
            If oFlag(0) And oFlag(1) And oFlag(2) And oFlag(3) And oFlag(4) And oFlag(5) And oFlag(6) And oFlag(7) And oFlag(8) Then
                Exit For
            End If
        Next

        GetColumnNumbers = oColumnNumbers
    End Function

    Private Sub ListToExcel()

        Dim oExcel As Object = CreateObject("Excel.Application")
		oExcel.Visible = True

        Dim oWorkbook As Object
        oWorkbook = oExcel.Workbooks.Add()

        Dim oSheet As Object
        oSheet = oWorkbook.Sheets.Item(1)

		oSheet.Range("A" & 1).Value = "Pfad"
		oSheet.Range("B" & 1).Value = "Artikelnummer"

        Dim oCount As Integer = 2 'starting row        
		For Each Family In Families
            For Each FamilyMember In Family.Members
                
				oSheet.Range("A" & oCount).Value = Family.Name & "*" & FamilyMember.oTableRow
                
				If FamilyMember.Artikelnummer = "NA" Then
				FamilyMember.Artikelnummer = FamilyMember.Artikelnummer1
				End If
				
				If FamilyMember.Artikelnummer = "NA" Then
				FamilyMember.Artikelnummer = FamilyMember.Artikelnummer2
				End If
				
				If FamilyMember.Artikelnummer = "NA" Then
				FamilyMember.Artikelnummer = FamilyMember.Artikelnummer3
				End If
				
				If FamilyMember.Artikelnummer = "NA" Then
				FamilyMember.Artikelnummer = FamilyMember.AIMD_PARTNO
				End If
				
				If FamilyMember.Artikelnummer = "NA" Then
				FamilyMember.Artikelnummer = FamilyMember.Artikel
				End If
				
				
				If FamilyMember.Artikelnummer = "" Then
				FamilyMember.Artikelnummer = "-"
				End If
				
				oSheet.Range("B" & oCount).Value = FamilyMember.Artikelnummer
                oSheet.Range("C" & oCount).Value = FamilyMember.Abmessung
				oSheet.Range("D" & oCount).Value = FamilyMember.Dateiname
				oSheet.Range("E" & oCount).Value = FamilyMember.ArtikelnummerDis


                oCount = oCount + 1
            Next
        Next
		
		oExcel = Nothing
    End Sub
