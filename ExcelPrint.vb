'This module must import the microsoft office interoperation namespace.
Imports Microsoft.Office.Interop

'#####################################################################################
'ExcelPrint Module
'#####################################################################################
'This module is designed to allow a designer to quickly and easily take all the items and subitems of a listview and place them
'in a formatted Excel spreadsheet. Additionally, this module will allow the designer to save the Excel document under a specific
'name and this module will allow a designer to print the Excel document after it has been created. In the event that the designer
'specifies a report name for this document to save, the Excel application will be run as an invisible background process. Otherwise
'the Excel application will be made visible and will open on the user's desktop.
'#####################################################################################

Module ExcelPrint
	Public Sub MakeExcelDoc(ByVal lv As ListView, Optional visible as boolean = True, Optional rptpath as string = "", 
			Optional printrpt as boolean = false)
        Try
            'Create a new excel application, workbook, and worksheet
            Dim oXL As New Excel.Application
            Dim oWB As Excel.Workbook = oXL.Workbooks.Add
            Dim osheet As Excel.Worksheet = oWB.ActiveSheet
            Dim i As Integer
            
            'Add Listview Data
            For i = 0 To lv.Columns.Count - 1
                'Make the header row (the first row) values equal to the values in the listview columns.
                osheet.Cells(1, i + 1) = lv.Columns(i).Text
            Next
            
            Dim count As Integer = 0
            For i = 0 To lv.Items.Count - 1
                For j as integer = 0 To lv.Items(i).SubItems.Count - 1
                    'Add Text To Cells From Listview items and subitems (listview subitem 0 is the same as listview.text)
                    osheet.Cells(i + 2, j + 1) = lv.Items(i).SubItems(j).Text
                Next
            Next
            'End Adding Listview Data

            'Determine the number of listview items and the last column value
            Dim coun As Integer = lv.Items.Count
            Dim lastchar As Char = Chr(64 + lv.Columns.Count)
            
            'Combine the last column value with the number of listview items to grab the name of the last cell that had text
            'entered.
            Dim last As String = lastchar & coun + 1
            
            'Format the header row so all the values are bold, the background color is light gray, a balck border
            'surrounds every particular header and the vertical and horizontal alignments are set to be centered.
            With osheet.Range("A1", lastchar & "1")
                .Font.Color = Color.Black
                .Interior.Color = Color.LightGray
                .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            'Format all cells with values in them to have a font type of arial and a font size of 12.
            With osheet.Range("A1", last)
                .Font.Name = "Arial"
                .Font.Size = 12
                .Columns.AutoFit()
            End With

            '''''''''''Page Setup values
            With osheet.PageSetup
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .LeftMargin = 0.25
                .RightMargin = 0.25
                'Ensure that the first row is printed on all pages
                .PrintTitleRows = "$1:$1"
                'Set the right header equal to today's date
                .RightHeader = Date.Today.ToLongDateString
                'Set the Right Footer equal to the current page number
                .RightFooter = "Page &P of &N"
            End With

            'If a report name is provided...
			If rptpath <> "" Then
				'Grab the user's documents directory and determine the excel file path.
                Dim yourdir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\"
                Dim filename As String = yourdir & saveasrptname & ".xlsx"

                'Save the excel document
                oWB.SaveAs(Filename:=filename)
                
				MessageBox.Show("Report Saved")
            End If

			'If print is equal to true...
            If print = True Then
                'print a copy of the excel document
                oWB.PrintOutEx()

				MessageBox.Show("Printing Report")

                'Provide a 2 second buffer timer to ensure that the print command gets loaded into the print queue
                Task.Delay(2000).Wait()
            End If

			oXL.Visible = visible
			
			'If visible is set to false, close the workbook
			If visible = false
				oWB.Close(False)
			End If

        Catch ex As Exception
            'If anything goes awry during this process, display the error message on the user's machine.
            MsgBox(ex.Message)
        End Try
    End Sub
End Module
