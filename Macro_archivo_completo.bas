Attribute VB_Name = "Module1"
Sub new_buildingid_excel()

Dim app As New Excel.Application
app.Visible = False

Dim new_excel As Excel.Workbook
Dim new_excel_sheet As Worksheet

Dim geography As String
Dim country As String
Dim owner As String
Dim buildingId As String
Dim company As String
Dim street As String
Dim address As String
Dim ruta_plantilla As String
Dim ruta_archivo_generado As String

Dim i As Integer
Dim row As Integer

ruta_plantilla = "C:\TEMPLATE.xlsx"
row = 2

Do Until row > 1276

    If (buildingId = Range("P" & row).Value) Then
            row = row + 1
        Else
            geography = Range("A" & row).Value
            country = Range("B" & row).Value
            owner = Range("AP" & row).Value
            buildingId = Range("P" & row).Value
            company = Range("W" & row).Value
            street = Range("AA" & row).Value
            address = Range("Z" & row).Value & ", " & Range("AB" & row).Value
            
            If (Range("Y" & row).Value <> vbNullString) Then
                address = address & ", " & Range("Y" & row).Value
            End If
            
            ruta_archivo_generado = "C:\Files\"
            
            If (owner = vbNullString) Then
                ruta_archivo_generado = ruta_archivo_generado & geography & " - " & country & " - " & buildingId & " - " & company & ".xlsx"
            Else:
                ruta_archivo_generado = ruta_archivo_generado & geography & " - " & country & " - " & owner & " - " & buildingId & " - " & company & ".xlsx"
            End If
            
            FileCopy ruta_plantilla, ruta_archivo_generado
            
            Set new_excel = app.Workbooks.Open(ruta_archivo_generado)
            Set new_excel_sheet = new_excel.Sheets("Offsite form")
            
            new_excel_sheet.Range("E5").Value = company
            new_excel_sheet.Range("E9").Value = buildingId
            new_excel_sheet.Range("E10").Value = street
            new_excel_sheet.Range("E12").Value = address
            new_excel_sheet.Range("E13").Value = country
            
            i = 0
            
            Do Until buildingId <> Range("P" & row).Offset(i, 0).Value
            
                new_excel_sheet.Range("C17").Offset(i, 0).Value = Range("D" & row).Offset(i, 0).Value
                new_excel_sheet.Range("D17").Offset(i, 0).Value = Range("L" & row).Offset(i, 0).Value
                new_excel_sheet.Range("E17").Offset(i, 0).Value = Range("M" & row).Offset(i, 0).Value
                new_excel_sheet.Range("F17").Offset(i, 0).Value = Range("F" & row).Offset(i, 0).Value
                
                i = i + 1
                
            Loop
            
            new_excel.Close SaveChanges:=True
            app.Quit
            Set app = Nothing
            
            row = row + 1
        
        End If
    Loop
End Sub
