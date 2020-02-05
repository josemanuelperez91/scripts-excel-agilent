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

ruta_plantilla = "C:\Users\Jose Manuel Perez\Google Drive\180119 - App Physical Inventory\TEMPLATE.xlsx"
ruta_archivo_generado = "C:\Users\Jose Manuel Perez\Google Drive\180119 - App Physical Inventory\Files\"

' Set master = Workbooks.Open("C:\Users\Jose Manuel Perez\Google Drive\180119 - App Physical Inventory\MASTER.xlsx")
' Set master_sheet = master.Sheets("Asset to be counted - macro")

geography = Range("A6").Value
country = Range("B6").Value
owner = Range("AP6").Value
buildingId = Range("P6").Value
company = Range("W6").Value
street = Range("AA6").Value
address = Range("Z6").Value & ", " & Range("AB6").Value

If (Range("Y6").Value <> vbNullString) Then
    address = address & ", " & Range("Y6").Value
End If

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

Do Until buildingId <> Range("P6").Offset(i, 0).Value

    new_excel_sheet.Range("C17").Offset(i, 0).Value = Range("D6").Offset(i, 0).Value
    new_excel_sheet.Range("D17").Offset(i, 0).Value = Range("L6").Offset(i, 0).Value
    new_excel_sheet.Range("E17").Offset(i, 0).Value = Range("M6").Offset(i, 0).Value
    new_excel_sheet.Range("F17").Offset(i, 0).Value = Range("F6").Offset(i, 0).Value
    
    i = i + 1
Loop

new_excel.Close SaveChanges:=True
app.Quit
Set app = Nothing
End Sub
