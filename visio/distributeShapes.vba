Sub DistributeComponents()
     
' Listing all pages
'    Dim page As Visio.page
'    Dim pages As Visio.pages
'    Dim document As Visio.document
'    Dim documents As Visio.documents
'
'    Set documents = Application.documents
'
'    For Each document In documents
'        Debug.Print document.FullName
'       Set pages = document.pages
'       For Each page In pages
'           Debug.Print Tab(5); page.Name
'       Next
'    Next

    
Dim selection As Visio.selection
Dim shape As Visio.shape
Set selection = Visio.ActiveWindow.selection

Dim y As Double, x As Double, xd As Double, yd As Double
Dim columns As Integer, i As Integer

columns = 12
i = 1
xd = 45 / 25.4
yd = 25 / 25.4

For Each shape In selection
    Debug.Print shape.Cells("PinX") & " " & shape.Cells("PinX")
    Debug.Print "x: " & x; Tab(3); "y: " & y
    If i Mod columns = 0 Then y = y + yd
    x = xd * (i Mod columns)
    shape.SetCenter x, y
    i = i + 1
Next

    
' Manual distribution
    
'    Dim y As Double, x As Double, xd As Double, yd As Double
'    Dim columns As Integer
'    Dim none As Visio.Shape
'
'    Visio.ActivePage.AutoSize = True
'
'    Set columns = 16
'    Set xd = 1.8
'    Set yd = 1
'    Set x = 0
'    Set y = 0
'
'    If Visio.ActiveWindow.Selection.Count > 0 Then
'
'        For i = 1 To Visio.ActiveWindow.Selection.Count
'            If i Mod columns = 0 Then y = y + yd
'            x = xd * (i Mod columns)
'            Visio.ActiveWindow.Selection(i).Cells("pinx") = x
'            Visio.ActiveWindow.Selection(i).Cells("piny") = y
'        Next
'
'    Else
'        MsgBox "No shapes selected. Nothing done." ' soft fail
'    End If

'Visio.ActiveWindow.Selection.Distribute visDistVertRight, False


End Sub
