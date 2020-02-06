Sub login()
Dim wbForm As Workbook
Set wbForm = ActiveWorkbook


Const Url$ = "https://moj.farmaprom.pl"

    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")

    With ie

        .navigate Url
        ieBusy ie
        .Visible = True

    'Set Elements = ie.document.getElementsByTagname("button")
   ' For Each e In Elements
    'If (e.getAttribute("id") = "submit") Then
    'e.Click
    'End If
    'Exit For
    'Next e

    End With
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    With ie
     Set Elements = ie.document.getElementsByTagname("button")
     For Each e In Elements
     If (e.innerText = "Zaloguj") Then
     e.Click
     Exit For
     
     End If
     Next e

     'Exit For
     
     End With
     
    Application.Wait (Now + TimeValue("0:00:03"))
    
     With ie
     Set Elements = ie.document.getElementsByTagname("span")
     For Each e In Elements
     If (e.innerText = "Raporty") Then
     e.Click
     Exit For
     
     End If
     Next e

     'Exit For
     
     End With
     Application.Wait (Now + TimeValue("0:00:03"))
     
     With ie
     Set Elements = ie.document.getElementsByTagname("span")
     For Each e In Elements
     If (e.innerText = "Liczba wizyt na klienta") Then
     e.Click
     Exit For
     
     End If
     Next e
     
     End With
     
     Application.Wait (Now + TimeValue("0:00:03"))
     
     With ie
     Set Elements = ie.document.getElementsByTagname("a")
     For Each e In Elements
     If (e.innerText = "Eksportuj do CSV") Then
     
     e.Click
     Exit For
     
     End If
     Next e
     
     End With
     
     'MsgBox ("   xxxxxxxxx" & vbNewLine & " xxxxxxxxx")
    

End Sub

Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.readyState < 4
        DoEvents
    Loop
End Sub