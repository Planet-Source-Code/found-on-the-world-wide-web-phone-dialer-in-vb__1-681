<div align="center">

## Phone Dialer in VB


</div>

### Description

Phone Dialer in VB for windows 95. 'thanks to Andre Obelink

'De Visual Basic Groep

'http://www.plus.nl/vbg/
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-phone-dialer-in-vb__1-681/archive/master.zip)

### API Declarations

Private Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAdress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)


### Source Code

```
'make a new project; 2 textboxen (index 0 & 1); 2 labels (index 0 & 1)
'1 command button
'Insert the next code in the right place (use Insert/File)
'Press F5
------------- code -------------------
Private Sub ChooseNumber(strNumber As String, strAppName As String, strName As String)
  Dim lngResult As Long
  Dim strBuffer As String
  lngResult = tapiRequestMakeCall&(strNumber, strAppName, strName, "")
  If lngResult <> 0 Then 'error
    strBuffer = "Error connecting to number: "
    Select Case lngResult
    Case -2&
      strBuffer = strBuffer & " 'PhoneDailer not installed?"
    Case -3&
      strBuffer = strBuffer & "Error : " & CStr(lngResult) & "."
    End Select
    MsgBox strBuffer
  End If
End Sub
Private Sub Command1_Click()
  Call ChooseNumber(Text1(0).Text, "PhoneDialer", Text1(1).Text)
End Sub
Private Sub Form_Load()
  Text1(0).Text = ""
  Text1(1).Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
```

