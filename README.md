<div align="center">

## Animation Status bar


</div>

### Description

Easy! without API. Only a few lines of code.You can scroll your Status bar panel's text and You can modify suit to you.I hope you enjoy & your correctons,comments,Ideas are very gratefully.

G.B.R.Nilantha Athurupana.

Wattarama,Imbulgasdeniya,Sri Lanka.(94-03744479, E-Mail: crysbro@cga.lk)
 
### More Info
 
'Create a new form (Any Name you can use). Add a Timer control (Timer1)Required 'follwing ActiveX OCX to create a Status Bar Control(It name must be (SB)

'Microsoft Windows Common controls 5.0(SP2)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[G\.B\.R\.Nilantha Athurupana](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/g-b-r-nilantha-athurupana.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/g-b-r-nilantha-athurupana-animation-status-bar__1-2742/archive/master.zip)





### Source Code

```
'General Deciarations
Dim C As String 'to store panel's text
Dim CO As Integer 'to store text length
Dim FS As Long 'to store Panels Width
Private Sub MDIForm_Load()
 Timer1.Interval = 100
 SB.Panels(1).Text = "Nilantha Athurupana"
 C = SB.Panels(1).Text
 CO = Len(C) + 1
 SB.Panels(1).Text = ""
 FS = SB.Panels(1).Width
End Sub
Private Sub Timer1_Timer()
On Error GoTo ATH
 Static C01 As Integer ' Counter 1
 Static CO2 As Integer ' Counter 2
 Static A As String 'To move text
 Dim R As String 'Restore text
 Dim T As String 'Restore text
XX:
 If CO > 0 Then
  C01 = CO
  T = Mid(C, C01, 1)
  CO = CO - 1
  R = " "
  Mid(R, 1) = T
  SB.Panels(1).Text = R & SB.Panels(1).Text
 Else
  A = A & " "
  R = " "
  Mid(R, 1) = A
  SB.Panels(1).Text = R & SB.Panels(1).Text
 End If
 If CO2 >= FS Then
  CO2 = 0
  CO = Len(C)
  SB.Panels(1).Text = ""
  GoTo XX
 Else
  CO2 = CO2 + 35 'please edit this value according to your text length.
 End If
 Exit Sub
ATH:
End Sub
```

