VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "i2c bus spy"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Advanced"
      Height          =   1335
      Left            =   2880
      TabIndex        =   14
      Top             =   3480
      Width           =   3615
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1800
         ScaleHeight     =   975
         ScaleWidth      =   1695
         TabIndex        =   21
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton Command10 
            Caption         =   "Read"
            Height          =   375
            Left            =   0
            TabIndex        =   23
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Write"
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Text            =   "208"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Value"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "clk"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Text            =   "208"
      Top             =   1005
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Send decimal"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "read"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   12675
      TabIndex        =   9
      Top             =   1440
      Width           =   12735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "0"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "drop SDA between clocks"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send character"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1000
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   400
      TabIndex        =   2
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop sequence"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start sequence"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   7920
      TabIndex        =   25
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Decoded bcd:"
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim draw_pos As Integer
Dim draw_inc As Integer

Dim old_sda As Boolean
Dim old_scl As Boolean

Dim shouted As Boolean

'Binary -> decimal
Public Function Bin2Dec(ByVal bin As String) As String
    Dim cont As Long
    Dim bstr As String
    Dim step As Long
    Dim neg As Boolean
    step = 1
    cont = 0
    If Left(bin, 1) = "-" Then
        neg = True
        bstr = Right(bin, Len(bin) - 1)
    Else
        neg = False
        bstr = bin
    End If
    While Left(bstr, 1) = "0"
        bstr = Right(bstr, Len(bstr) - 1)
    Wend
    If Len(bstr) = 0 Then
        Bin2Dec = 0
        Exit Function
    End If
    While Len(bstr) > 0
        If Right(bstr, 1) = "1" Then
            cont = cont + step
        End If
        bstr = Left(bstr, Len(bstr) - 1)
        step = step * 2
    Wend
    If neg = True Then
        cont = "-" & cont
    End If
    Bin2Dec = cont
End Function
Function Dec2Bin(ByVal n As Long) As String
Do Until n = 0
    If (n Mod 2) Then Dec2Bin = "1" & Dec2Bin Else Dec2Bin = "0" & Dec2Bin
    n = n \ 2
Loop
For i = 1 To 8 - Len(Dec2Bin)
    Dec2Bin = "0" & Dec2Bin
Next i
End Function
Private Function clear_area()
Picture1.Cls
Picture1.BackColor = RGB(255, 255, 255)
Picture1.CurrentX = 100
Picture1.CurrentY = 400
Picture1.Print "SDA"

Picture1.CurrentX = 100
Picture1.CurrentY = 1200
Picture1.Print "SCL"

draw_pos = 500
End Function

Private Function clkout()
scl (0)
sda (1)
get_bit_and_draw
delay = HScroll1.value
wait (delay)

scl (1)
res = get_bit_and_draw(True)
'here we should see a 1 on sda or 0

scl (0)
get_bit_and_draw

clkout = res
End Function



Private Function draw_bits(sda As Boolean, scl As Boolean)
Picture1.Line (draw_pos, 300)-(draw_pos, 1700), RGB(230, 230, 230)

If old_sda <> sda Then Picture1.Line (draw_pos, 500)-(draw_pos, 700)
If old_scl <> scl Then Picture1.Line (draw_pos, 1500)-(draw_pos, 1300)

If sda = True Then
    Picture1.Line (draw_pos, 500)-(draw_pos + draw_inc, 500)
End If

If sda = False Then
    Picture1.Line (draw_pos, 700)-(draw_pos + draw_inc, 700)
End If


If scl = True Then
    Picture1.Line (draw_pos, 1300)-(draw_pos + draw_inc, 1300)
End If

If scl = False Then
    Picture1.Line (draw_pos, 1500)-(draw_pos + draw_inc, 1500)
End If


If old_sda = True And sda = False And scl = True Then
    'start sequence
    Picture1.CurrentX = draw_pos
    Picture1.CurrentY = 100
    Picture1.Print "<"
End If

If old_sda = False And sda = True And scl = True Then
    'stop sequence
    Picture1.CurrentX = draw_pos
    Picture1.CurrentY = 100
    Picture1.Print ">"
End If

If old_sda = sda And scl = True And scl = True And old_scl = False Then
    'bit and transition
    Picture1.CurrentX = draw_pos
    Picture1.CurrentY = 120
    If sda = True Then Picture1.Print "1" Else Picture1.Print "0"
End If

old_sda = sda
old_scl = scl

draw_pos = draw_pos + draw_inc
End Function


Private Function get_bit() As String
MSComm1.Output = "z" 'request a bit
While MSComm1.InBufferCount = 0
    DoEvents
Wend
get_bit = MSComm1.Input
End Function


Private Function get_bit_and_draw(Optional sda_sdl As Boolean = True)
Dim c_scl As Boolean
Dim c_sda As Boolean

z = Dec2Bin(Asc(get_bit))
If Mid(z, 1, 1) = "1" Then c_scl = True Else c_scl = False
If Mid(z, 5, 1) = "1" Then c_sda = True Else c_sda = False
draw_bits c_sda, c_scl

If sda_sdl = True Then get_bit_and_draw = c_sda
If sda_sdl = False Then get_bit_and_draw = c_sdl
End Function

Private Function graphing_space()
get_bit_and_draw
get_bit_and_draw
get_bit_and_draw
End Function

Private Function idlebus()
sda (1)
scl (1)
End Function

Private Function read_byte(Optional last As Boolean = False)
Dim response As String

For i = 1 To 8
    n = clkout
    If n = True Then response = response & "1" Else response = response & "0"
Next i

high = Bin2Dec(Mid(response, 1, 4))
low = Bin2Dec(Mid(response, 5, 4))
decoded_bcd = Int(high & low)
Label6.Caption = decoded_bcd

response = Bin2Dec(response)

send_bit (last)
'pull line low to send ACK bit (or high, if it's the end of data)

read_byte = response
End Function

Public Function sda(value As Integer)
If value = 1 Then MSComm1.Output = "a"
If value = 0 Then MSComm1.Output = "b"
End Function

Public Function scl(value As Integer)
If value = 1 Then MSComm1.Output = "c"
If value = 0 Then MSComm1.Output = "d"
End Function
Private Function send_bit(bit As Boolean)
delay = HScroll1.value
scl (0)

If bit = True Or bit = 1 Then sda (1) Else sda (0)
wait (delay)
get_bit_and_draw

scl (1)
wait (delay)
get_bit_and_draw

scl (0)
wait (delay)
get_bit_and_draw

If Check1.value = 1 Then sda (0)
End Function

Private Function send_byte(newbyte)
z = Dec2Bin(newbyte)
For i = 1 To Len(z)

    If Mid(z, i, 1) = "1" Then send_bit (True) Else send_bit (False)

Next i

scl (0)
End Function

Private Function start_seq()
delay = HScroll1.value

'bring them in the actual position for a start sequence
scl (0) 'we need the clock to be low before changing data

sda (1)
wait (delay)
get_bit_and_draw

'raise the clock, now we have the initial position
scl (1)
wait (delay)
get_bit_and_draw

'this is the actual sequence. sda goes low while clock is high
sda (0)
wait (delay)
get_bit_and_draw

scl (0)
get_bit_and_draw
End Function

Private Function stop_seq()
delay = HScroll1.value
'a stop sequence starts with both pins low. make sure they are
scl (0)
sda (0) 'remember, no altering sda before the clock is low
get_bit_and_draw
wait (delay)

'start the sequence, first raise the clock
scl (1)
get_bit_and_draw
wait (delay)
sda (1) 'this is the second exception of clock transitions. a stop sequence pulls sda high
        'while the scl is high
get_bit_and_draw

idlebus
get_bit_and_draw
End Function

Private Sub Command1_Click()
start_seq
End Sub

Private Sub Command10_Click()
Form1.Enabled = False
clear_area

stop_seq
start_seq

addy = Text3.Text
reg = Text4.Text
Text5.Text = ""

'address device
send_byte (addy)


'wait acknowledge
z = clkout
If z = True Then
    MsgBox "Device " & Str(addy) & "(" & Str(Int(addy / 2)) & " ) did not respond", vbExclamation, "i2c spy"
    Form1.Enabled = True
    Exit Sub
End If

'space
graphing_space

'send address byte
send_byte (reg)
z = clkout
If z = True Then
    MsgBox "Device " & Str(addy) & "(" & Str(Int(addy / 2)) & " ) did not acknowledge register " & Str(reg), vbExclamation, "i2c spy"
    Form1.Enabled = True
    Exit Sub
End If

'space
graphing_space

'second start (restart)
start_seq

'address the device, this time for writing, so it will be an odd byte
addy = addy + 1

'address device in write mode
send_byte (addy)


'wait acknowledge
z = clkout
If z = True Then
    MsgBox "Device " & Str(addy - 1) & "(" & Str(Int((addy - 1) / 2)) & " ) did not respond in write mode", vbExclamation, "i2c spy"
    Form1.Enabled = True
    Exit Sub
End If

'space
graphing_space

'read byte
n = read_byte(True)

Text5.Text = n

Form1.Enabled = True
stop_seq
End Sub

Private Sub Command2_Click()
stop_seq
End Sub


Private Sub Command3_Click()
newbyte = Text1.Text
If newbyte = "" Then Exit Sub
send_byte (Asc(newbyte))
End Sub

Private Sub Command4_Click()
send_bit (1)
End Sub

Private Sub Command5_Click()
send_bit (0)
End Sub


Private Sub Command6_Click()
get_bit_and_draw
End Sub

Private Sub Command7_Click()
newbyte = Text2.Text
If newbyte = "" Then Exit Sub

z = Dec2Bin(newbyte)
For i = 1 To Len(z)

    If Mid(z, i, 1) = "1" Then send_bit (True) Else send_bit (False)

Next i

scl (0)
End Sub

Private Sub Command8_Click()
Form1.Caption = clkout
End Sub

Private Sub Command9_Click()
Form1.Enabled = False
clear_area

stop_seq
start_seq

addy = Text3.Text
reg = Text4.Text
newval = Text5.Text

'address device
send_byte (addy)


'wait acknowledge
z = clkout
If z = True Then
    MsgBox "Device " & Str(addy) & "(" & Str(Int(addy / 2)) & " ) did not respond", vbExclamation, "i2c spy"
    Form1.Enabled = True
    Exit Sub
End If

'space
graphing_space

'send register byte
send_byte (reg)
z = clkout
If z = True Then
    MsgBox "Device " & Str(addy) & "(" & Str(Int(addy / 2)) & " ) did not acknowledge register " & Str(reg), vbExclamation, "i2c spy"
    Form1.Enabled = True
    Exit Sub
End If

'space
graphing_space

'send data byte
send_byte (newval)
z = clkout
If z = True Then
    MsgBox "Device " & Str(addy) & "(" & Str(Int(addy / 2)) & " ) did not acknowledge after writing", vbExclamation, "i2c spy"
    Form1.Enabled = True
    Exit Sub
End If

'space
graphing_space

stop_seq
Form1.Enabled = True
End Sub

Private Sub Form_Load()
With MSComm1
   .Handshaking = comNone
   .EOFEnable = False
   .RTSEnable = False
   .Settings = "57600,n,8,1"
   .InBufferSize = 512
   .RThreshold = 256
   .PortOpen = True
End With

draw_inc = 75
clear_area

shouted = False ' notification message when the user clicks the i2c addy
End Sub


Public Function wait(dur As Integer)
start = Timer * 100
durend = start + dur

While (Timer * 100 < durend)
    DoEvents
Wend
End Function

Private Sub HScroll1_Change()
Label1.Caption = "delay: " & HScroll1.value
End Sub


Private Sub Picture1_Click()
clear_area
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub


Private Sub Text3_Click()
If shouted = False Then MsgBox "Please write the device address in 8 bit format, clearing bit 0 (R/W=W). the software will set this bit when it will need to write" & vbCrLf & "So, adressing a DS1307 realtime clock will be 208 (11010000) instead of 104 (1101000) + R/W", vbInformation, "8 bit address"
shouted = True
End Sub


