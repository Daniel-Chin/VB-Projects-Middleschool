VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Web 
   Caption         =   "WEB"
   ClientHeight    =   5412
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9852
   LinkTopic       =   "Form1"
   ScaleHeight     =   5412
   ScaleWidth      =   9852
   StartUpPosition =   1  '����������
   Begin VB.TextBox Prompt 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4572
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   4092
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   9480
      Top             =   360
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      LocalPort       =   1875
   End
End
Attribute VB_Name = "Web"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�����Ƿ�װ���룬�����������
Option Explicit

Const MaxConnection As Integer = 999
Public Port As Integer
Private Gate(1 To MaxConnection) As Boolean

Public Sub Listen(Optional NewPort As Integer = -1)
If NewPort <> -1 Then
    Port = NewPort
End If
With Sock(0)
    .LocalPort = Port
    .Listen
End With
Pop "��ʼ�������˿�: " & Port
End Sub

Private Sub Form_Resize()
Prompt.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim k As Integer
For k = 1 To Sock.ubound
    If Gate(k) Then
        Sock(k).Close
    End If
Next k
Sock(0).Close
End Sub

Private Sub Sock_Close(Index As Integer)
Gate(Index) = False
Unload Sock(Index)
Pop String(Index - 1, "-") & "�Ͽ�: ID " & Index
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim k As Integer
Do Until k = MaxConnection
    k = k + 1
    If Not Gate(k) Then
        Exit Do
    End If
Loop
If k = MaxConnection Then
    Pop "�����������������µ�����"
Else
    Load Sock(k)
    Gate(k) = True
    Sock(k).Accept requestID
    Pop "����: ID " & k
    DoEvents
End If
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Sock(Index).GetData strData, vbString
Pop String(Index - 1, "-") & "From ID " & Index & ": " + strData
Dim FirstLen As Integer
strData = Right(strData, Len(strData) - 1)
FirstLen = InStr(strData, Chr(2))
Do Until FirstLen = 0
    ReadData Index, Left(strData, FirstLen - 1)
    strData = Right(strData, Len(strData) - InStr(strData, Chr(2)))
    FirstLen = InStr(strData, Chr(2))
Loop
ReadData Index, strData
End Sub

Private Sub Pop(MSG As String)
Prompt = Prompt & MSG & vbCrLf
Prompt.SelStart = Len(Prompt)
Debug.Print MSG
End Sub

'�����Ƿ�װ���룬����������ġ�
'�����Ǳ�ģ���ʹ��

Private Sub ReadData(Index As Integer, strData As String)
'����¼��������Կͻ��˵���Ϣ���յ�����Ϣ����strData��
'Index���ǿͻ��˵�ID��
End Sub

Public Function IsGate(ID As Integer) As Boolean
'��������жϱ��ΪID�Ŀͻ����Ƿ�������
IsGate = Gate(ID)
End Function

Public Sub Send(ToWho As Integer, strData As String)
'����Send��������ͻ��˷����ı���Ϣ��
'ToWho��Ŀ��ͻ��˵ı�ţ�strData����Ϣ����
If Gate(ToWho) Then
    DoEvents
    Sock(ToWho).SendData Chr(2) & strData
    DoEvents
    Pop String(ToWho - 1, "-") & "To ID " & ToWho & ": " & strData
Else
    Pop "����Ŀ���ѶϿ�����: " & ToWho & "���޷�����" & strData
End If
End Sub
