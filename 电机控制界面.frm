VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   " "
   ClientHeight    =   10248
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   18240
   ClipControls    =   0   'False
   FillColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10248
   ScaleWidth      =   18240
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8400
      TabIndex        =   33
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8400
      TabIndex        =   32
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8400
      TabIndex        =   24
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000013&
      Caption         =   "״̬��ʾ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.4
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   11760
      TabIndex        =   20
      Top             =   2880
      Width           =   4695
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         Caption         =   " ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   2160
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000018&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000018&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Run 
      BackColor       =   &H80000014&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton ZR 
      BackColor       =   &H80000014&
      Caption         =   "��ת"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   18
      Top             =   8280
      Width           =   975
   End
   Begin VB.CommandButton Stop 
      Caption         =   "ͣ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton FR 
      Caption         =   "��ת"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14520
      TabIndex        =   16
      Top             =   8280
      Width           =   975
   End
   Begin VB.CommandButton BRK 
      Caption         =   "�ƶ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   8280
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "ģ�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4572
      Left            =   4200
      TabIndex        =   0
      Top             =   2280
      Width           =   6375
      Begin VB.Frame Frame3 
         BackColor       =   &H80000013&
         Caption         =   "�����趨"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.4
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3732
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   3012
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   960
            TabIndex        =   2
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   960
            TabIndex        =   1
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton SET 
            Caption         =   "ȷ��"
            Height          =   375
            Left            =   960
            TabIndex        =   12
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "ȱʡ-����"
            Height          =   180
            Left            =   2160
            TabIndex        =   37
            Top             =   2520
            Width           =   816
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "1-���� "
            Height          =   180
            Left            =   2160
            TabIndex        =   36
            Top             =   2280
            Width           =   648
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000013&
            Caption         =   "  ����"
            Height          =   492
            Left            =   120
            TabIndex        =   35
            Top             =   2640
            Width           =   732
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000013&
            Caption         =   "  ����"
            Height          =   372
            Left            =   120
            TabIndex        =   34
            Top             =   2160
            Width           =   732
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000013&
            Caption         =   "r/min"
            Height          =   252
            Left            =   2040
            TabIndex        =   31
            Top             =   1200
            Width           =   732
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000013&
            Caption         =   "r/min"
            Height          =   252
            Left            =   2040
            TabIndex        =   30
            Top             =   1680
            Width           =   852
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000013&
            Caption         =   "����ת��"
            Height          =   372
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   852
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000013&
            Caption         =   "����ת��"
            Height          =   732
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   852
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   1800
            TabIndex        =   13
            Top             =   720
            Width           =   852
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "ʵʱ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.4
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2895
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1080
            TabIndex        =   3
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Textsend1 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ���ȡ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   10
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   9
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "ת��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "r/min"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   7
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   615
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   10800
      Top             =   7560
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10680
      Top             =   6720
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      InBufferSize    =   11
      OutBufferSize   =   22
      RThreshold      =   1
      SThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "r/min"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   9480
      TabIndex        =   27
      Top             =   4080
      Width           =   612
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "r/min"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9480
      TabIndex        =   26
      Top             =   4080
      Width           =   612
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "r/min"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9480
      TabIndex        =   25
      Top             =   4080
      Width           =   612
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "������ϵͳ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5880
      TabIndex        =   23
      Top             =   480
      Width           =   8640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   2760
      X2              =   2760
      Y1              =   2040
      Y2              =   9600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   17760
      X2              =   2760
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   17760
      X2              =   17760
      Y1              =   2040
      Y2              =   9600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   2760
      X2              =   17760
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   852
      Left            =   12720
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   852
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   852
      Left            =   6000
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   852
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000005&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   852
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   852
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   852
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private savetime As Double


Sub timeover(JianGe As Long)
     'ʱ����
     savetime = timeGetTime
     While timeGetTime < savetime + JianGe
     DoEvents
     Wend
End Sub




Private Sub MSComm1_OnComm()
 Select Case MSComm1.CommEvent
   Case comEvCD
   Case comEvCTS
   Case comEvDSR
   Case comEvRing
   Case comEvReceive
   Case comEvSend
 End Select
End Sub
Public Function DEC_to_HEX(Dec As Long) As String
    Dim a As String
    DEC_to_HEX = ""
    Do While Dec > 0
        a = CStr(Dec Mod 16)
        Select Case a
            Case "10": a = "A"
            Case "11": a = "B"
            Case "12": a = "C"
            Case "13": a = "D"
            Case "14": a = "E"
            Case "15": a = "F"
        End Select
        DEC_to_HEX = a & DEC_to_HEX
        Dec = Dec \ 16
    Loop
End Function

                                '''''''''''''''''�ֵ����'''''''''''''''''δ����Ӧ��'''''''

Private Sub SET_Click()
 Dim Send_Current(8) As Byte            'Ҫ���͵ĵ�������
 
 Dim Send1_RPM(8) As Byte               'Ҫ���͵Ĺ���ת������                             !!!!!
 Dim Send0_RPM(8) As Byte               'Ҫ���͵Ĵ���ת������                             !!!!!
 
 Dim Send_up(8) As Byte               'Ҫ���͵���������                             !!!!!
 Dim Send_down(8) As Byte               'Ҫ���͵Ľ�������                             !!!!!
 
 
 
 
 Dim C_data(8) As Byte             '���巢�͵�������������
 
 Dim R_data(8) As Byte           '���巢�͹���ת�ٲ���������
 
 Dim d_data(8) As Byte                                                        '���巢�ʹ���ת�ٲ���������  !!!!!!
 
 Dim up_data(8) As Byte                                '���������趨
 Dim down_data(8) As Byte                             '���彵���趨
 
 Dim Answer(6) As String         '������ȷ�ظ�������
 
 
 Dim i As Integer
 Dim j As Integer
 Dim k As Integer
 Dim m As Integer
 Dim l As Integer
 Dim u As Integer
 Dim w As Integer
 
  
 Dim Current As Single       'Ҫ���õĵ���
  Current = Text2.Text
 Dim RPM1 As Long             'Ҫ���õĹ���ת��
 Dim RPM0 As Long             'Ҫ���õĴ���ת��
  
  If Text3.Text = "" Then MsgBox "���趨����ת��"
     
  If Text4.Text = "" Then MsgBox "���趨����ת��"
  
  RPM1 = Text3.Text
  RPM0 = Text4.Text
  
  Dim up As Single
     If Text5.Text = "1" Then up = 1
     Else: up = 0
     End If
     
  Dim down As Single
     If Text6.Text = "1" Then down = 1
     Else: down = 0
     End If
     
 
 Dim Len_Current As Integer   '�����û�����ĵ����ı��ĳ���
 Dim Len_RPM1 As Integer       '�����û�����Ĺ���ת���ı��ĳ���
 Dim Len_RPM0 As Integer       '�����û�����Ĵ���ת���ı��ĳ���
 Dim crt_str As String        '�������ת��Ϊ�ַ���
 Dim rpm1_str As String        '���幤��ת��ת��Ϊ�ַ���
 Dim rpm0_str As String        '�������ת��ת��Ϊ�ַ���
 

 
 
 If Current > 20 Or Current < 0.2 Then
   MsgBox "�����ķ�ΧΪ0.2��20��"
 End If                          '�趨ֵ�ķ�Χ
 
 If RPM1 > 1500 Or RPM1 < 120 Then
   MsgBox "ת�ٵķ�ΧΪ120��1500��"
 End If
 If RPM0 > 1500 Or RPM0 < 120 Then
   MsgBox "ת�ٵķ�ΧΪ120��1500��"
 End If
  
crt_str = DEC_to_HEX(Current * 10)                                          'ת�ٺ͵����Ķ��嵽���Ƕ��٣���������������������������
rpm1_str = DEC_to_HEX(RPM1 / 12)
rpm0_str = DEC_to_HEX(RPM0 / 12)
 
 
 Len_Current = Len(crt_str)          '�õ��û�����ĵ����ı��ĳ���
 Len_RPM1 = Len(rpm1_str)              '�õ��û�����Ĺ���ת���ı��ĳ���
 Len_RPM0 = Len(rpm0_str)              '�õ��û�����Ĵ���ת���ı��ĳ���
 
 If (Len_Current < 4) Then       '***************************
   For i = Len_Current To 3      '
    crt_str = "0" + crt_str    ' ������ĵ������ֲ�����λ��ǰ�油0
   Next i                        '
 End If                          '****************************
 
 
 If (Len_RPM1 < 4) Then           '***************************
   For i = Len_RPM1 To 3      '
     rpm_str = "0" + rpm1_str          ' ������Ĺ���ת�����ֲ�����λ��ǰ�油0
   Next i                           '
 End If
 If (Len_RPM0 < 4) Then           '***************************
   For i = Len_RPM0 To 3      '
     rpm_str = "0" + rpm0_str          ' ������Ĵ���ת�����ֲ�����λ��ǰ�油0
   Next i                           '
 End If
                            
 '***************�趨�����**********************
  C_data(0) = &HFA
  C_data(1) = &H1                                               '���͵�������ز���
  C_data(2) = &H57
  C_data(3) = &H2
  C_data(4) = "&H" + Left(crt_str, 2)                           '�õ������ĸ�2λ
  C_data(5) = "&H" + Right(crt_str, 2)                          '�õ������ĵ�2λ
  C_data(6) = C_data(1) Xor C_data(2) Xor C_data(3) Xor C_data(4) Xor C_data(5)
  C_data(7) = &HFB                                                '****************************************
  
  For j = 0 To 7
    Send_Current(j) = Val(C_data(j))
  Next j
 MSComm1.Output = Send_Current  '��com�ڷ��͵�����������
 
 '***************�趨����ת��**********************
  
 Dim R() As Byte
 Dim R_stuff As String
 Dim R_strdata As String
 Dim p As Integer
                                                                   
  R_data(0) = &HFA
  R_data(1) = &H1                                                 '����ת�ٵ���ز���
  R_data(2) = &H57
  R_data(3) = &H0
  R_data(4) = "&H" + Left(rpm1_str, 2)                                  '�õ�ת�ٵĸ�2λ
  R_data(5) = "&H" + Right(rpm1_str, 2)                                '�õ�ת�ٵĵ�2λ
  R_data(6) = R_data(1) Xor R_data(2) Xor R_data(3) Xor R_data(4) Xor R_data(5)
  R_data(7) = &HFB
  For m = 0 To 7
   Send1_RPM(m) = Val(R_data(m))                '���͵�ת�ٲ�����
  Next m
  
 MSComm1.Output = Send1_RPM     '��com�ڷ���ת�ٲ�������
 
 
  '***************�趨����ת��**********************                                         '!!!!!!
  
 Dim d() As Byte
 Dim d_stuff As String
 Dim d_strdata As String
 Dim h As Integer
                                                                   
  d_data(0) = &HFA
  d_data(1) = &H1                                                 '����ת�ٵ���ز���
  d_data(2) = &H57
  d_data(3) = &H1
  d_data(4) = "&H" + Left(rpm0_str, 2)                                  '�õ�ת�ٵĸ�2λ
  d_data(5) = "&H" + Right(rpm0_str, 2)                                '�õ�ת�ٵĵ�2λ
  d_data(6) = d_data(1) Xor d_data(2) Xor d_data(3) Xor d_data(4) Xor d_data(5)
  d_data(7) = &HFB
  For l = 0 To 7
   Send0_RPM(l) = Val(d_data(l))                '���͵�ת�ٲ�����
  Next l
  
 MSComm1.Output = Send0_RPM     '��com�ڷ���ת�ٲ�������
 
 '***************�趨���ٹ���**********************
  up_data(0) = &HFA
  up_data(1) = &H1
  up_data(2) = &H57
  up_data(3) = &H3
  up_data(4) = "&H" + 0
  up_data(5) = "&H" + up
  up_data(6) = up_data(1) Xor up_data(2) Xor up_data(3) Xor up_data(4) Xor up_data(5)
  up_data(7) = &HFB                                                '****************************************
  
  For u = 0 To 7
    Send_up(u) = Val(up_data(u))
  Next u
 MSComm1.Output = Send_up  '��com�ڷ��͵�����������
 
  '***************�趨���ٹ���**********************
  down_data(0) = &HFA
  down_data(1) = &H1
  down_data(2) = &H57
  down_data(3) = &H2
  down_data(4) = "&H" + 0
  down_data(5) = "&H" + down
  down_data(6) = down_data(1) Xor down_data(2) Xor down_data(3) Xor down_data(4) Xor down_data(5)
  down_data(7) = &HFB                                                '****************************************
  
  For w = 0 To 7
    Send_down(w) = Val(down_data(w))
  Next w
 MSComm1.Output = Send_down  '��com�ڷ��͵�����������
 
 

 MSComm1.OutBufferCount = 0   '��շ��ͻ�����
 MSComm1.InBufferCount = 0

End Sub

                                 '********************************��ת����*************Э����û��***********************
Private Sub ZR_Click()
 Dim clw_run(6) As Byte             '���巢�͵���ת�����ݰ�
 Dim S_data(6) As String          '���͵�����
 Dim i As Integer
  S_data(0) = &HFA
  S_data(1) = &H1               'ȥ��������
  S_data(2) = &H46
  S_data(3) = &H1
  S_data(4) = S_data(1) Xor S_data(2) Xor S_data(3)
  S_data(5) = &HFB
  For i = 0 To 5
   clw_run(i) = Val(S_data(i))
  Next i
 MSComm1.Output = clw_run
 
 MSComm1.OutBufferCount = 0   '��շ��ͻ�����
 MSComm1.InBufferCount = 0
End Sub
                                  '**************************************��ʱ��*********************************************
Private Sub Timer2_Timer()
Call getdata
End Sub
                                  '**************************************������������������������ʾ���ݣ�����������������������������������������ԭ����asm�ļ���ͬ�����޸�*********************************************
Sub getdata()
 Dim send_read(6) As Byte            '���巢�͵Ķ�ȡ���ݵ����ݰ�
 Dim buffer As String           '���յ����ݰ�
 Dim S_data(6) As Byte           '���͵�����
 
 Dim inbyte(11) As Byte          '������յ�����------------û���õ���
 
 Dim rev As Long            'ת��
 Dim Current As Long       '����
 Dim state As Byte
 Dim alarm As Byte
 Dim Pcurrent As String          'Ҫ���õĵ���
 Pcurrent = Text2.Text
 
 Dim i As Integer
 Dim s() As Byte
 Dim stuff As String
 Dim strdata As String
 Dim k As Integer
 Dim j As Integer
 
 
 
  S_data(0) = &HFA
  S_data(1) = &H1
  S_data(2) = &H4D
  S_data(3) = &H0
  S_data(4) = S_data(1) Xor S_data(2) Xor S_data(3)
  S_data(5) = &HFB
  For j = 0 To 5
    send_read(j) = Val(S_data(j))
  Next j
 MSComm1.Output = send_read        '��com�ڷ��Ͷ�ȡָ��

    MSComm1.InputLen = 0
    stuff = MSComm1.Input          '���շ��ص�����
    s() = stuff
  For k = 0 To UBound(s())
    If Len(Hex(s(k))) = 1 Then
   s_strdata = s_strdata & "0" & Hex(s(k))
   Else
   s_strdata = s_strdata & Hex(s(k))
   End If
  Next
      
   'If (Val("&H" & Mid(s_strdata, 5, 2)) = &H4D) Then
    ' rev = Val("&H" & Mid(s_strdata, 9, 2)) * 100 + Val("&H" & Mid(s_strdata, 13, 2))                'ת��
     'Current = Val("&H" & Mid(s_strdata, 21, 2))                                                     '����
     'Text1.Text = Current * 0.1
     'Textsend1.Text = rev
    'alarm = Val("&H" & Mid(s_strdata, 25, 2))                                                        '״̬��Ϣ
    
    ' If (alarm = &H20) Then
     ' Shape5.FillColor = QBColor(10)
     ' Shape8.FillColor = QBColor(7)                                                                  '��ת
    ' Else
     ' If (alarm = &H30) Then
   '    Shape8.FillColor = QBColor(10)
    '   Shape5.FillColor = QBColor(7)                                                                 '��ת
'      End If
 '   End If
  '   If (alarm = &H62 Or alarm = &H72) Then                                                          '����
   '   Shape1.FillColor = QBColor(12)
    ' End If
  '   If (alarm = &H24 Or alarm = &H34) Then                                                          '����
   '   Shape2.FillColor = QBColor(12)
    ' Else
 '     Shape2.FillColor = QBColor(7)
  '   End If
  ' End If
'End Sub

If (Val("&H" & Mid(s_strdata, 5, 2)) = &H4D) Then
     rev = Val("&H" & Mid(s_strdata, 7, 2)) * 100 + Val("&H" & Mid(s_strdata, 9, 2))                'ת��
     Current = Val("&H" & Mid(s_strdata, 13, 2))                                                     '����
     Text1.Text = Current * 0.1
     Textsend1.Text = rev
    alarm = Val("&H" & Mid(s_strdata, 17, 2))                                                        '״̬��Ϣ
    
     If (alarm = &H20) Then
      Shape5.FillColor = QBColor(10)
      Shape8.FillColor = QBColor(7)                                                                  '��ת
     Else
      If (alarm = &H30) Then
       Shape8.FillColor = QBColor(10)
       Shape5.FillColor = QBColor(7)                                                                 '��ת
      End If
    End If
     If (alarm = &H62 Or alarm = &H72) Then                                                          '����
      Shape1.FillColor = QBColor(12)
     End If
     If (alarm = &H24 Or alarm = &H34) Then                                                          '����
      Shape2.FillColor = QBColor(12)
     Else
      Shape2.FillColor = QBColor(7)
     End If
   End If
End Sub
                


Private Sub Form_Load()                      '��ʼ��
      With MSComm1
       MSComm1.CommPort = 1                         'ʹ��COM1
       MSComm1.Settings = "9600,N,8,1"               '����ͨ�ſڲ���
       MSComm1.InBufferSize = 512                    '����MSComm1���ջ�����Ϊ512�ֽ�
       MSComm1.OutBufferSize = 512                    '����MSComm1���ͻ�����Ϊ512�ֽ�
       MSComm1.InputMode = comInputModeText       '���ý�������ģʽΪ�ı���ʽ
       MSComm1.InputLen = 0                        '����Input һ�δӽ��ջ����ȡ�ֽ���Ϊȫ��
       MSComm1.SThreshold = 1                       '����Output һ�δӷ��ͻ����ȡ�ֽ���Ϊ1
       MSComm1.RThreshold = 1               '���ý���һ���ֽڲ���OnComm�¼�
      If MSComm1.PortOpen = False Then     '�ж�ͨ�ſ��Ƿ��
       MSComm1.PortOpen = True              '��ͨ�ſ�
          If Err Then                '������
            MsgBox "����ͨ����Ч"
            Exit Sub
          End If
       End If
    End With
Shape1.FillColor = QBColor(7)
Shape2.FillColor = QBColor(7)
Shape4.FillColor = QBColor(7)
Shape5.FillColor = QBColor(7)
Shape6.FillColor = QBColor(7)
Shape8.FillColor = QBColor(7)
 
End Sub

                                          '**************************************ֹͣ����*********************************************
Private Sub Stop_Click()
 Dim B_stop(6) As Byte             '���巢�͵�ֹͣ�����ݰ�
 Dim S_data(6) As String          '���͵�����
 Dim i As Integer
  S_data(0) = &HFA
  S_data(1) = &H1               'ȥ��������
  S_data(2) = &H53
  S_data(3) = &H0
  S_data(4) = S_data(1) Xor S_data(2) Xor S_data(3)
  S_data(5) = &HFB
 For i = 0 To 5
   B_stop(i) = Val(S_data(i))
 Next i
 MSComm1.Output = B_stop
Shape6.FillColor = QBColor(7)
Shape4.FillColor = QBColor(7)
Shape5.FillColor = QBColor(7)
Shape8.FillColor = QBColor(7)
Shape1.FillColor = QBColor(7)
Shape2.FillColor = QBColor(7)
   MSComm1.InBufferCount = 0
   MSComm1.OutBufferCount = 0   '��շ��ͻ�����

End Sub

Private Sub BRK_Click()
 Dim Break(6) As Byte              '���巢�͵��ƶ������ݰ�
 Dim S_data(6) As String          '���͵�����
 Dim i As Integer
  S_data(0) = &HFA
  S_data(1) = &H1                'ȥ��������
  S_data(2) = &H53
  S_data(3) = &H1
  S_data(4) = S_data(1) Xor S_data(2) Xor S_data(3)
  S_data(5) = &HFB
 For i = 0 To 5
  Break(i) = Val(S_data(i))
 Next i
 MSComm1.Output = Break
 Shape4.FillColor = QBColor(12)
 Shape6.FillColor = QBColor(7)
 Shape5.FillColor = QBColor(7)
 Shape8.FillColor = QBColor(7)
 Shape1.FillColor = QBColor(7)
 Shape2.FillColor = QBColor(7)
  MSComm1.InBufferCount = 0
  MSComm1.OutBufferCount = 0   '��շ��ͻ�����

End Sub
                               '**************************************������*********************************************
Private Sub Run_Click()
 Dim Q_Run(6) As Byte              '���巢�͵����������ݰ�
 Dim S_data(6) As String          '���͵�����
 Dim i As Integer
  S_data(0) = &HFA
  S_data(1) = &H1               'ȥ��������
  S_data(2) = &H52
  S_data(3) = &H0
  S_data(4) = S_data(1) Xor S_data(2) Xor S_data(3)
  S_data(5) = &HFB
  For i = 0 To 5
     Q_Run(i) = Val(S_data(i))
  Next i
 MSComm1.Output = Q_Run
 
 Shape6.FillColor = QBColor(10)
 Shape4.FillColor = QBColor(7)
  MSComm1.OutBufferCount = 0   '��շ��ͻ�����
  MSComm1.InBufferCount = 0
 
 End Sub

                     '**************************************��ת����**********************Э����û��***********************
Private Sub FR_Click()
 Dim F_R(6) As Byte                 '���巢�͵ķ�ת�����ݰ�
 Dim S_data(6) As String          '���͵�����
 Dim i As Integer
  S_data(0) = &HFA
  S_data(1) = &H1                'ȥ��������
  S_data(2) = &H46
  S_data(3) = &H0
  S_data(4) = S_data(1) Xor S_data(2) Xor S_data(3)
  S_data(5) = &HFB
 For i = 0 To 5
  F_R(i) = Val(S_data(i))
 Next i
 MSComm1.Output = F_R
 
MSComm1.OutBufferCount = 0
MSComm1.InBufferCount = 0
End Sub

