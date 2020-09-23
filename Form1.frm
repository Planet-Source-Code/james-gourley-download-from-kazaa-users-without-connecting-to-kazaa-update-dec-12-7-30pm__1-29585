VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form KaZaA 
   Caption         =   "My KaZaA"
   ClientHeight    =   6372
   ClientLeft      =   132
   ClientTop       =   708
   ClientWidth     =   9372
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6372
   ScaleWidth      =   9372
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock SearchSpecify 
      Left            =   48
      Top             =   240
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   528
      Top             =   3216
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   48
      Top             =   3216
      _ExtentX        =   804
      _ExtentY        =   804
      _Version        =   393216
   End
   Begin VB.ListBox SearchList 
      Height          =   3120
      ItemData        =   "Form1.frx":058A
      Left            =   0
      List            =   "Form1.frx":058C
      TabIndex        =   16
      Top             =   3168
      Visible         =   0   'False
      Width           =   1980
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4668
      Left            =   2016
      TabIndex        =   13
      Top             =   1632
      Width           =   7356
      ExtentX         =   12975
      ExtentY         =   8234
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scanning"
      ForeColor       =   &H00008000&
      Height          =   1548
      Left            =   2016
      TabIndex        =   2
      Top             =   48
      Width           =   7356
      Begin MSWinsockLib.Winsock Winsock1 
         Index           =   0
         Left            =   1584
         Top             =   1152
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1920
         Top             =   1152
      End
      Begin VB.Frame Frame3 
         Caption         =   "Scan Speed"
         Height          =   492
         Left            =   2112
         TabIndex        =   29
         Top             =   1008
         Width           =   1452
         Begin MSComctlLib.Slider Slider1 
            Height          =   252
            Left            =   96
            TabIndex        =   30
            Top             =   192
            Width           =   1308
            _ExtentX        =   2307
            _ExtentY        =   445
            _Version        =   393216
            Min             =   -1000
            Max             =   -1
            SelStart        =   -1
            TickStyle       =   3
            Value           =   -1
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search"
         ForeColor       =   &H00008000&
         Height          =   1548
         Left            =   3612
         TabIndex        =   17
         Top             =   0
         Width           =   3756
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   1728
            Top             =   528
         End
         Begin VB.CheckBox SmartSearch 
            BackColor       =   &H00FFC0C0&
            Caption         =   "SmartSearch"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   204
            Left            =   96
            MouseIcon       =   "Form1.frx":058E
            MousePointer    =   99  'Custom
            TabIndex        =   31
            Top             =   1276
            Width           =   1548
         End
         Begin VB.CheckBox SearchButton 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   204
            Left            =   2880
            MouseIcon       =   "Form1.frx":06E0
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   576
            Width           =   756
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H0000FF00&
            Caption         =   "Search On Find"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   156
            Left            =   96
            MouseIcon       =   "Form1.frx":0832
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   1054
            Width           =   1548
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H0000FF00&
            Caption         =   "Current User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   156
            Left            =   96
            MouseIcon       =   "Form1.frx":0984
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   824
            Width           =   1548
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0000FF00&
            Caption         =   "All Users Found"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   156
            Left            =   96
            MouseIcon       =   "Form1.frx":0AD6
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   604
            Width           =   1548
         End
         Begin VB.TextBox Search 
            Height          =   288
            Left            =   96
            TabIndex        =   18
            Top             =   192
            Width           =   3612
         End
         Begin VB.ListBox SmartSearchList 
            Height          =   240
            ItemData        =   "Form1.frx":0C28
            Left            =   1584
            List            =   "Form1.frx":0C2F
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Shape Shape11 
            BorderColor     =   &H00FF0000&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   252
            Left            =   48
            Shape           =   4  'Rounded Rectangle
            Top             =   1248
            Width           =   1644
         End
         Begin VB.Shape Shape10 
            BorderColor     =   &H0000C000&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   216
            Left            =   48
            Top             =   1016
            Width           =   1644
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H0000C000&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   214
            Left            =   48
            Shape           =   4  'Rounded Rectangle
            Top             =   796
            Width           =   1644
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H0000C000&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   214
            Left            =   48
            Shape           =   4  'Rounded Rectangle
            Top             =   566
            Width           =   1644
         End
         Begin VB.Image Image1 
            Height          =   348
            Index           =   3
            Left            =   2592
            Picture         =   "Form1.frx":0C44
            Stretch         =   -1  'True
            Top             =   480
            Width           =   252
         End
         Begin VB.Shape Shape8 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   2688
            Shape           =   4  'Rounded Rectangle
            Top             =   528
            Width           =   1020
         End
         Begin VB.Label ShowResults 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Results"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   348
            Left            =   2928
            MouseIcon       =   "Form1.frx":150E
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   912
            Width           =   780
         End
         Begin VB.Image Image1 
            Height          =   348
            Index           =   2
            Left            =   2592
            Picture         =   "Form1.frx":1660
            Stretch         =   -1  'True
            Top             =   912
            Width           =   252
         End
         Begin VB.Label Searching 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Not Currently Searching"
            Height          =   192
            Left            =   2028
            TabIndex        =   25
            Top             =   1296
            Width           =   1680
         End
         Begin VB.Shape Shape7 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   348
            Left            =   2688
            Shape           =   4  'Rounded Rectangle
            Top             =   912
            Width           =   1020
         End
      End
      Begin VB.TextBox EndGroup4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1824
         TabIndex        =   11
         Text            =   "255"
         Top             =   672
         Width           =   396
      End
      Begin VB.TextBox EndGroup3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   10
         Text            =   "255"
         Top             =   672
         Width           =   396
      End
      Begin VB.TextBox EndGroup2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1056
         TabIndex        =   9
         Text            =   "66"
         Top             =   672
         Width           =   396
      End
      Begin VB.TextBox EndGroup1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   672
         TabIndex        =   8
         Text            =   "24"
         Top             =   672
         Width           =   396
      End
      Begin VB.TextBox StartGroup4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1824
         TabIndex        =   7
         Text            =   "0"
         Top             =   288
         Width           =   396
      End
      Begin VB.TextBox StartGroup3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Text            =   "0"
         Top             =   288
         Width           =   396
      End
      Begin VB.TextBox StartGroup2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1056
         TabIndex        =   5
         Text            =   "66"
         Top             =   288
         Width           =   396
      End
      Begin VB.TextBox StartGroup1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   672
         TabIndex        =   4
         Text            =   "24"
         Top             =   288
         Width           =   396
      End
      Begin VB.Image Image1 
         Height          =   396
         Index           =   1
         Left            =   2256
         Picture         =   "Form1.frx":1F2A
         Stretch         =   -1  'True
         Top             =   624
         Width           =   300
      End
      Begin VB.Label Command2 
         BackStyle       =   0  'Transparent
         Caption         =   "Clear List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   204
         Left            =   2592
         MouseIcon       =   "Form1.frx":27F4
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   720
         Width           =   876
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   672
         Width           =   1116
      End
      Begin VB.Image Image1 
         Height          =   396
         Index           =   0
         Left            =   2256
         Picture         =   "Form1.frx":2946
         Stretch         =   -1  'True
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Command1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   204
         Left            =   2592
         MouseIcon       =   "Form1.frx":3210
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   336
         Width           =   876
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   2352
         Shape           =   4  'Rounded Rectangle
         Top             =   288
         Width           =   1164
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 KaZaA User(s) Found"
         Height          =   192
         Left            =   96
         TabIndex        =   15
         Top             =   1104
         Width           =   1668
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Currently Scanning"
         Height          =   192
         Left            =   96
         TabIndex        =   14
         Top             =   1296
         Width           =   1620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End At:"
         Height          =   204
         Left            =   96
         TabIndex        =   12
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start At:"
         Height          =   192
         Left            =   96
         TabIndex        =   3
         Top             =   336
         Width           =   540
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   444
         Left            =   48
         Top             =   1056
         Width           =   2028
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   48
         Shape           =   4  'Rounded Rectangle
         Top             =   288
         Width           =   2364
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   48
         Shape           =   4  'Rounded Rectangle
         Top             =   672
         Width           =   2316
      End
   End
   Begin VB.ListBox IPList 
      Height          =   2736
      ItemData        =   "Form1.frx":3362
      Left            =   0
      List            =   "Form1.frx":3364
      TabIndex        =   0
      Top             =   192
      Width           =   1980
   End
   Begin VB.ListBox HiddenList 
      Height          =   1008
      ItemData        =   "Form1.frx":3366
      Left            =   96
      List            =   "Form1.frx":3368
      TabIndex        =   24
      Top             =   3168
      Visible         =   0   'False
      Width           =   1884
   End
   Begin VB.ListBox SearchHiddenList 
      Height          =   1200
      ItemData        =   "Form1.frx":336A
      Left            =   144
      List            =   "Form1.frx":336C
      TabIndex        =   27
      Top             =   4416
      Visible         =   0   'False
      Width           =   1596
   End
   Begin VB.ListBox AddressList 
      Height          =   1008
      ItemData        =   "Form1.frx":336E
      Left            =   0
      List            =   "Form1.frx":3370
      TabIndex        =   26
      Top             =   4224
      Visible         =   0   'False
      Width           =   1980
   End
   Begin RichTextLib.RichTextBox Hidden 
      Height          =   876
      Index           =   0
      Left            =   0
      TabIndex        =   28
      Top             =   5424
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1545
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":3372
   End
   Begin RichTextLib.RichTextBox SaveBuffer 
      Height          =   1212
      Left            =   624
      TabIndex        =   33
      Top             =   4512
      Visible         =   0   'False
      Width           =   876
      _ExtentX        =   1545
      _ExtentY        =   2138
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":3454
   End
   Begin VB.Label SearchLabel 
      Caption         =   "Search Results:"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   204
      Left            =   0
      TabIndex        =   23
      Top             =   2976
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "KaZaA/Morpheus Users:"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1764
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save User List"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open User List"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
         Begin VB.Menu mnuUserList 
            Caption         =   "User List"
         End
         Begin VB.Menu mnuClearSearch 
            Caption         =   "Search Field"
         End
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "Advanced"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Begin VB.Menu mnuNewSearch 
            Caption         =   "New Search"
         End
         Begin VB.Menu Sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAllUsers 
            Caption         =   "All Users Found"
         End
         Begin VB.Menu mnuCurrent 
            Caption         =   "Current User"
         End
         Begin VB.Menu mnuSearchonfind 
            Caption         =   "Search on Find"
         End
         Begin VB.Menu Sep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSmartSearch 
            Caption         =   "SmartSearch"
         End
      End
      Begin VB.Menu mnuUserSearch 
         Caption         =   "User Search"
         Begin VB.Menu mnuSearchForUser 
            Caption         =   "Search For A User"
         End
      End
      Begin VB.Menu mnuIPSearch 
         Caption         =   "IP Search"
         Begin VB.Menu mnuFindIP 
            Caption         =   "Search for Certain IP"
         End
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpecifyIP 
         Caption         =   "Specify User IP"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About 'My KaZaA'"
      End
   End
End
Attribute VB_Name = "KaZaA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TimesAround, SearchIndex, i

Sub SearchForText(SearchText As String, OtherInfo)
On Error Resume Next
Form_Resize
If Option1.Value = True Or Option3.Value = True Then
SearchList.Clear
SearchHiddenList.Clear
SearchList.Clear
  'Create buffer of all new shared folders
  Counter = 0
    For i = Hidden.ubound To IPList.ListCount - 1
        Load Hidden(Hidden.ubound + 1) 'Load buffer for new shared folder
        Hidden(i).Text = Inet1.OpenURL("http://" & AddressList.List(i) & ":1214")  'add contents to buffer
        Searching = "Reading " & IPList.List(i)
        Counter = 0
Wait:
        Counter = Counter + 1
        If Counter > 5000 Then GoTo StopWaiting
        If Inet1.StillExecuting = True Then GoTo Wait 'wait until Inet is finished getting shared folder
StopWaiting:
    Next i
    SearchList.Clear
  'search the buffers
    For i = 0 To Hidden.ubound - 1
        Searching = "Searching " & IPList.List(i)
        If SmartSearch.Value = 1 Then
            For j = 0 To SmartSearchList.ListCount - 1
                If InStr(1, LCase(Hidden(i).Text), LCase(SmartSearchList.List(j)), vbTextCompare) Then
                    SearchList.AddItem IPList.List(i)
                    SearchHiddenList.AddItem AddressList.List(i)
                    Exit For
                End If
            Next j
        Else
            If InStr(1, LCase(Hidden(i).Text), LCase(SearchText), vbTextCompare) Then
                SearchList.AddItem IPList.List(i)
                SearchHiddenList.AddItem AddressList.List(i)
            End If
        End If
    Next i
    
    If Option1.Value = True Then
        MsgBox "Search Complete" & vbCrLf & SearchList.ListCount & " Result(s) found."
        SearchButton.Value = 0
        ShowResults_Click
        ShowResults_Click
    End If
    
ElseIf Option2.Value = True Then
'if searching current user
Dim Temp As String

    Temp = Inet1.OpenURL(WebBrowser1.LocationURL)
Wait2:
    If Inet1.StillExecuting = True Then GoTo Wait2 'wait until shared folder is loaded
    'search for text
    If InStr(1, LCase(Temp), LCase(SearchText), vbTextCompare) Then
        MsgBox "Search text found!"
    Else
        MsgBox "Search text not found."
    End If
End If
End Sub

Private Sub Command1_Click()
Select Case Command1.Caption
'check if it is already scanning
Case "Scan"
    Timer1.Enabled = True
    Command1.Caption = "Stop Scanning"
    Exit Sub

Case "Stop Scanning"
    Timer1.Enabled = False
    Command1.Caption = "Scan"
End Select
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.FontBold = True
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.FontBold = False
End Sub

Private Sub Command2_Click()
IPList.Clear
AddressList.Clear
SearchHiddenList.Clear
For i = 1 To Hidden.ubound
    Unload Hidden(i)
Next i
i = 0
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.FontBold = True
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.FontBold = False
End Sub

Private Sub Form_Load()
TimesAround = 0
WebBrowser1.Navigate2 "http://www.kazaa.com"
End Sub

Private Sub Form_Resize()
'resize everything to look nice on the screen
On Error Resume Next
If SearchList.Visible = True Then
IPList.Height = Me.Height / 2
SearchLabel.Top = IPList.Top + IPList.Height
SearchList.Top = SearchLabel.Top + SearchLabel.Height + 50
SearchList.Height = Me.Height - SearchList.Top - 680
Else
IPList.Height = Me.Height - IPList.Top - 650
End If
WebBrowser1.Height = Me.Height - WebBrowser1.Top - 720
WebBrowser1.Width = Me.Width - WebBrowser1.Left - 120
Frame1.Width = WebBrowser1.Width
Frame2.Width = Frame1.Width - Frame2.Left
Search.Width = Frame2.Width - (Search.Left * 2)
SearchButton.Left = Search.Left + Search.Width - SearchButton.Width - 100
Searching.Left = SearchButton.Left + SearchButton.Width - Searching.Width
Image1(3).Left = SearchButton.Left - 300
Shape8.Left = Image1(3).Left + 100
Image1(2).Left = Image1(3).Left
ShowResults.Left = Image1(2).Left + Image1(2).Width + 100
Shape7.Left = Image1(2).Left + 100
Frame3.Left = Frame2.Left - Frame3.Width
End Sub

Private Sub IPList_Click()
'display other person's shared folder
WebBrowser1.Navigate "http://" & AddressList.List(IPList.ListIndex) & ":1214"
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuAllUsers_Click()
Select Case Option1.Value
Case True
Option1.Value = False
Case False
Option1.Value = True
End Select
End Sub

Private Sub mnuClearSearch_Click()
SmartSearchList.Clear
SearchList.Clear
HiddenList.Clear
Search = ""
For i = 1 To Hidden.ubound
 Unload Hidden(i)
Next i
End Sub

Private Sub mnuCurrent_Click()
Select Case Option2.Value
Case True
Option2.Value = False
Case False
Option2.Value = True
End Select
End Sub

Private Sub mnuFindIP_Click()
'Search all found ip addresses
Temp = InputBox("Search for IP:", "IP Search")
For i = 0 To AddressList.ListCount - 1
    If AddressList.List(i) = Temp Then
        IPList.ListIndex = i
        IPList_Click
        Exit Sub
    End If
Next i
MsgBox "IP Address not yet found." & vbCrLf & "Try using the 'Specify IP' command in the Advanced Menu.", vbInformation + vbOKOnly, "User not found"
End Sub

Private Sub mnuNewSearch_Click()
'open new search dialog
Search = InputBox("Search For:", "New Search")
SearchButton.Value = 1
SearchForText Search, ""
End Sub

Private Sub mnuOpen_Click()
Dim strtextline As String

CommonDialog1.CancelError = False
CommonDialog1.DialogTitle = "Open KaZaA User list"
CommonDialog1.Filter = "KaZaA User Lists (*.kul) | *.kul; | All Files (*.*) | *.*;"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then GoTo ErrorHappened
Open CommonDialog1.FileName For Input As #1
i = 0
        Do While Not EOF(1)
            Line Input #1, strtextline 'read file line by line
          'check header on line to sort data
            If Left(strtextline, 5) = "User:" Then IPList.AddItem (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "Addr:" Then AddressList.AddItem (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "Srt1:" Then StartGroup1.Text = (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "Srt2:" Then StartGroup2.Text = (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "Srt3:" Then StartGroup3.Text = (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "Srt4:" Then StartGroup4.Text = (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "End3:" Then EndGroup3.Text = (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "End4:" Then EndGroup4.Text = (Right(strtextline, Len(strtextline) - 5))
            If Left(strtextline, 5) = "Sped:" Then Slider1.Value = Val(Right(strtextline, Len(strtextline) - 6))
        Loop
    Close #1
ErrorHappened:
End Sub

Private Sub mnuQuit_Click()
End
End Sub

Private Sub mnuSave_Click()
On Error Resume Next
CommonDialog1.FileName = ""
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Save KaZaA User list"
CommonDialog1.Filter = "KaZaA User Lists (*.kul) | *.kul; | All Files (*.*) | *.*;"
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then GoTo ErrorHappened
'add all info to one list needed to save file
With SaveBuffer
    For i = 0 To IPList.ListCount - 1
        .Text = .Text & vbCrLf & "User:" & IPList.List(i)
        .Text = .Text & vbCrLf & "Addr:" & AddressList.List(i)
    Next i
    .Text = .Text & vbCrLf & "Srt1:" & StartGroup1
    .Text = .Text & vbCrLf & "Srt2:" & StartGroup2
    .Text = .Text & vbCrLf & "Srt3:" & StartGroup3
    .Text = .Text & vbCrLf & "Srt4:" & StartGroup4
    .Text = .Text & vbCrLf & "End3:" & EndGroup3
    .Text = .Text & vbCrLf & "End4:" & EndGroup4
    .Text = .Text & vbCrLf & "Sped:" & Slider1.Value & vbCrLf
End With
Kill CommonDialog1.FileName 'delete old file
Open CommonDialog1.FileName For Output As #1 'open file for writing
    Write #1, SaveBuffer.Text 'write to file
Close #1

ErrorHappened:
End Sub

Private Sub mnuSearchForUser_Click()
'search for user in userlist
Temp = InputBox("Search for:", "User Search")
For i = 0 To IPList.ListCount - 1
    If IPList.List(i) = Temp Then
        IPList.ListIndex = i
        IPList_Click
        Exit Sub
    End If
Next i
MsgBox "User not yet found.", vbInformation + vbOKOnly, "User not found"
End Sub

Private Sub mnuSearchonfind_Click()
Select Case Option3.Value
Case True
Option3.Value = False
Case False
Option3.Value = True
End Select
End Sub

Private Sub mnuSmartSearch_Click()
Select Case mnuSmartSearch.Checked
Case True
SmartSearch.Value = 0
Case False
SmartSearch.Value = 1
End Select
End Sub

Private Sub mnuSpecifyIP_Click()
On Error GoTo X
'connect to specific IP address
SearchSpecify.Close
SearchSpecify.Connect InputBox("Enter Remote Kazaa/Morpheus User's IP Address", "Specify IP Address"), 1214
Exit Sub
X:
MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub mnuUserList_Click()
IPList.Clear
AddressList.Clear
End Sub

Private Sub Search_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SearchButton_Click
    SearchButton.Value = 1
    KeyAscii = 0
End If
End Sub

Private Sub SearchButton_Click()
Select Case SearchButton.Value
Case 1
i = 0
    If Option2.Value = True Then
        Search = Replace(Search, " ", "+", 1, Len(Search), vbTextCompare)
        Search = Replace(Search, "*", "", 1, Len(Search), vbTextCompare)
        SearchForText Search, ""
        Exit Sub
    End If
    ShowResults.Caption = "Hide Results"
  'clear and show searchbox
    SearchIndex = 0
    'SearchList.Clear
    'HiddenList.Clear
    SearchList.Visible = True
    SearchLabel.Visible = True
    Search.Enabled = False
  If SmartSearch.Value = 1 Then
  'replace spaces and *'s with different characters (add more if you want)
    SmartSearchList.Clear
    SmartSearchList.AddItem Replace(Search, " ", "+", 1, Len(Search), vbTextCompare)
    SmartSearchList.AddItem Replace(Search, " ", "_", 1, Len(Search), vbTextCompare)
    SmartSearchList.AddItem Replace(Search, " ", ".", 1, Len(Search), vbTextCompare)
    SmartSearchList.AddItem Replace(Search, " ", "", 1, Len(Search), vbTextCompare)
    SmartSearchList.AddItem Replace(Search, "*", "+", 1, Len(Search), vbTextCompare)
    SmartSearchList.AddItem Replace(Search, "*", "_", 1, Len(Search), vbTextCompare)
    SmartSearchList.AddItem Replace(Search, "*", ".", 1, Len(Search), vbTextCompare)
    SmartSearchList.AddItem Replace(Search, "*", "", 1, Len(Search), vbTextCompare)
  Else
  'replace spaces with + signs to match internet explorer's spaces
  'remove all *'s
    Search = Replace(Search, " ", "+", 1, Len(Search), vbTextCompare)
    Search = Replace(Search, "*", "", 1, Len(Search), vbTextCompare)
  End If
    HiddenList.Clear
    SearchForText Search, ""
Case 0
    Searching = "Searching Cancelled"
  'hide searchbox
    SearchList.Visible = False
    SearchLabel.Visible = False
    Search.Enabled = True
End Select
Form_Resize 'resize listboxes to look nice
End Sub

Private Sub SearchList_Click()
On Error Resume Next
WebBrowser1.Navigate "http://" & SearchHiddenList.List(SearchList.ListIndex) & ":1214" 'go to selected shared folder
End Sub

Private Sub SearchSpecify_Connect()
'send string which will get user info
SearchSpecify.SendData "PASS Admin" & vbCrLf & "NICK M{iN}M" & vbCrLf & "USER KaZaAClone " & SearchSpecify.LocalIP & ":KaZaA"
End Sub

Private Sub SearchSpecify_DataArrival(ByVal bytesTotal As Long)
On Error GoTo X
Dim Data As String
SearchSpecify.GetData Data, vbString
'clean up data containing username
Data = Replace(Data, "HTTP/1.0 501 Not Implemented", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Username: ", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, SearchSpecify.RemoteHostIP, "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Network: KaZaA", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, " ", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-IP:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, ":1214", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, vbCrLf, "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-SupernodeIP:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Network:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "MusicCity", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, ".", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "0", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "1", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "2", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "3", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "4", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "5", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "6", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "7", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "8", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "9", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, Chr(10), "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, " ", "", 1, Len(Data), vbTextCompare)
'add data to listboxes
IPList.AddItem Data
AddressList.AddItem SearchSpecify.RemoteHostIP 'add ip to a hidden list

If SearchButton.Value = 1 And Option3.Value = True Then
    Load Hidden(Hidden.ubound + 1) 'Load buffer for new shared folder
        Hidden(Hidden.ubound).Text = Inet1.OpenURL("http://" & SearchSpecify.RemoteHostIP & ":1214")  'add contents to buffer
        Searching = "Reading " & IPList.List(IPList.ListCount - 1)
        Counter = 0
Wait:
        Counter = Counter + 1
        If Counter > 5000 Then GoTo StopWaiting
        If Inet1.StillExecuting = True Then GoTo Wait 'wait until Inet is finished getting shared folder
StopWaiting:
  'search the buffers
    If SmartSearch.Value = 1 Then
        For j = 0 To SmartSearchList.ListCount - 1
            If InStr(1, LCase(Hidden(Hidden.ubound).Text), LCase(SmartSearchList.List(j)), vbTextCompare) Then
                SearchList.AddItem IPList.List(IPList.ListCount - 1)
                SearchHiddenList.AddItem AddressList.List(AddressList.ListCount - 1)
                Exit For
            End If
        Next j
    Else
        If InStr(1, LCase(Hidden(Hidden.ubound).Text), LCase(Search), vbTextCompare) Then
            SearchList.AddItem IPList.List(IPList.ListCount - 1)
            SearchHiddenList.AddItem AddressList.List(AddressList.ListCount - 1)
        End If
    End If
        
    'HiddenList.AddItem Winsock1(Index).RemoteHostIP 'add ip to hidden list (acts as a buffer of ip addresses if you are finding hosts faster than you can scan them)
End If
X:
SearchSpecify.Close 'close connection (avoid's multiple listings of single ip)
IPList.ListIndex = IPList.ListCount - 1
IPList_Click
End Sub

Private Sub ShowResults_Click()
Select Case ShowResults.Caption
Case "Show Results"
    SearchList.Visible = True
    SearchLabel.Visible = True
    Form_Resize
    ShowResults.Caption = "Hide Results"
    Exit Sub
Case "Hide Results"
    SearchList.Visible = False
    SearchLabel.Visible = False
    Form_Resize
    ShowResults.Caption = "Show Results"
End Select
End Sub

Private Sub ShowResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowResults.FontBold = True
End Sub

Private Sub ShowResults_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowResults.FontBold = False
End Sub

Private Sub StartGroup1_Change()
EndGroup1 = StartGroup1
End Sub

Private Sub StartGroup2_Change()
EndGroup2 = StartGroup2
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

If Option1.Value = True Then
    mnuAllUsers.Checked = True
    mnuCurrent.Checked = False
    mnuSearchonfind.Checked = False
ElseIf Option2.Value = True Then
    mnuAllUsers.Checked = False
    mnuCurrent.Checked = True
    mnuSearchonfind.Checked = False
ElseIf Option3.Value = True Then
    mnuAllUsers.Checked = False
    mnuCurrent.Checked = False
    mnuSearchonfind.Checked = True
End If
If SmartSearch.Value = 1 Then mnuSmartSearch.Checked = True

Timer1.Interval = -Slider1.Value
Label5.Caption = IPList.ListCount & " KaZaA user(s) found." 'display number of kazaa users

TimesAround = TimesAround + 1

Load Winsock1(TimesAround) 'load new winsock
If TimesAround > 50 Then Unload Winsock1(TimesAround - 50) 'unload winsock control (time out)

If Val(StartGroup4) < Val(EndGroup4) Then
    StartGroup4.Text = StartGroup4.Text + 1 'increase current ip address by one
ElseIf Val(StartGroup4) = Val(EndGroup4) Then
    StartGroup3 = StartGroup3 + 1 'increase 3rd group in ip address by one
    StartGroup4 = 0 'reset last group in ip address
End If

If Val(StartGroup3) > Val(EndGroup3) Then 'check if scan is complete
    MsgBox "Scan Complete"
    Timer1.Enabled = False
    Command1.Caption = "Scan"
End If

Winsock1(TimesAround).Connect StartGroup1 & "." & StartGroup2 & "." & StartGroup3 & "." & StartGroup4, 1214 'connect to potential kazaa user
Label4.Caption = "Scanning " & StartGroup1 & "." & StartGroup2 & "." & StartGroup3 & "." & StartGroup4 'display current ip address
End Sub

Private Sub Timer2_Timer()
If Inet1.StillExecuting = False Then
    If Option1.Value = True And SearchButton.Value = 1 Then
        i = Hidden.ubound + 1
        SearchForText Search, ""
    End If
    If Option3.Value = True And SearchButton.Value = 1 Then
        i = Hidden.ubound + 1
        SearchForText Search, ""
    End If

    If IPList.ListCount > Hidden.ubound And Option3.Value = True Then SearchForText Search, ""
End If
Timer2.Enabled = False
End Sub

Private Sub Winsock1_Connect(Index As Integer)
Winsock1(Index).SendData "PASS Admin" & vbCrLf & "NICK M{iN}M" & vbCrLf & "USER KaZaAClone " & Winsock1(Index).LocalIP & ":KaZaA"
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo X
Dim Data As String
Winsock1(Index).GetData Data, vbString
'clean up data containing username
Data = Replace(Data, "HTTP/1.0 501 Not Implemented", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Username: ", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, Winsock1(Index).RemoteHostIP, "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Network: KaZaA", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, " ", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-IP:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, ":1214", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, vbCrLf, "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-SupernodeIP:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Network:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "MusicCity", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, ".", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "0", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "1", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "2", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "3", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "4", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "5", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "6", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "7", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "8", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "9", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, Chr(10), "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, " ", "", 1, Len(Data), vbTextCompare)
'add data to listboxes
IPList.AddItem Data
AddressList.AddItem Winsock1(Index).RemoteHostIP 'add ip to a hidden list

If SearchButton.Value = 1 And Option3.Value = True Then
    Load Hidden(Hidden.ubound + 1) 'Load buffer for new shared folder
        Hidden(Hidden.ubound).Text = Inet1.OpenURL("http://" & Winsock1(Index).RemoteHostIP & ":1214")  'add contents to buffer
        Searching = "Reading " & IPList.List(IPList.ListCount - 1)
        Counter = 0
Wait:
        Counter = Counter + 1
        If Counter > 5000 Then GoTo StopWaiting
        If Inet1.StillExecuting = True Then GoTo Wait 'wait until Inet is finished getting shared folder
StopWaiting:
  'search the buffers
    If SmartSearch.Value = 1 Then
        For j = 0 To SmartSearchList.ListCount - 1
            If InStr(1, LCase(Hidden(Hidden.ubound).Text), LCase(SmartSearchList.List(j)), vbTextCompare) Then
                SearchList.AddItem IPList.List(IPList.ListCount - 1)
                SearchHiddenList.AddItem AddressList.List(AddressList.ListCount - 1)
                Exit For
            End If
        Next j
    Else
        If InStr(1, LCase(Hidden(Hidden.ubound).Text), LCase(Search), vbTextCompare) Then
            SearchList.AddItem IPList.List(IPList.ListCount - 1)
            SearchHiddenList.AddItem AddressList.List(AddressList.ListCount - 1)
        End If
    End If
        
    'HiddenList.AddItem Winsock1(Index).RemoteHostIP 'add ip to hidden list (acts as a buffer of ip addresses if you are finding hosts faster than you can scan them)
End If
Winsock1(Index).Close 'close connection (avoid's multiple listings of single ip)
X:
End Sub
