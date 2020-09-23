VERSION 5.00
Object = "{BF18F2A4-8B30-11D3-A95C-00008639BD6E}#1.0#0"; "APToolkit.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " PDF Stamp Utility"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleMode       =   0  'User
   ScaleWidth      =   9075
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   8835
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2670
         TabIndex        =   15
         ToolTipText     =   "Use the ""input""  button to browse for an existing input file"
         Top             =   900
         Width           =   4485
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   7440
         TabIndex        =   9
         Top             =   4620
         Width           =   1095
      End
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   3195
         Left            =   -390
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   2070
         Width           =   2805
         _cx             =   5080
         _cy             =   5080
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Input"
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   870
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Use the ""Output"" button to browse for a folder"
         Top             =   2040
         Width           =   4485
      End
      Begin VB.CheckBox Check2 
         Caption         =   "View the file after it has been created"
         Height          =   255
         Left            =   2670
         TabIndex        =   5
         ToolTipText     =   "Acrobat Reader will be used to open the file if this option is enabled"
         Top             =   3600
         Value           =   1  'Checked
         Width           =   3105
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   4230
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Set a password to open the file"
         Height          =   285
         Left            =   2670
         TabIndex        =   6
         ToolTipText     =   "Use this option to set a password to open the pdf file"
         Top             =   3870
         Width           =   2955
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Output"
         Height          =   375
         Left            =   7440
         TabIndex        =   2
         Top             =   2010
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Use stationery for all the pages."
         Height          =   285
         Index           =   0
         Left            =   2670
         TabIndex        =   3
         Top             =   2490
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Use stationery on the first page only."
         Height          =   285
         Index           =   1
         Left            =   2700
         TabIndex        =   4
         Top             =   2760
         Width           =   3075
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   375
         Left            =   7440
         TabIndex        =   10
         Top             =   3630
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1620
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Caption         =   "Browse for a PDF input file"
         Height          =   255
         Left            =   2700
         TabIndex        =   16
         Top             =   630
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Input"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   450
         Left            =   630
         Picture         =   "Form1.frx":0442
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Kurt Koenig Belgium    03/2006"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   3750
         TabIndex        =   13
         Top             =   5040
         Width           =   2385
      End
      Begin PETOCXLib.PETOCX PETOCX1 
         Left            =   120
         Top             =   420
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Browse for a folder and enter a name for the PDF file to save:"
         Height          =   285
         Left            =   2700
         TabIndex        =   11
         Top             =   1740
         Width           =   4485
      End
      Begin VB.Image Image2 
         Height          =   1830
         Left            =   360
         Picture         =   "Form1.frx":05EC
         Top             =   150
         Width           =   1590
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' // You have to install the commercial ActivePDF toolkit object for this to work.
' // This program won't work at all without that. You can download a free
' // trial version from http://www.activepdf.com/
'
'
' // The installation of Adobe Reader 7 is also required for this program.
' // (the Acropdf.dll is used by the thumbnail window)
'
'
'// BTW: THE INPUT FILE YOU USE MUST HAVE A TRANSPARENT BACKGROUND
'// USE THE INPUT EXAMPLE IN THIS PROGRAM'S FOLDER IF YOU CAN'T FIND ONE

' ----------------------------------------------------------------------------------

'// To open the "saveas" dialog with "My Documents" as default folder

Private Const CSIDL_DOCUMENTS = 5

'// Next line is to open the generated PDF with the asociated program (Adobe Reader)

Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Dim strPDFName, fso, strPath, TK, R, VarExisting, VarInputPDF, VarLogo, MyFile, pgs, startMess, strPasswd, varBackground, MyDocFolder


Private Sub form_load()

'// put location of "My Documents" in a variable

MyDocFolder = fGetSpecialFolder(CSIDL_DOCUMENTS)
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") & "\"

'// the next three variables are for use with the ActivePDF Toolkit object

VarInputPDF = App.Path & "\input.pdf"

'// Below is the pdf file that is used as stamp, or watermark, for your own pdf files
'// You can replace this with your own company stationery, but you have to keep the
'// file name. Remember you need to aquire a license for the activepdf toolkit if
'// you are going to use this tool in real life.


VarLogo = ".\_stationery.pdf"
varBackground = "1" '1 = background, 0 = foreground

Text2.Enabled = False
Text2.BackColor = &H8000000F
Label3.Caption = "Input"


'// Acrobat Reader 7 MUST be installed to use this next object
'// else a "Run time Error" occurs

AcroPDF1.LoadFile VarInputPDF
    With AcroPDF1
    
        .setViewRect 0, 0, 0, 0
        .gotoFirstPage
        .setZoom 19 'or other percentage, the 19 is to fit my thumbnail window
        .setPageMode "OneSinglePage"
        .setShowToolbar False
        .setShowScrollbars False
        .setLayoutMode "pageonly"
        
    End With
    
End Sub


Private Sub Command3_Click()

    SelectSourcePDF
    
End Sub


Private Sub Preview()

'// this will open the output document in Acrobat Reader

    ShellExecute 0, vbNullString, strPDFName, vbNullString, vbNullString, vbMaximizedFocus

End Sub


Private Sub Command4_Click()

    CloseAll
    
End Sub


Private Sub Form_Terminate()

    CloseAll
    
End Sub



Private Sub Text1_Click()

    startMess = MsgBox(vbCrLf & "Please use the " & Chr(34) & "Output" & Chr(34) & " button first, then enter a name for the file in the " & Chr(34) & "Choose folder" & Chr(34) & vbCrLf & "dialog.", 48, "PDF Stamp Utility")

End Sub


Private Sub Text3_Click()

    startMess = MsgBox(vbCrLf & "Please use the " & Chr(34) & "Input" & Chr(34) & " button first, then enter a name for the file in the " & Chr(34) & "Choose folder" & Chr(34) & vbCrLf & "dialog.", 48, "PDF Stamp Utility")

End Sub


Private Sub Command1_Click()

    SelectTargetPDF

End Sub


Private Sub Command2_Click()

If VarInputPDF = "" Then
    startMess = MsgBox(vbCrLf & "You didn't choose a source PDF file.", 48, "Action Canceled")

Else

    If strPDFName = "" Then
    startMess = MsgBox(vbCrLf & "You didn't enter a target file name or the file path cannot be found." & vbCrLf & "Please use the " & Chr(34) & "Output" & Chr(34) & " button and enter a name for the file" & vbCrLf & "that will be created first!", 16, "PDF Stamp Utility")
        Exit Sub
    End If

    If Check1.Value Then
        strPasswd = Text2.Text
    Else
        strPasswd = ""
    End If

        If strPDFName <> "" Then
            If (fso.FileExists(strPDFName)) Then
                PDFexists
            Else
                KillProcess ("Acrord32.exe")
                SaveFile
                
            End If
        End If
 End If
End Sub


Private Sub Check1_Click()

    If Check1.Value Then
        Text2.BackColor = &HFFFFFF
        Text2.Enabled = True
        Text2.SetFocus
    Else
        Text2.Enabled = False
        Text2.BackColor = &H8000000F
        Text2.Text = ""
    End If

End Sub


Private Sub SelectTargetPDF()

On Error GoTo Handler
  CD1.DialogTitle = "Choose a folder and enter a file name to save your PDF document"
  CD1.InitDir = MyDocFolder
  CD1.Filter = "PDF Files *.pdf|*.pdf"
  CD1.CancelError = True
  CD1.ShowOpen
  
  strPDFName = CD1.FileName
  Text1.Text = strPDFName
  
  Exit Sub
  
Handler:
  startMess = MsgBox(vbCrLf & "You didn't enter a file name or pressed " & Chr(34) & "Cancel" & Chr(34), 48, "Action Canceled")
  strPDFName = ""
  Text1.Text = ""
  Exit Sub
  
End Sub


Private Sub SelectSourcePDF()

On Error GoTo Handler
  CD1.DialogTitle = "Choose a folder and a source PDF file"
  CD1.InitDir = MyDocFolder
  CD1.Filter = "PDF Files *.pdf|*.pdf"
  CD1.CancelError = True
  CD1.ShowOpen
  
  VarInputPDF = CD1.FileName
  Text3.Text = VarInputPDF
  
  '// again: Acrobat Reader 7 must be installed to use this next object
  
  AcroPDF1.LoadFile VarInputPDF
    With AcroPDF1
    .setViewRect 0, 0, 0, 0
        .gotoFirstPage
        .setZoom 19 'or other
        .setPageMode "OneSinglePage"
        .setShowToolbar False
        .setShowScrollbars False
        .setLayoutMode "pageonly"
    End With

  Exit Sub
  
Handler:
  startMess = MsgBox(vbCrLf & "You didn't choose a source PDF file or pressed " & Chr(34) & "Cancel" & Chr(34), 48, "Action Canceled")
  VarInputPDF = App.Path & "\input.pdf"
  Text3.Text = ""
  Exit Sub
  
End Sub


Private Sub SaveFile()

If Option1(0) Then


' // create a new instance of the activePDFtoolkit object

    Set TK = CreateObject("APToolkit.Object")
    
' // use 128 bit security, you can enter a password for the output file if you like

    TK.SetOutputSecurity128 strPasswd, "", True, False, False, False, False, False, False, True
    R = TK.OpenOutputFile(App.Path & "\temporary.pdf")
    R = TK.OpenInputFile(VarInputPDF)
' // this is for the pdf's properties window
    TK.SetInfo "stationery for e-papers", "Created by PDF Stamp Utility", "Kurt Koenig   http://kurtkoenig.homeunix.net", vbCrLf & "      PDF Stamp Utility created by Kurt Koenig Belgium using ActivePDF Toolkit Component and Visual Basic 6.0"
' // here the stationery pdf file is stamped to your input file (in the background)
    R = TK.AddLogo(App.Path & VarLogo, varBackground)
    R = TK.CopyForm(0, 0)
    R = TK.CloseOutputFile()
    Set TK = Nothing
    fso.MoveFile (App.Path & "\temporary.pdf"), strPDFName
' // the thumbnail view shows your output file
    AcroPDF1.LoadFile strPDFName
                AcroPDF1.gotoFirstPage
                    With AcroPDF1
                        .setViewRect 0, 0, 0, 0
                        .setZoom 19
                        .setPageMode "OneSinglePage"
                        .setShowToolbar False
                        .setShowScrollbars False
                        .setLayoutMode "pageonly"

                    End With
                      Label3.Caption = "Output"
    If Check2.Value Then
        Preview
    End If
  
    
    Exit Sub
Else

' // basically the same as the previous part, but now the stationery
' // is only used on the first page

    Set TK = CreateObject("APToolkit.Object")
    pgs = TK.NumPages(VarInputPDF)
    TK.SetOutputSecurity128 strPasswd, "", True, False, False, False, False, False, False, True
    R = TK.OpenOutputFile(App.Path & "\temporary.pdf")
    R = TK.OpenInputFile(VarInputPDF)
    TK.SetInfo "Integra stationery for e-papers and certificates", "Created by Integra PDF Maker", "Kurt Koenig   http://kurtkoenig.homeunix.net", vbCrLf & "      PDF Stamp Utility created for INTEGRA-bvba Statiestraat 164, Berchem" & vbCrLf & "      PO 2600 Belgium using ActivePDF Toolkit Component and Visual Basic 6.0"
    R = TK.AddLogo(App.Path & VarLogo, varBackground)
    R = TK.CopyForm(1, 1)
    R = TK.ClearLogosAndImages
    
    If pgs > 1 Then
           R = TK.CopyForm(2, 0)
        Else
            startMess = MsgBox(vbCrLf & "There is only one page in this PDF document.", 48, "ATTENTION: Single page!")
        End If
        
    R = TK.CloseOutputFile()
    Set TK = Nothing
    fso.MoveFile (App.Path & "\temporary.pdf"), strPDFName
    AcroPDF1.LoadFile strPDFName
                AcroPDF1.gotoFirstPage
                    With AcroPDF1
                        .setViewRect 0, 0, 0, 0
                        .setZoom 19
                        .setPageMode "OneSinglePage"
                        .setShowToolbar False
                        .setShowScrollbars False
                        .setLayoutMode "pageonly"

                    End With
                    Label3.Caption = "Output"
    If Check2.Value Then
        Preview
    End If

    
End If
  
End Sub


Function PDFexists()
  
    VarExisting = MsgBox(vbCrLf & "The file " & Chr(34) & strPDFName & Chr(34) & " already exists. Do you want to overwrite the existing file?", 68, "Existing filename!")

        If VarExisting = vbNo Then
        
            SelectTargetPDF
       
        Else
           
            Set MyFile = fso.GetFile(strPDFName)
            MyFile.Delete
                
            SaveFile
          
        End If
    
End Function


Private Sub CloseAll()

' // Remove the objects and also close AcroRd32.exe, otherwise it stays open in the back

  KillProcess ("Acrord32.exe")
  Set fso = Nothing
  Set TK = Nothing


End
  
End Sub


