VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataType 
   Caption         =   "Data Type Selection"
   ClientHeight    =   1320
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   2535
   OleObjectBlob   =   "frmDataType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim output As Integer
Dim source As String

Private Sub cmdCancel_Click()
    output = -1
    execute (output)
End Sub

Private Sub cmdDates_Click()
    output = 1
    execute (output)
End Sub

Private Sub cmdGeneral_Click()
    output = 2
    execute (output)
End Sub

Private Sub cmdText_Click()
    output = 3
    execute (output)
End Sub


Private Sub execute(runOption As Integer)
        Me.Hide
        'FilteredPasteValuesInPlace (runOption)
        setDataType (runOption)
        Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 200
    Me.Left = Application.Left + 50 'Application.Width - Me.Width
     
End Sub
 
