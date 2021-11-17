VERSION 5.00
Begin VB.Form frmQuestSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elegir Quest"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmQuestSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub PonerListaQuest(ByVal rData As String)

List1.Clear

Dim j As Integer, k As Integer
For j = 0 To List1.ListCount - 1
    Me.List1.RemoveItem 0
Next j
k = CInt(ReadField(1, rData, 44))

For j = 1 To k
    List1.AddItem ReadField(1 + j, rData, 44)
Next j

Me.Show , frmMain

End Sub
Private Sub List1_Click()
    Call SendData("INFD" & List1.ListIndex + 1)
End Sub
