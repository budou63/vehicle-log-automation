VERSION 5.00
Begin VB.UserForm frmVehicleReview
   Caption         =   "車両選択候補の確認"
   ClientHeight    =   6900
   ClientWidth     =   9000
   ShowModal       =   0   'False
   StartUpPosition =   0  'Manual
   Begin VB.Label lblProgress
      Caption         =   "0件"
      Height          =   300
      Left            =   240
      Top             =   180
      Width           =   8400
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   300
      Left            =   240
      Top             =   480
      Width           =   8400
   End
   Begin VB.Label lblDetails
      Caption         =   ""
      Height          =   4800
      Left            =   240
      Top             =   840
      Width           =   8400
      WordWrap        =   -1  'True
   End
   Begin VB.CommandButton cmdGoToRow
      Caption         =   "対象行へ移動"
      Height          =   420
      Left            =   240
      Top             =   5880
      Width           =   1440
   End
   Begin VB.CommandButton cmdApply
      Caption         =   "候補車両へ修正"
      Height          =   420
      Left            =   1800
      Top             =   5880
      Width           =   1800
   End
   Begin VB.CommandButton cmdReject
      Caption         =   "修正しない"
      Height          =   420
      Left            =   3720
      Top             =   5880
      Width           =   1440
   End
   Begin VB.CommandButton cmdDefer
      Caption         =   "保留・次の候補"
      Height          =   420
      Left            =   5280
      Top             =   5880
      Width           =   1680
   End
   Begin VB.CommandButton cmdFinish
      Caption         =   "確認終了"
      Height          =   420
      Left            =   7080
      Top             =   5880
      Width           =   1560
   End
End
Attribute VB_Name = "frmVehicleReview"
Attribute VB_PredeclaredId = True
Option Explicit

Private candidates As Collection
Private currentIndex As Long
Private closing As Boolean

Public Sub 候補を読み込む(ByVal sourceCandidates As Collection)
    Set candidates = sourceCandidates
    currentIndex = 1
    Me.Caption = "車両選択候補の確認"
    Me.Top = 80
    Me.Left = 80
    Call 現在候補を表示
End Sub

Private Sub 現在候補を表示()
    Dim item As Object
    If candidates Is Nothing Or currentIndex > candidates.Count Then
        lblProgress.Caption = "確認対象はありません"
        lblStatus.Caption = "確認が完了しました。"
        lblDetails.Caption = ""
        cmdGoToRow.Enabled = False
        cmdApply.Enabled = False
        cmdReject.Enabled = False
        cmdDefer.Enabled = False
        Exit Sub
    End If
    Set item = candidates(currentIndex)
    lblProgress.Caption = CStr(currentIndex) & "件目 / " & CStr(candidates.Count) & "件"
    lblStatus.Caption = "記録シートを確認して判断してください。"
    cmdApply.Caption = CStr(item("候補車両")) & "号車へ修正"
    lblDetails.Caption = "車両選択ミスの可能性があります。" & vbCrLf & vbCrLf & _
                         "回答番号：" & item("回答番号") & vbCrLf & _
                         "運転日：" & item("運転日") & vbCrLf & _
                         "使用時間：" & item("使用時間") & vbCrLf & _
                         "運転者：" & item("運転者") & vbCrLf & _
                         "行先：" & item("行先") & vbCrLf & _
                         "用務：" & item("用務") & vbCrLf & vbCrLf & _
                         "登録車両：" & item("登録車両") & vbCrLf & _
                         "候補車両：" & item("候補車両") & vbCrLf & vbCrLf & _
                         "【" & item("登録車両") & "号車】" & vbCrLf & item("登録推移") & vbCrLf & vbCrLf & _
                         "【" & item("候補車両") & "号車】" & vbCrLf & item("候補推移") & vbCrLf & vbCrLf & _
                         item("候補車両") & "号車として扱う方がメーター推移と時間関係が自然です。"
End Sub

Private Sub cmdGoToRow_Click()
    If candidates Is Nothing Or currentIndex > candidates.Count Then Exit Sub
    Call 車両候補の対象行へ移動(CStr(candidates(currentIndex)("回答番号")))
End Sub

Private Sub cmdApply_Click()
    Dim item As Object
    If candidates Is Nothing Or currentIndex > candidates.Count Then Exit Sub
    Set item = candidates(currentIndex)
    If 車両候補を承認して修正(CStr(item("回答番号")), CStr(item("登録車両")), CStr(item("候補車両"))) Then
        lblStatus.Caption = "修正済み：最終確認終了時に1回だけ整理します。"
        currentIndex = currentIndex + 1
        Call 現在候補を表示
    Else
        lblStatus.Caption = "修正しませんでした。内容チェックをやり直してください。"
    End If
End Sub

Private Sub cmdReject_Click()
    currentIndex = currentIndex + 1
    Call 現在候補を表示
End Sub

Private Sub cmdDefer_Click()
    currentIndex = currentIndex + 1
    Call 現在候補を表示
End Sub

Private Sub cmdFinish_Click()
    closing = True
    Call 車両候補確認を終了
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not closing Then
        closing = True
        Call 車両候補確認を終了
    End If
End Sub
