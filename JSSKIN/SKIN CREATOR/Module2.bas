Attribute VB_Name = "Module1"

    Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Sub DO_skin(FRM As Form)
On Error Resume Next
With FRM
    .TOPLEFT.Move .ScaleLeft, .ScaleTop
    .TOPRIGHT.Move .ScaleWidth - .TOPRIGHT.Width, .ScaleTop
    .TOPMID.Move .TOPLEFT.Width, .ScaleTop, .ScaleWidth - .TOPLEFT.Width - .TOPRIGHT.Width, .TOPLEFT.Height
    .LEFTTOP.Move .ScaleLeft, .ScaleTop + .TOPLEFT.Height
    .LEFTBOT.Move .ScaleLeft, .ScaleHeight - .LEFTBOT.Height
    .LEFTMID.Move .ScaleLeft, .ScaleTop + .TOPLEFT.Height + .LEFTTOP.Height, .LEFTTOP.Width, .ScaleHeight - .TOPLEFT.Height - .LEFTTOP.Height - .LEFTBOT.Height
    .RIGHTMID.Width = .RIGHTTOP.Width
    .RIGHTTOP.Move .ScaleWidth - .RIGHTTOP.Width, .ScaleTop + .TOPLEFT.Height
    .RIGHTBOT.Move .ScaleWidth - .RIGHTBOT.Width, .ScaleHeight - .RIGHTBOT.Height
    .RIGHTMID.Move .ScaleWidth - .RIGHTMID.Width, .ScaleTop + .TOPLEFT.Height + .RIGHTTOP.Height, .RIGHTTOP.Width, .ScaleHeight - .TOPLEFT.Height - .RIGHTTOP.Height - .RIGHTBOT.Height
    .BOT.Move .LEFTBOT.Width, .ScaleHeight - .BOT.Height, .ScaleWidth - .LEFTBOT.Width - .RIGHTBOT.Width, .LEFTBOT.Height
    .CLOSEBOX.Move .ScaleWidth - .CLOSEBOX.Width - CInt(FRMMAIN.TXTRIGHT), CInt(FRMMAIN.TXTTOP)
    .MAXRESBOX.Move .CLOSEBOX.Left - CInt(FRMMAIN.TXTGAP) - .MAXRESBOX.Width, CInt(FRMMAIN.TXTTOP)
    .MINBOX.Move .MAXRESBOX.Left - CInt(FRMMAIN.TXTGAP) - .MINBOX.Width, CInt(FRMMAIN.TXTTOP)
    .ONTOPBOX.Move .MINBOX.Left - CInt(FRMMAIN.TXTGAP) - .ONTOPBOX.Width, CInt(FRMMAIN.TXTTOP)
    .Label1.Move CInt(FRMMAIN.TXTCLEFT), CInt(FRMMAIN.TXTCTOP)
End With
End Sub

