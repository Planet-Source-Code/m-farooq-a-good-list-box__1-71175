Attribute VB_Name = "WriteOnPictureBox"
Public Sub PrintToCenter(Msg As String, PicBox As PictureBox)
        With PicBox
           .AutoRedraw = -1
           .Font = "Courier New"
           .FontSize = 10
           .FontBold = True
           .ForeColor = vbBlack
           
           HalfWidth = .TextWidth(Msg) / 2     ' Calculate one-half width.
           HalfHeight = .TextHeight(Msg) / 2   ' Calculate one-half height.
           .CurrentX = .ScaleWidth / 2 - HalfWidth   ' Set X.
           .CurrentY = .ScaleHeight / 2 - HalfHeight ' Set Y.
        End With
        
    PicBox.Print Msg   ' Print message.

End Sub


