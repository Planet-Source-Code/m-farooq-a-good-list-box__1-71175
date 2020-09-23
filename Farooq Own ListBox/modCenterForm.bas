Attribute VB_Name = "modCenterForm"
'Centering the form
Public Function CForm(Frm As Form)
    Frm.Left = (Screen.Width / 2) - (Frm.Width / 2)
    Frm.Top = (Screen.Height / 2) - (Frm.Height / 2)
End Function

