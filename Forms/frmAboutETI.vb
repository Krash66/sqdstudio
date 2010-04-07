Public Class frmAboutETI
    Inherits frmBlank

    Private Sub frmAboutETI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Label1.Text = "About ETI CDC Studio"
        Label2.Text = ""
        Label3.Text = MyBase.ProductName & " V3 ™  "
        Label4.Text = "Build " & MyBase.ProductVersion
        Label5.Text = "Copyright  © 2008. All rights reserved. Reproduced by Evolutionary Technologies International, Inc. under permission. Warning: This computer program is protected by copyright law and international treaties. Unauthorized reproduction or distribution of this program, or any portion of it, may result in severe civil and criminal penalties, and will be prosecuted to the maximum extent possible under law."

        cmdCancel.Visible = False
        cmdHelp.Visible = False

    End Sub

End Class
