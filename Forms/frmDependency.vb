Public Class frmDependency
    Inherits SQDStudio.frmBlank

    Dim objCopy As INode '//object copied in the clipboard
    Dim cTargetNode As TreeNode '//under this node object will be pasted


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents tvDep As System.Windows.Forms.TreeView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDependency))
        Me.tvDep = New System.Windows.Forms.TreeView
        Me.Label9 = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.ImageLocation = "C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstu" & _
            "dio\images\FormTop\sq_skyblue.jpg"
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Text = "Finish"
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        '
        'Label1
        '
        Me.Label1.Text = "Missing Dependencies"
        '
        'Label2
        '
        Me.Label2.Text = "This dialogbox will allow you to anyalyze missing dependencies on the target envi" & _
            "ronment. ""Red"" node indicates that it's missing on the target environment. "
        '
        'tvDep
        '
        Me.tvDep.Location = New System.Drawing.Point(16, 96)
        Me.tvDep.Name = "tvDep"
        Me.tvDep.Size = New System.Drawing.Size(480, 176)
        Me.tvDep.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(16, 75)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(320, 21)
        Me.Label9.TabIndex = 41
        Me.Label9.Text = "Missing dependencies on the target environment"
        '
        'frmDependency
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(514, 335)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.tvDep)
        Me.Name = "frmDependency"
        Me.Text = "SQData Studio V3"
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.tvDep, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.Label9, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    '/obj is object being pasted. and cNode is target node underwhich object is being pasted
    'newObjParent is new parent of object being pasted. 
    Public Function AddMissingDependency(ByVal obj As INode, ByVal cNode As TreeNode, ByVal newObjParent As INode) As Boolean

        objCopy = obj

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                AddMissingDependency = True
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
        End Select
    End Function

    '//obj = Object to be copied to the destination
    Function CreateMissingDependency(ByVal obj As INode, ByVal destNode As INode) As Boolean

        Select Case obj.Type
            Case NODE_SOURCEDATASTORE
                Dim o As clsDatastore = Nothing
                Dim sel As clsDSSelection

                For Each sel In o.ObjSelections

                Next
        End Select

        Return True
    End Function

    '//cNode is the node under which search has to be done
    '//obj is object which has to be searched. generally we will go by matching name
    'Function FindMatchingObject(ByVal cNode As INode, ByVal obj As INode) As INode

    'End Function

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            '/* Code to create selected missing dependency */

            DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try
    End Sub



    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        ShowHelp(modHelp.HHId.H_Add_an_Environment)
    End Sub
End Class
