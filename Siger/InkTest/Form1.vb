Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.IO
Imports Siger
Imports Microsoft.Ink

Public Class Form1
    Inherits System.Windows.Forms.Form

    Private propertyGrid As propertyGrid
    Private customReco As SigerRecognizer

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        InkPicture2.SetGestureStatus(Microsoft.Ink.ApplicationGesture.AllGestures, True)

        propertyGrid = New PropertyGrid
        With propertyGrid
            .CommandsVisibleIfAvailable = True
            .Location = PropertyGridPlaceHolder.Location
            .Size = PropertyGridPlaceHolder.Size
            .Anchor = PropertyGridPlaceHolder.Anchor
            .Visible = True
        End With

        PropertyGridPlaceHolder.Parent.Controls.Add(propertyGrid)
        propertyGrid.BringToFront()

        customReco = New SigerRecognizer
        customReco.RecognizerList.Add(New ScratchOut)
        customReco.RecognizerList.Add(New Delete)
        customReco.RecognizerList.Add(New LineBreak)
        customReco.RecognizerList.Add(New Lowercase)
        customReco.RecognizerList.Add(New HorizontalLine)
        customReco.RecognizerList.Add(New Transpose)
        customReco.RecognizerList.Add(New ChevronDown)
        customReco.RecognizerList.Add(New ChevronUp)
        customReco.RecognizerList.Add(New Tick)

    End Sub

#Region " Windows Form Designer generated code "

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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents ExitMenu As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ClearMenu As System.Windows.Forms.MenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents InkPicture1 As Microsoft.Ink.InkPicture
    Friend WithEvents Splitter3 As System.Windows.Forms.Splitter
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents InkPicture2 As Microsoft.Ink.InkPicture
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Splitter4 As System.Windows.Forms.Splitter
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents InkPicture3 As Microsoft.Ink.InkPicture
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents PropertyGridPlaceHolder As System.Windows.Forms.Label
    Friend WithEvents SaveMenu As System.Windows.Forms.MenuItem
    Friend WithEvents OpenMenu As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.SaveMenu = New System.Windows.Forms.MenuItem
        Me.OpenMenu = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.ExitMenu = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ClearMenu = New System.Windows.Forms.MenuItem
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.PropertyGridPlaceHolder = New System.Windows.Forms.Label
        Me.Splitter4 = New System.Windows.Forms.Splitter
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.InkPicture3 = New Microsoft.Ink.InkPicture
        Me.Label3 = New System.Windows.Forms.Label
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.InkPicture2 = New Microsoft.Ink.InkPicture
        Me.Splitter3 = New System.Windows.Forms.Splitter
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.InkPicture1 = New Microsoft.Ink.InkPicture
        Me.Panel1.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.SaveMenu, Me.OpenMenu, Me.MenuItem5, Me.ExitMenu})
        Me.MenuItem1.Text = "&File"
        '
        'SaveMenu
        '
        Me.SaveMenu.Index = 0
        Me.SaveMenu.Text = "Save Stroke &As..."
        '
        'OpenMenu
        '
        Me.OpenMenu.Index = 1
        Me.OpenMenu.Text = "&Open"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.Text = "-"
        '
        'ExitMenu
        '
        Me.ExitMenu.Index = 3
        Me.ExitMenu.Text = "E&xit"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ClearMenu})
        Me.MenuItem3.Text = "&Edit"
        '
        'ClearMenu
        '
        Me.ClearMenu.Index = 0
        Me.ClearMenu.Text = "&Clear"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Panel7)
        Me.Panel1.Controls.Add(Me.Splitter4)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.Splitter2)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel1.Location = New System.Drawing.Point(496, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 576)
        Me.Panel1.TabIndex = 0
        '
        'Panel7
        '
        Me.Panel7.Controls.Add(Me.PropertyGridPlaceHolder)
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel7.Location = New System.Drawing.Point(0, 315)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(200, 261)
        Me.Panel7.TabIndex = 4
        '
        'PropertyGridPlaceHolder
        '
        Me.PropertyGridPlaceHolder.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PropertyGridPlaceHolder.Location = New System.Drawing.Point(0, 0)
        Me.PropertyGridPlaceHolder.Name = "PropertyGridPlaceHolder"
        Me.PropertyGridPlaceHolder.Size = New System.Drawing.Size(200, 264)
        Me.PropertyGridPlaceHolder.TabIndex = 15
        Me.PropertyGridPlaceHolder.Text = "Property Grid Place Holder"
        '
        'Splitter4
        '
        Me.Splitter4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter4.Location = New System.Drawing.Point(0, 312)
        Me.Splitter4.Name = "Splitter4"
        Me.Splitter4.Size = New System.Drawing.Size(200, 3)
        Me.Splitter4.TabIndex = 3
        Me.Splitter4.TabStop = False
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Label4)
        Me.Panel4.Controls.Add(Me.TextBox1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(0, 163)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(200, 149)
        Me.Panel4.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(0, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 24)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Recognized as:"
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Location = New System.Drawing.Point(0, 24)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(200, 128)
        Me.TextBox1.TabIndex = 14
        Me.TextBox1.Text = ""
        '
        'Splitter2
        '
        Me.Splitter2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter2.Location = New System.Drawing.Point(0, 160)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(200, 3)
        Me.Splitter2.TabIndex = 1
        Me.Splitter2.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.InkPicture3)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(200, 160)
        Me.Panel3.TabIndex = 0
        '
        'InkPicture3
        '
        Me.InkPicture3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.InkPicture3.Location = New System.Drawing.Point(0, 40)
        Me.InkPicture3.MarginX = -2147483648
        Me.InkPicture3.MarginY = -2147483648
        Me.InkPicture3.Name = "InkPicture3"
        Me.InkPicture3.Size = New System.Drawing.Size(200, 120)
        Me.InkPicture3.TabIndex = 16
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(0, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 18)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Last Gesture"
        '
        'Splitter1
        '
        Me.Splitter1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Splitter1.Location = New System.Drawing.Point(493, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 576)
        Me.Splitter1.TabIndex = 1
        Me.Splitter1.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Panel6)
        Me.Panel2.Controls.Add(Me.Splitter3)
        Me.Panel2.Controls.Add(Me.Panel5)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(493, 576)
        Me.Panel2.TabIndex = 2
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.Label2)
        Me.Panel6.Controls.Add(Me.InkPicture2)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel6.Location = New System.Drawing.Point(0, 323)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(493, 253)
        Me.Panel6.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 24)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Built-in Recognizer"
        '
        'InkPicture2
        '
        Me.InkPicture2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.InkPicture2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.InkPicture2.CollectionMode = Microsoft.Ink.CollectionMode.GestureOnly
        Me.InkPicture2.Location = New System.Drawing.Point(0, 33)
        Me.InkPicture2.MarginX = -2147483648
        Me.InkPicture2.MarginY = -2147483648
        Me.InkPicture2.Name = "InkPicture2"
        Me.InkPicture2.Size = New System.Drawing.Size(488, 215)
        Me.InkPicture2.TabIndex = 5
        '
        'Splitter3
        '
        Me.Splitter3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Splitter3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter3.Location = New System.Drawing.Point(0, 320)
        Me.Splitter3.Name = "Splitter3"
        Me.Splitter3.Size = New System.Drawing.Size(493, 3)
        Me.Splitter3.TabIndex = 2
        Me.Splitter3.TabStop = False
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.Label1)
        Me.Panel5.Controls.Add(Me.InkPicture1)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel5.Location = New System.Drawing.Point(0, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(493, 320)
        Me.Panel5.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 23)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "siger Recognizer:"
        '
        'InkPicture1
        '
        Me.InkPicture1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.InkPicture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.InkPicture1.Location = New System.Drawing.Point(0, 32)
        Me.InkPicture1.MarginX = -2147483648
        Me.InkPicture1.MarginY = -2147483648
        Me.InkPicture1.Name = "InkPicture1"
        Me.InkPicture1.Size = New System.Drawing.Size(488, 288)
        Me.InkPicture1.TabIndex = 4
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(696, 576)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel1)
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.Text = "Simple Gesture Recognition (SiGeR) Test Bench"
        Me.Panel1.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub InkPicture1_Stroke(ByVal sender As System.Object, ByVal e As Microsoft.Ink.InkCollectorStrokeEventArgs) Handles InkPicture1.Stroke
        RecognizeStroke(InkPicture1)
    End Sub

    Private Sub RecognizeStroke(ByVal inkPict As InkPicture)
        Dim stroke As Stroke = inkPict.Ink.Strokes(0)
        Try
            ' Todo: throws an exception if stroke has less than 11 points.  must fix.
            Console.WriteLine(stroke.PacketCount)
            Dim si As New StrokeInfo(stroke)
            propertyGrid.SelectedObject = si

            TextBox1.Text = ""

            For Each gesture As CustomGesture In customReco.Recognize(stroke)
                TextBox1.Text &= gesture.Name & vbCrLf
            Next

            If TextBox1.Text = String.Empty Then
                TextBox1.Text = "No Gesture: " + stroke.PacketCount.ToString()
            End If

            CopyToLastStrokeBox(inkPict)
            ClearInkPicture(inkPict)

            'TextBox1.Text = "too small: " + stroke.PacketCount.ToString()
            'CopyToLastStrokeBox(inkPict)
            'ClearInkPicture(inkPict)

        Catch ex As Exception
            Dim down As Boolean = True
            Dim i As Integer = 1
            Dim first As Point = stroke.GetPoint(0)
            Dim second As Point = stroke.GetPoint(1)

            While i < stroke.PacketCount AndAlso down = True
                If (second.Y <= first.Y) Then
                    down = False
                End If
                i = i + 1
            End While

            If (down = True) Then
                TextBox1.Text = "comma"
            Else
                TextBox1.Text = "exception: " + stroke.PacketCount.ToString()
            End If

            CopyToLastStrokeBox(inkPict)
            ClearInkPicture(inkPict)
        End Try
    End Sub

    Private Sub CopyToLastStrokeBox(ByVal inkPict As InkPicture)
        Dim inkCopy As Ink = inkPict.Ink.Clone
        Dim strokeCopy As Stroke = inkCopy.Strokes(0)
        Dim rect As System.Drawing.Rectangle = strokeCopy.GetBoundingBox
        strokeCopy.Move(-rect.Left, -rect.Top)
        InkPicture3.InkEnabled = False
        InkPicture3.Ink = inkCopy
        InkPicture3.InkEnabled = True
    End Sub

    Private Sub InkPicture2_Gesture(ByVal sender As System.Object, ByVal e As Microsoft.Ink.InkCollectorGestureEventArgs) Handles InkPicture2.Gesture
        TextBox1.Text = ""
        propertyGrid.SelectedObject = Nothing
        CopyToLastStrokeBox(InkPicture2)
        For Each recognizedGesture As Microsoft.Ink.Gesture In e.Gestures
            TextBox1.Text &= recognizedGesture.Id.ToString() & " - " & _
            recognizedGesture.Confidence.ToString() & vbCrLf
        Next
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub

    Private Sub ExitMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitMenu.Click
        End
    End Sub

    Private Sub ClearMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearMenu.Click
        ClearInkPicture(InkPicture1)
        TextBox1.Text = ""
    End Sub

    Private Sub ClearInkPicture(ByVal inkPic As InkPicture)
        inkPic.Ink.DeleteStrokes()
        inkPic.Invalidate()
    End Sub

    Private Sub PropertyGridPlaceHolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub SaveMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveMenu.Click
        Dim fo As New SaveFileDialog
        If fo.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim b() As Byte = InkPicture3.Ink.Save(Microsoft.Ink.PersistenceFormat.InkSerializedFormat)
            Dim s As New FileStream(fo.FileName, FileMode.Create)
            s.Write(b, 0, b.Length)
            s.Close()
        End If

    End Sub

    Private Sub OpenMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenMenu.Click
        Dim fo As New OpenFileDialog
        If fo.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim fi As New FileInfo(fo.FileName)
            Dim b(fi.Length) As Byte
            Dim s As New FileStream(fo.FileName, FileMode.Open)
            s.Read(b, 0, fi.Length)
            s.Close()
            Dim ink As New Microsoft.Ink.Ink
            ink.Load(b)

            InkPicture1.InkEnabled = False
            InkPicture1.Ink = ink
            InkPicture1.InkEnabled = True

            RecognizeStroke(InkPicture1)
        End If
    End Sub
End Class
