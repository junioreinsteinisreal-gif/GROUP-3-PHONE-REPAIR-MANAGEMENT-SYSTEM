Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO

''' <summary>Central colours, fonts and control helpers for the whole app.</summary>
Public Module UITheme

    ' ---------- Palette ----------
    Public ReadOnly Navy As Color = Color.FromArgb(10, 25, 60)
    Public ReadOnly RoyalBlue As Color = Color.FromArgb(37, 83, 200)
    Public ReadOnly LightGrey As Color = Color.FromArgb(240, 242, 245)
    Public ReadOnly CardWhite As Color = Color.White
    Public ReadOnly SuccessGreen As Color = Color.FromArgb(34, 120, 60)
    Public ReadOnly WarningAmber As Color = Color.FromArgb(245, 166, 35)
    Public ReadOnly DangerRed As Color = Color.FromArgb(160, 20, 30)
    Public ReadOnly TextDark As Color = Color.FromArgb(40, 44, 52)
    Public ReadOnly TextMuted As Color = Color.FromArgb(100, 116, 139)

    ' ---------- Fonts ----------
    ' Change FontScale to make the whole app bigger or smaller (1.0 = original).
    Public Const FontScale As Single = 1.2F

    Public Function FontRegular(Optional size As Single = 10) As Font
        Return New Font("Segoe UI Semibold", size * FontScale, FontStyle.Regular)
    End Function

    Public Function FontBold(Optional size As Single = 10) As Font
        Return New Font("Segoe UI", size * FontScale, FontStyle.Bold)
    End Function

    ' ---------- Control helpers ----------
    Public Function MakeLabel(text As String, size As Single, bold As Boolean,
                              textColor As Color) As Label
        Dim lbl As New Label()
        lbl.Text = text
        lbl.AutoSize = True
        lbl.BackColor = Color.Transparent
        lbl.ForeColor = textColor
        lbl.Font = If(bold, FontBold(size), FontRegular(size))
        Return lbl
    End Function

    Public Function MakeCenteredLabel(text As String, size As Single, bold As Boolean,
                                      textColor As Color, width As Integer,
                                      height As Integer) As Label
        Dim lbl As Label = MakeLabel(text, size, bold, textColor)
        lbl.AutoSize = False
        lbl.Size = New Size(width, height)
        lbl.TextAlign = ContentAlignment.MiddleCenter
        Return lbl
    End Function

    Public Function MakeTextBox(width As Integer, Optional size As Single = 12) As TextBox
        Dim txt As New TextBox()
        txt.Font = FontRegular(size)
        txt.Width = width
        txt.BorderStyle = BorderStyle.FixedSingle
        txt.ForeColor = TextDark
        Return txt
    End Function

    Public Function MakeButton(text As String, backColor As Color,
                               width As Integer, height As Integer) As Button
        Dim btn As New Button()
        btn.Text = text
        btn.Size = New Size(width, height)
        btn.BackColor = backColor
        btn.ForeColor = Color.White
        btn.Font = FontBold(11)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.FlatAppearance.MouseOverBackColor = ControlPaint.Dark(backColor, 0.1F)
        btn.Cursor = Cursors.Hand
        Return btn
    End Function

    Public Function MakeCard(width As Integer, height As Integer) As Panel
        Dim pnl As New Panel()
        pnl.Size = New Size(width, height)
        pnl.BackColor = CardWhite
        Return pnl
    End Function

    ' ---------- Centering ----------
    Public Sub CenterIn(ctrl As Control, container As Control)
        ctrl.Left = Math.Max(0, (container.ClientSize.Width - ctrl.Width) \ 2)
        ctrl.Top = Math.Max(0, (container.ClientSize.Height - ctrl.Height) \ 2)
    End Sub

    ''' <summary>Keeps a control centred in a form, even when it is resized.</summary>
    Public Sub KeepCentered(frm As Form, ctrl As Control)
        AddHandler frm.Resize, Sub(s, ev) CenterIn(ctrl, frm)
        AddHandler frm.Shown, Sub(s, ev) CenterIn(ctrl, frm)
        CenterIn(ctrl, frm)
    End Sub

    ' ---------- Backdrop (navy gradient + optional faded photo) ----------
    Private backdropImage As Image = Nothing
    Private backdropLoaded As Boolean = False

    Private Function GetBackdropImage() As Image
        If backdropLoaded Then Return backdropImage
        backdropLoaded = True
        Try
            Dim folder As String = Path.GetDirectoryName(DatabaseManager.DbPath)
            Dim imgPath As String = Path.Combine(folder, "background.jpg")
            If File.Exists(imgPath) Then
                Using fs As New FileStream(imgPath, FileMode.Open, FileAccess.Read)
                    Using tmp As Image = Image.FromStream(fs)
                        backdropImage = New Bitmap(tmp)
                    End Using
                End Using
            End If
        Catch
            backdropImage = Nothing
        End Try
        Return backdropImage
    End Function

    ''' <summary>Paints the login-style background. Call after Me.DoubleBuffered = True.</summary>
    Public Sub ApplyBackdrop(frm As Form)
        frm.BackColor = Navy
        AddHandler frm.Paint,
            Sub(s, ev)
                Dim g As Graphics = ev.Graphics
                Dim r As Rectangle = frm.ClientRectangle
                If r.Width < 2 OrElse r.Height < 2 Then Return

                Using br As New LinearGradientBrush(r, Navy, RoyalBlue, 45.0F)
                    g.FillRectangle(br, r)
                End Using

                Dim img As Image = GetBackdropImage()
                If img IsNot Nothing Then
                    Dim cm As New ColorMatrix()
                    cm.Matrix33 = 0.3F   ' 30% opacity
                    Using attr As New ImageAttributes()
                        attr.SetColorMatrix(cm, ColorMatrixFlag.Default, ColorAdjustType.Bitmap)
                        g.DrawImage(img, r, 0, 0, img.Width, img.Height,
                                    GraphicsUnit.Pixel, attr)
                    End Using
                End If

                ' Transparent navy overlay keeps text readable
                Using ov As New SolidBrush(Color.FromArgb(110, Navy))
                    g.FillRectangle(ov, r)
                End Using
            End Sub
        AddHandler frm.Resize, Sub(s, ev) frm.Invalidate()
    End Sub

End Module