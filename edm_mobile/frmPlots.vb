Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlServerCe
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class frmPlots
    Dim points_table As New DataTable
    Dim datum_table As New DataTable
    Dim units_table As New DataTable

    Dim minx As Double = 100000000000.0
    Dim maxx As Double = -100000000000.0
    Dim miny As Double = 100000000000.0
    Dim maxy As Double = -100000000000.0
    Dim minz As Double = 100000000000.0
    Dim maxz As Double = -100000000000.0

    Dim originx As Double = 0
    Dim originy As Double = 0
    Dim widthx As Double = 1
    Dim heighty As Double = 1

    Dim xaxis As String = "X"
    Dim yaxis As String = "Y"

    Dim pixel_size As Integer = 1
    Dim zoom As Integer = 0
    Dim clickx As Integer
    Dim clicky As Integer
    Dim show_grid As Boolean = True
    Dim show_info As Boolean = False

    Dim show_overlay As Boolean = False

    Private Sub frmPlots_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        XZ.Left = XY.Left
        XZ.Top = XY.Top
        YZ.Left = XY.Left
        YZ.Top = XY.Top

        load_points()
        range_points()
        center_on_last()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub load_points()

        Dim sql As String = "SELECT * FROM " + options(22)

        Using points As New SqlCeConnection("datasource='" + options(1) + "'")
            points.Open()
            Using dad As New SqlCeDataAdapter(sql, points)
                dad.Fill(points_table)
            End Using
            points.Close()
        End Using

        sql = "SELECT * FROM edm_datums"

        Using datums As New SqlCeConnection("datasource='" + options(1) + "'")
            datums.Open()
            Using dad As New SqlCeDataAdapter(sql, datums)
                dad.Fill(datum_table)
            End Using
            datums.Close()
        End Using

        sql = "SELECT * FROM edm_units"

        Using units As New SqlCeConnection("datasource='" + options(1) + "'")
            units.Open()
            Using dad As New SqlCeDataAdapter(sql, units)
                dad.Fill(units_table)
            End Using
            units.Close()
        End Using

    End Sub

    Private Sub range_points()

        Dim row As DataRow

        If points_table.Rows.Count > 0 Then

            For Each row In points_table.Rows
                If CDbl(row("X")) < minx Then minx = CDbl(row("X"))
                If CDbl(row("X")) > maxx Then maxx = CDbl(row("X"))
                If CDbl(row("Y")) < miny Then miny = CDbl(row("Y"))
                If CDbl(row("Y")) > maxy Then maxy = CDbl(row("Y"))
                If CDbl(row("Z")) < minz Then minz = CDbl(row("Z"))
                If CDbl(row("Z")) > maxz Then maxz = CDbl(row("Z"))
            Next

        End If

        ' If no data points, then try datums and units table
        If minx > maxx Then
            If datum_table.Rows.Count > 0 Then

                For Each row In datum_table.Rows
                    If CDbl(row("X")) < minx Then minx = CDbl(row("X"))
                    If CDbl(row("X")) > maxx Then maxx = CDbl(row("X"))
                    If CDbl(row("Y")) < miny Then miny = CDbl(row("Y"))
                    If CDbl(row("Y")) > maxy Then maxy = CDbl(row("Y"))
                    If CDbl(row("Z")) < minz Then minz = CDbl(row("Z"))
                    If CDbl(row("Z")) > maxz Then maxz = CDbl(row("Z"))
                Next

            End If

            ' If no datum table or points table, try units to set limits
            If units_table.Rows.Count > 0 Then

                For Each row In units_table.Rows
                    If Not row("minx").Equals(DBNull.Value) And Not row("miny").Equals(DBNull.Value) And Not row("maxx").Equals(DBNull.Value) And Not row("maxy").Equals(DBNull.Value) Then
                        If CDbl(row("minx")) <> CDbl(row("maxx")) And CDbl(row("miny")) <> CDbl(row("maxy")) Then
                            If CDbl(row("maxx")) - CDbl(row("minx")) <= 100 Then
                                If CDbl(row("minx")) < minx Then minx = CDbl(row("minx"))
                                If CDbl(row("maxx")) > maxx Then maxx = CDbl(row("maxx"))
                                If CDbl(row("miny")) < miny Then miny = CDbl(row("miny"))
                                If CDbl(row("maxy")) > maxy Then maxy = CDbl(row("maxy"))
                            End If
                        End If
                    End If
                Next

            End If
        End If

    End Sub

    Private Sub plot_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles plot.MouseDown

        If zoom <> 0 Then
            clickx = e.X
            clicky = plot.Height - e.Y
        End If

    End Sub

    Private Sub plot_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles plot.MouseUp

        If zoom <> 0 Then
            If System.Math.Abs(e.X - clickx) < 10 Or System.Math.Abs((plot.Height - e.Y) - clicky) < 10 Then
                If zoom = 1 And widthx >= 0.2 Then
                    originx = map_x(clickx) - widthx / 4
                    originy = map_y(clicky) - heighty / 4
                    widthx = widthx / 2
                    heighty = heighty / 2
                ElseIf zoom = -1 And widthx < 500 Then
                    originx = map_x(clickx) - widthx
                    originy = map_y(clicky) - heighty
                    heighty = heighty * 2
                    widthx = widthx * 2
                End If
            ElseIf zoom = 1 Then
                Dim new_scale As Integer
                new_scale = System.Math.Max(System.Math.Abs(e.X - clickx), System.Math.Abs((plot.Height - e.Y) - clicky))
                widthx = new_scale / plot.Width * widthx
                heighty = widthx
                originx = map_x(System.Math.Min(e.X, clickx))
                originy = map_y(System.Math.Min(clicky, plot.Height - e.Y))
            End If
            zoom = 0
            clickx = -1
            clicky = -1
            plot.Invalidate()

        ElseIf zoom = 0 And show_info Then
            Dim row As DataRow
            Dim closest_row As DataRow
            Dim closest_d As Integer = plot.Height
            Dim d As Integer
            Dim n As Integer = -1
            Dim x As Integer = -1
            Dim y As Integer = -1
            Dim closest_is_datum As Boolean = False

            For Each row In points_table.Rows
                d = System.Math.Abs(pos_x(CDbl(row(xaxis))) - e.X) + System.Math.Abs((plot.Height - pos_y(CDbl(row(yaxis)))) - e.Y)
                If d < closest_d Then
                    closest_d = d
                    closest_row = row
                End If
            Next

            For Each row In datum_table.Rows
                d = System.Math.Abs(pos_x(CDbl(row(xaxis))) - e.X) + System.Math.Abs((plot.Height - pos_y(CDbl(row(yaxis)))) - e.Y)
                If d < closest_d Then
                    closest_d = d
                    closest_row = row
                    closest_is_datum = True
                End If
            Next

            If closest_d < plot.Height * 0.1 Then
                If closest_is_datum Then
                    infobox.Text = closest_row("name").ToString
                Else
                    infobox.Text = closest_row("UNIT").ToString + "-" + closest_row("ID").ToString + "(" + closest_row("SUFFIX").ToString + ")"
                End If
                infobox.Width = CInt(infobox.Font.Size) * Len(infobox.Text)
                infobox.Text = infobox.Text + vbCrLf + FormatNumber(closest_row("X"), 3)
                infobox.Text = infobox.Text + vbCrLf + FormatNumber(closest_row("Y"), 3)
                infobox.Text = infobox.Text + vbCrLf + FormatNumber(closest_row("Z"), 3)
                infobox.Height = CInt(infobox.Font.Size) * 8
                infobox.Width = System.Math.Max(infobox.Width, CInt(infobox.Font.Size) * Len(FormatNumber(map_x(e.X), 3)))
            Else
                infobox.Text = FormatNumber(map_x(e.X), 3) + vbCrLf + FormatNumber(map_y(plot.Height - e.Y), 3)
                infobox.Height = CInt(infobox.Font.Size) * 4
                infobox.Width = CInt(infobox.Font.Size) * System.Math.Max(Len(FormatNumber(map_y(plot.Height - e.Y), 3)), Len(FormatNumber(map_x(e.X), 3)))
            End If
            If e.X < plot.Width / 2 Then infobox.Left = e.X Else infobox.Left = e.X - infobox.Width
            If e.Y > plot.Height / 2 Then infobox.Top = e.Y - infobox.Height Else infobox.Top = e.Y
            infobox.Visible = True
        End If

    End Sub

    Private Sub plot_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles plot.Paint

        Cursor.Current = Cursors.WaitCursor

        Dim row As DataRow

        Dim x As Double = -1
        Dim y As Double = -1

        Dim white As New Pen(Color.White)
        Dim red As New Pen(Color.Red)
        Dim yellow As New SolidBrush(Color.Yellow)
        Dim lightblue As New Pen(Color.Aqua)

        Dim point_size_x As Integer = System.Math.Max(scale_x(0.025), pixel_size)
        Dim point_size_y As Integer = System.Math.Max(scale_y(0.025), pixel_size)
        Dim half_point_size As Integer = CInt(System.Math.Max(point_size_x / 2, 1))

        'Plot the points themselves if there are any
        If points_table.Rows.Count > 0 Then
            Dim artifact As New List(Of System.Drawing.Point)
            Dim pointno As Integer = 0
            For Each row In points_table.Rows
                If CInt(row("SUFFIX")) = 0 Then
                    If artifact.Count = 1 Then
                        e.Graphics.DrawEllipse(white, artifact.Item(0).X - half_point_size, artifact.Item(0).Y - half_point_size, point_size_x, point_size_y)
                    ElseIf artifact.Count > 1 Then
                        e.Graphics.DrawLines(white, artifact.ToArray)
                    End If
                    artifact.Clear()
                End If
                artifact.Add(New Point(pos_x(CDbl(row(xaxis))), plot.Height - pos_y(CDbl(row(yaxis)))))
            Next
            If artifact.Count = 1 Then
                e.Graphics.DrawEllipse(red, artifact.Item(0).X - half_point_size, artifact.Item(0).Y - half_point_size, point_size_x, point_size_y)
            ElseIf artifact.Count > 1 Then
                e.Graphics.DrawLines(white, artifact.ToArray)
                e.Graphics.DrawEllipse(red, artifact.Item(artifact.Count - 1).X - half_point_size, artifact.Item(artifact.Count - 1).Y - half_point_size, point_size_x, point_size_y)
            End If
        End If

        'Plot the datums if there are any
        If datum_table.Rows.Count > 0 Then
            Dim p(2) As System.Drawing.Point
            Dim xplot As Integer
            Dim yplot As Integer

            For Each row In datum_table.Rows
                xplot = pos_x(CDbl(row(xaxis)))
                yplot = pos_y(CDbl(row(yaxis)))
                p(0).X = System.Math.Max(xplot - 5, 0) : p(0).Y = plot.Height - System.Math.Max(yplot - 5, 0)
                p(1).X = xplot : p(1).Y = plot.Height - System.Math.Min(yplot + 5, plot.Height)
                p(2).X = System.Math.Min(xplot + 5, plot.Width) : p(2).Y = plot.Height - System.Math.Max(yplot - 5, 0)
                e.Graphics.FillPolygon(yellow, p)
            Next

        End If

        'Add grid lines if asked for
        If show_grid Then
            Dim gridlines As New Pen(Color.Gray)
            gridlines.DashStyle = Drawing2D.DashStyle.Dash
            Dim grid_interval As Integer = System.Math.Floor(widthx / 10)
            If grid_interval < 1 Then grid_interval = 1
            For x = Int(originx) To Int(originx + widthx) Step grid_interval
                e.Graphics.DrawLine(gridlines, pos_x(x), 0, pos_x(x), plot.Height)
            Next
            For y = Int(originy) To Int(originy + heighty) Step grid_interval
                e.Graphics.DrawLine(gridlines, 0, plot.Height - pos_y(y), plot.Width, plot.Height - pos_y(y))
            Next
        End If

        'Add units
        If units_table.Rows.Count > 0 And xaxis = "X" And yaxis = "Y" Then
            Dim x1 As Integer
            Dim y1 As Integer
            Dim x2 As Integer
            Dim y2 As Integer

            For Each row In units_table.Rows
                If Not row("minx").Equals(DBNull.Value) And Not row("miny").Equals(DBNull.Value) And Not row("maxx").Equals(DBNull.Value) And Not row("maxy").Equals(DBNull.Value) Then
                    x1 = pos_x(CDbl(row("minx")))
                    y1 = plot.Height - pos_y(CDbl(row("miny")))
                    x2 = pos_x(CDbl(row("maxx")))
                    y2 = plot.Height - pos_y(CDbl(row("maxy")))
                    e.Graphics.DrawRectangle(lightblue, x1, y1, (x2 - x1) + 1, (y2 - y1) + 1)
                End If
            Next
        End If

        reset_menus()

        Cursor.Current = Cursors.Default

    End Sub

    Private Function pos_x(ByVal x As Double) As Integer

        Return (CInt(CSng((x - originx) / widthx) * CSng(plot.Width)))

    End Function

    Private Function scale_x(ByVal x As Double) As Integer

        Return (System.Math.Max(CInt(CSng(x / widthx) * CSng(plot.Width)), pixel_size))

    End Function

    Private Function pos_y(ByVal y As Double) As Integer

        Return (CInt(CSng((y - originy) / heighty) * CSng(plot.Height)))

    End Function

    Private Function scale_y(ByVal y As Double) As Integer

        Return (System.Math.Max(CInt(CSng(y / heighty) * CSng(plot.Height)), pixel_size))

    End Function

    Private Function map_x(ByVal x As Integer) As Double

        Return (CDbl(x) / CDbl(plot.Width) * widthx + originx)

    End Function

    Private Function map_y(ByVal y As Integer) As Double

        Return (CDbl(y) / CDbl(plot.Height) * heighty + originy)

    End Function

    Private Sub full_extent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles full_extent.Click

        Cursor.Current = Cursors.WaitCursor
        If xaxis = "X" And yaxis = "Y" Then
            originx = System.Math.Floor(minx)
            originy = System.Math.Floor(miny)
            widthx = System.Math.Max(System.Math.Floor(maxx - minx + 2), System.Math.Floor(maxy - miny + 2))
            heighty = widthx
        ElseIf xaxis = "Y" And yaxis = "Z" Then
            originx = System.Math.Floor(miny)
            originy = System.Math.Floor(minz)
            widthx = System.Math.Max(System.Math.Floor(maxy - miny + 2), System.Math.Floor(maxz - minz + 2))
            heighty = widthx
        ElseIf xaxis = "X" And yaxis = "Z" Then
            originx = System.Math.Floor(minx)
            originy = System.Math.Floor(minz)
            widthx = System.Math.Max(System.Math.Floor(maxx - minx + 2), System.Math.Floor(maxz - minz + 2))
            heighty = widthx
        End If
        zoom = 0
        reset_menus()
        plot.Invalidate()

    End Sub

    Private Sub zoom_in_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles zoom_in.Click

        If zoom = 1 Then zoom = 0 Else zoom = 1
        clickx = -1
        clicky = -1
        reset_menus()

    End Sub

    Private Sub zoom_out_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles zoom_out.Click

        If zoom = -1 Then zoom = 0 Else zoom = -1
        clickx = -1
        clicky = -1
        reset_menus()

    End Sub

    Private Sub reset_menus()

        If zoom = 1 Then zoom_in.BackColor = Color.Gray Else zoom_in.BackColor = Color.White
        If zoom = -1 Then zoom_out.BackColor = Color.Gray Else zoom_out.BackColor = Color.White
        If show_info Then info.BackColor = Color.Gray Else info.BackColor = Color.White
        If Not show_info Then infobox.Visible = False
        If show_grid Then grid.BackColor = Color.Gray Else grid.BackColor = Color.White

    End Sub

    Private Sub grid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grid.Click

        If show_grid Then show_grid = False Else show_grid = True
        plot.Invalidate()

    End Sub

    Private Sub info_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles info.Click

        If show_info Then show_info = False Else show_info = True
        plot.Invalidate()

    End Sub


    Private Sub XY_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles XY.Click

        XY.Visible = False
        XZ.Visible = True

        xaxis = "X"
        yaxis = "Z"

        center_on_last()
        plot.Invalidate()

    End Sub

    Private Sub XZ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles XZ.Click

        XZ.Visible = False
        YZ.Visible = True

        xaxis = "Y"
        yaxis = "Z"

        center_on_last()
        plot.Invalidate()

    End Sub

    Private Sub YZ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles YZ.Click

        YZ.Visible = False
        XY.Visible = True

        xaxis = "X"
        yaxis = "Y"

        center_on_last()
        plot.Invalidate()

    End Sub

    Private Sub center_on_last()

        If points_table.Rows.Count > 0 Then
            Dim row As DataRow = points_table.Rows(points_table.Rows.Count - 1)
            originx = CDbl(row(xaxis)) - 2
            originy = CDbl(row(yaxis)) - 2
            widthx = 4
            heighty = 4
        End If

    End Sub

    Private Sub MenuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuExit.Click

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        OpenFileDialog.Filter = "*.txt|*.txt|*.csv|*.csv"
        If OpenFileDialog.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If OpenFileDialog.FileName <> "" Then
                overlayfile = OpenFileDialog.FileName
                show_overlay = True
                plot.Invalidate()
            End If
        End If

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        If show_overlay Then
            show_overlay = False
        Else
            show_overlay = True
        End If
        plot.Invalidate()

    End Sub
End Class