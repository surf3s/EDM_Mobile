Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlServerCe
Imports System.Windows.Forms
Public Class frmDBGrids

    Private Sub frmDBGrids_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        Cursor.Current = Cursors.WaitCursor
        Dim db As New SqlCeConnection("datasource='" + options(1) + "'")
        db.Open()

        Dim sql As New SqlCeCommand("SELECT * FROM " + options(22), db)
        Dim dataadapter As SqlCeDataAdapter = New SqlCeDataAdapter(sql)
        Dim dspoints As New DataSet
        dataadapter.Fill(dspoints, options(22))
        bspoints.DataSource = dspoints
        bspoints.DataMember = options(22)
        pointgrid.DataSource = bspoints

        sql.CommandText = "SELECT * FROM edm_datums"
        Dim dsdatums As New DataSet
        dataadapter.Fill(dsdatums, "edm_datums")
        bsdatums.DataSource = dsdatums
        bsdatums.DataMember = "edm_datums"
        datumgrid.DataSource = bsdatums

        sql.CommandText = "SELECT * FROM edm_poles"
        Dim dsprisms As New DataSet
        dataadapter.Fill(dsprisms, "edm_poles")
        bsprisms.DataSource = dsprisms
        bsprisms.DataMember = "edm_poles"
        prismsgrid.DataSource = bsprisms

        sql.CommandText = "SELECT * FROM edm_units"
        Dim dsunits As New DataSet
        dataadapter.Fill(dsunits, "edm_units")
        bsunits.DataSource = dsunits
        bsunits.DataMember = "edm_units"
        unitsgrid.DataSource = bsunits

        db.Close()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub
End Class