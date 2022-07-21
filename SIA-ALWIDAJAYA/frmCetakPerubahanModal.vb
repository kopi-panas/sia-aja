﻿Imports System.Data.OleDb
Public Class frmCetakPerubahanModal

    Private Sub SetCboPeriode()
        Dim mTahun As String
        Dim mPeriode As String

        mTahun = Format(Now.Year)
        mPeriode = mTahun

        Try
            Query = "SELECT DISTINCT tblPerubahanModal.Periode, IIf(Right([Periode],2)= '" & "01" & "', '" & "Januari" & "',IIf(Right([Periode],2)='" & "02" & "','" & "Februari" & "',IIf(Right([Periode],2)='" & "03" & "','" & "Maret" & "',IIf(Right([Periode],2)='" & "04" & "','" & "April" & "',IIf(Right([Periode],2)='" & "05" & "','" & "Mei" & "',IIf(Right([Periode],2)='" & "06" & "','" & "Juni" & "',IIf(Right([Periode],2)='" & "07" & "','" & "Juli" & "',IIf(Right([Periode],2)='" & "08" & "','" & "Agustus" & "',IIf(Right([Periode],2)='" & "09" & "','" & "September" & "',IIf(Right([Periode],2)='" & "10" & "','" & "Oktober" & "',IIf(Right([Periode],2)='" & "11" & "','" & "November" & "',IIf(Right([Periode],2)='" & "12" & "','" & "Desember" & "','" & "Bulan tidak dikenal" & "')))))))))))) AS Bulan FROM(tblPerubahanModal) GROUP BY tblPerubahanModal.Periode, IIf(Right([Periode],2)='" & "01" & "','" & "Januari" & "',IIf(Right([Periode],2)='" & "02" & "','" & "Februari" & "',IIf(Right([Periode],2)='" & "03" & "','" & "Maret" & "',IIf(Right([Periode],2)='" & "04" & "','" & "April" & "',IIf(Right([Periode],2)='" & "05" & "','" & "Mei" & "',IIf(Right([Periode],2)='" & "06" & "','" & "Juni" & "',IIf(Right([Periode],2)='" & "07" & "','" & "Juli" & "',IIf(Right([Periode],2)='" & "08" & "','" & "Agustus" & "',IIf(Right([Periode],2)='" & "09" & "','" & "September" & "',IIf(Right([Periode],2)='" & "10" & "','" & "Oktober" & "',IIf(Right([Periode],2)='" & "11" & "','" & "November" & "',IIf(Right([Periode],2)='" & "12" & "','" & "Desember" & "','" & "Bulan tidak dikenal" & "')))))))))))) ORDER BY tblPerubahanModal.Periode DESC"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            cboPeriode.Items.Clear()
            For a = 0 To dsData.Tables(0).Rows.Count - 1
                With cboPeriode
                    .Items.Add(dsData.Tables(0).Rows(a).Item(0) & " : " & dsData.Tables(0).Rows(a).Item(1))
                End With
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub frmCetakPerubahanModal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            KoneksiKeAccess()
            SetCboPeriode()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdCetak_Click(sender As Object, e As EventArgs) Handles cmdCetak.Click
        If Len(cboPeriode.Text) = 0 Then
            MsgBox("Pilih periode yang akan dicetak", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Pesan")
            cboPeriode.Focus()
        Else
            Try
                frmRptPerubahanModal.CrystalReportViewer1.SelectionFormula = "{tblPerubahanModal.Periode} = '" & Mid(cboPeriode.Text, 1, 6) & "'"
                frmRptPerubahanModal.CrystalReportViewer1.Dock = DockStyle.Fill
                frmRptPerubahanModal.CrystalReportViewer1.RefreshReport()
                frmRptPerubahanModal.ShowDialog()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub cmdKeluar_Click(sender As Object, e As EventArgs) Handles cmdKeluar.Click
        Me.Close()
    End Sub
End Class