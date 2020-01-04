Imports System.Data.OleDb
Public Class Form1
    Dim Conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim LokasiDB As String
    Sub Koneksi()
        LokasiDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=barang.accdb"
        Conn = New OleDbConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Koneksi()
        da = New OleDbDataAdapter("Select * from Data", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "Data")
        DataGridView1.DataSource = (ds.Tables("Data"))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            Dim CMD As OleDbCommand
            Call Koneksi()
            Dim simpan As String = "insert into Data values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "')"
            CMD = New OleDbCommand(simpan, Conn)
            CMD.ExecuteNonQuery()
            MsgBox("Input data berhasil")
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader
            CMD = New OleDbCommand("Select * From Data  where KodeBarang='" & TextBox1.Text & "'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Kode Barang Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("NamaBarang")
                TextBox3.Text = RD.Item("HargaBarang")
                TextBox4.Text = RD.Item("JumlahBarang")
                ComboBox1.Text = RD.Item("SatuanBarang")
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Koneksi()
        Dim CMD As OleDbCommand
        Dim edit As String = "update Data set NamaBarang='" & TextBox2.Text & "',HargaBarang='" & TextBox3.Text & "',JumlahBarang='" & TextBox4.Text & "',SatuanBarang='" & ComboBox1.Text & "' where KodeBarang='" & TextBox1.Text & "'"
        CMD = New OleDbCommand(edit, Conn)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil diUpdate")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus dengan Masukan NIM dan ENTER")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As OleDbCommand
                Dim hapus As String = "delete From Data  where KodeBarang='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Refresh()
    End Sub
End Class
