Imports System.Data.OleDb
Public Class Form1
    Dim Conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim LokasiDB As String
    Sub Koneksi()
        LokasiDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb"
        Conn = New OleDbConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub

    Sub loaddata()
        da = New OleDbDataAdapter("Select * from barang", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "barang")
        DataGridView1.DataSource = (ds.Tables("barang"))
    End Sub

    Sub kosong()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Koneksi()
        loaddata()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
            kosong()

        Else
            Dim dan As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menambah data ini?", "pesan", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If dan = DialogResult.No Then

            ElseIf dan = DialogResult.Yes Then
                Dim CMD As OleDbCommand
                Koneksi()
                Dim simpan As String = "insert into barang values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "')"
                CMD = New OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
                MsgBox("Input data berhasil")
                kosong()
                loaddata()
            End If


        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value
        TextBox2.Text = DataGridView1.CurrentRow.Cells(1).Value
        TextBox3.Text = DataGridView1.CurrentRow.Cells(2).Value
        TextBox4.Text = DataGridView1.CurrentRow.Cells(3).Value
        ComboBox1.Text = DataGridView1.CurrentRow.Cells(4).Value
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus dengan Masukan Kode Barang dan ENTER")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As OleDbCommand
                Dim hapus As String = "delete from barang where KodeBarang='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                loaddata()
                kosong()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dan As DialogResult = MessageBox.Show("Apakah Anda yakin ingin mengubah data ini?", "pesan", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If dan = DialogResult.No Then

        ElseIf dan = DialogResult.Yes Then
            Call Koneksi()
            Dim CMD As OleDbCommand
            Dim edit As String = "update barang set NamaBarang='" & TextBox2.Text & "',HargaBarang='" & TextBox3.Text & "',JumlahBarang='" & TextBox4.Text & "',SatuanBarang='" & ComboBox1.Text & "' where KodeBarang='" & TextBox1.Text & "'"
            CMD = New OleDbCommand(edit, Conn)
            CMD.ExecuteNonQuery()
            MsgBox("Data Berhasil diUpdate")
            kosong()
            loaddata()

        End If


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        kosong()
    End Sub
End Class
