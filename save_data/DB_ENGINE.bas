Attribute VB_Name = "db_engine"
' if add reference sub fail
' dont forget to add referece to dao (3.6) manual

Global error_var As String
Dim eng As Object
Dim db As Object
Dim rs As Object


Private Sub add_reference(ByVal db_name As String, Optional ByVal tbl_name As String)
Set eng = CreateObject("dao.DBEngine.36")
Set db = eng.OpenDatabase(db_name, 0, 0, ";pwd=5323")
If Len(Trim(tbl_name)) <> 0 Then

Set rs = db.OpenRecordset(tbl_name)
End If
End Sub






'Dim db As DAO.Database
'Dim rs As DAO.Recordset


Function SAVE_DATA(ByVal frm_name As Form, ByVal TBLNAME As String) As Boolean
On Error GoTo er

add_reference "schools.mdb", "students"


rs.AddNew

For Each Control In frm_name
If TypeOf Control Is TextBox Then
rs.Fields(Control.Name).Value = Control.Text
End If
Next


rs.Update

SAVE_DATA = True

close_data

'----------------------------------------
Exit Function
er:
error_var = Err.Description
SAVE_DATA = False
Exit Function
End Function







Sub create_textbox(ByVal frm_name As Form, ByVal tbl_name As String)
add_reference "schools.mdb"





frm_name.ScaleMode = 3
Dim obj As Object
Dim x As Integer
x = 65


For i = 0 To db.TableDefs(tbl_name).Fields.Count - 1

Set obj = frm_name.Controls.Add("vb.textbox", db.TableDefs(tbl_name).Fields(i).Name)
With obj
.Left = 96
.Top = x
.Height = 22
.Width = db.TableDefs(tbl_name).Fields(i).ValidationText
.Appearance = 0
.MaxLength = db.TableDefs(tbl_name).Fields(i).Size
.TabIndex = i
.Visible = True

End With

Set obj = frm_name.Controls.Add("vb.label", db.TableDefs(tbl_name).Fields(i).Name & "_lbl")
With obj
.Left = 96 - 70
.Top = x
.Height = 20
.Width = 200
.Visible = True
.Caption = db.TableDefs(tbl_name).Fields(i).Name & ":"
End With


x = x + 30

Next

c = 0
x = 0

End Sub




Function empty_data(frm_name As Form) As Boolean
For Each Control In frm_name
If TypeOf Control Is TextBox Then
If Len(Trim(Control.Text)) = 0 Then
Control.BackColor = &HFFC0FF
empty_data = True
Else
Control.BackColor = vbWhite
End If
End If
Next
End Function





Sub close_data()
rs.Close
db.Close
Set eng = Nothing
Set rs = Nothing
Set db = Nothing
End Sub



