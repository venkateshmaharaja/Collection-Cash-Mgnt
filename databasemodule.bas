Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As ADODB.Recordset
Public UserName As String
Public Rights As String
Public Status As String
Sub Connect()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
        & App.Path & "\project.mdb;Persist Security Info=False"
rs.Open "Select * from SALES ", con, adOpenDynamic, adLockOptimistic


End Sub



