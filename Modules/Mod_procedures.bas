Attribute VB_Name = "Mod_procedures"
Option Explicit

Public Function IfNull(valeur, pardef)

If IsNull(valeur) Then
    IfNull = pardef
Else
    IfNull = valeur
End If

End Function

Public Function Update_TB(F As Form, Table As String, txt As Integer, Chk As Integer, cbo As Integer, mask As Integer, dtp As Integer, condition As String, code_table As String, exe As Integer) As Boolean

Dim Sql As String
Dim oText As TextBox
Dim oCheck As CheckBox
Dim oCbo As ComboBox
'Dim oMask As MaskEdBox
Dim rep

Update_TB = True

''on error goto erreur

rep = MsgBox("Etes vous sûr de modifier cet enregistrement ?", vbYesNo)
If rep <> 6 Then
    Update_TB = False
    Exit Function
End If
Sql = "update " & Table & " set "
  
  
'If mask = 1 Then
'    For Each oMask In F.MaskEdBox
'        If Mid(oMask.Tag, 3) = "1" Or oMask.Tag = "" Then
'            Select Case oMask.Text
'                Case "__/__/____"
'                    sql = sql & oMask.DataField & "=NULL,"
'                Case "   :  "
'                    sql = sql & oMask.DataField & "=NULL,"
'                Case "     "
'                    sql = sql & oMask.DataField & "=NULL,"
'                Case Else
'                    Select Case oMask.mask
'                        Case "#####"
'                            sql = sql & oMask.DataField & "=" & Replace(oMask.Text, " ", "") & ","
'                        Case Else
'                            sql = sql & oMask.DataField & "='" & oMask.Text & "',"
'                    End Select
'            End Select
'        End If
'    Next
'End If

If txt = 1 Then
    For Each oText In F.txtFields
        If Trim$(oText.DataField) <> "" Then
            If Mid(oText.Tag, 3) = "1" Or oText.Tag = "" Then
                If Trim$(oText.Text) = "<Rien>" Or Trim$(oText.Text) = "" Then
                    Sql = Sql & oText.DataField & "=NULL,"
                Else
                    Sql = Sql & oText.DataField & "='" & Trim$(Replace(oText.Text, "'", "''")) & "',"
                End If
            End If
        End If
    Next
End If

If Chk = 1 Then
    For Each oCheck In F.ChkFields
        If Mid(oCheck.Tag, 3) = "1" Or oCheck.Tag = "" Or oCheck.Tag = "V" Then
            Sql = Sql & oCheck.DataField & "='" & oCheck.Value & "',"
        End If
    Next
End If

If cbo = 1 Then
    For Each oCbo In F.CboFields
        If Trim$(oCbo.DataField) <> "" Then
            If Mid(oCbo.Tag, 3) = "1" Or oCbo.Tag = "" Then
              If Trim$(oCbo.Text) <> "" And Trim$(oCbo.Text) <> "<Rien>" Then
                  Sql = Sql & oCbo.DataField & "='" & Trim$(Replace(oCbo.Text, "'", "''")) & "',"
              Else
                  Sql = Sql & oCbo.DataField & "=NULL,"
              End If
            End If
        End If
    Next
End If

Sql = Mid(Sql, 1, Len(Sql) - 1)

Sql = Sql & " where " & condition

If exe = 1 Then
    db.Execute Sql
    db.Execute "insert ZZ_QUERY (SQL) values ('" & Replace(Sql, "'", "''") & "')"
Else
    If Table = "GRID" Then
        Frm_grille_maj.TXTSQL.Text = Sql
    End If
    If Table = "GRID_COLONNE" Then
        Frm_grille_colonne_maj.TXTSQL.Text = Sql
    End If
    If Table = "CONTROLE" Then
        Frm_grille_colonne_maj.TXTSQL.Text = Sql
    End If
End If
Exit Function

erreur:
MsgBox Err.Description
Update_TB = False


End Function

Public Function Insert_TB(F As Form, Table As String, txt As Integer, Chk As Integer, cbo As Integer, mask As Integer, dtp As Integer, List As Integer, code_table As String, exe As Integer) As Integer
Dim Sql As String
Dim oText As TextBox
Dim oCheck As CheckBox
Dim oCbo As ComboBox
'Dim oMask As MaskEdBox
Dim oList As ListBox
Dim rep

Insert_TB = True

'on error goto erreur

rep = MsgBox("Etes vous sûr d'ajouter cet enregistrement ?", vbYesNo)
If rep <> 6 Then
    Insert_TB = False
    Exit Function
End If
        
Sql = "insert " & Table & "  ("
  
If List = 1 Then
    For Each oList In F.ListBox
      If Mid(oList.Tag, 3) = "1" Or oList.Tag = "" Or oList.Tag = "C" Or oList.Tag = "F" Then
        If Trim$(oList.Text) <> "" And Trim$(oList.Text) <> "<Rien>" Then Sql = Sql & oList.DataField & ","
      End If
    Next
End If
    
'If mask = 1 Then
'    For Each oMask In F.MaskEdBox
'      If Mid(oMask.Tag, 3) = "1" Or oMask.Tag = "" Or oMask.Tag = "C" Or oMask.Tag = "F" Then
'        If oMask.Text <> "__/__/____" And oMask.Text <> "   :  " And oMask.Text <> "     " Then sql = sql & oMask.DataField & ","
'      End If
'    Next
'End If
  
If txt = 1 Then
    For Each oText In F.txtFields
      If Trim$(oText.DataField) <> "" Then
        If Mid(oText.Tag, 3) = "1" Or oText.Tag = "" Or oText.Tag = "C" Or oText.Tag = "F" Then
              'If Trim$(oText.Text) <> "" And Trim$(oText.Text) <> "<Rien>" Then
              Sql = Sql & oText.DataField & ","
        End If
       End If
    Next
End If
If Chk = 1 Then
    For Each oCheck In F.ChkFields
      If Mid(oCheck.Tag, 3) = "1" Or oCheck.Tag = "" Or oCheck.Tag = "C" Or oCheck.Tag = "F" Or oCheck.Tag = "V" Then
        Sql = Sql & oCheck.DataField & ","
      End If
    Next
End If
If cbo = 1 Then
    For Each oCbo In F.CboFields
      If Trim$(oCbo.DataField) <> "" Then
        If Mid(oCbo.Tag, 3) = "1" Or oCbo.Tag = "" Or oCbo.Tag = "C" Or oCbo.Tag = "F" Then
          If Trim$(oCbo.Text) <> "" And Trim$(oCbo.Text) <> "<Rien>" Then Sql = Sql & oCbo.DataField & ","
        End If
      End If
    Next
End If
Sql = Mid(Sql, 1, Len(Sql) - 1)
Sql = Sql & ") values ("

'List1.list (List1.ListIndex)

If List = 1 Then
    For Each oList In F.ListBox
      If Mid(oList.Tag, 3) = "1" Or oList.Tag = "" Or oList.Tag = "C" Or oList.Tag = "F" Then
        If Trim$(oList.Text) <> "" And Trim$(oList.Text) <> "<Rien>" Then
            Sql = Sql & "'" & Trim$(oList.List(oList.ListIndex)) & "',"
        Else
            Sql = Sql & "NULL,"
        End If
      End If
    Next
End If

'If mask = 1 Then
'    For Each oMask In F.MaskEdBox
'      If Mid(oMask.Tag, 3) = "1" Or oMask.Tag = "" Or oMask.Tag = "C" Or oMask.Tag = "F" Then
'        If oMask.Text <> "__/__/____" And oMask.Text <> "   :  " And oMask.Text <> "     " Then sql = sql & "'" & Trim$(Replace(oMask.Text, "'", "''")) & "',"
'      End If
'    Next
'End If

If txt = 1 Then
    For Each oText In F.txtFields
      If Mid(oText.Tag, 3) = "1" Or oText.Tag = "" Or oText.Tag = "C" Or oText.Tag = "F" Then
        If Trim$(oText.Text) <> "" And Trim$(oText.Text) <> "<Rien>" Then
            Sql = Sql & "'" & Trim$(Replace(oText.Text, "'", "''")) & "',"
        Else
            Sql = Sql & "NULL,"
        End If
      End If
    Next
End If
 
If Chk = 1 Then
    For Each oCheck In F.ChkFields
      If Mid(oCheck.Tag, 3) = "1" Or oCheck.Tag = "" Or oCheck.Tag = "C" Or oCheck.Tag = "F" Or oCheck.Tag = "V" Then
        Sql = Sql & "'" & oCheck.Value & "',"
      End If
    Next
End If

If cbo = 1 Then
    For Each oCbo In F.CboFields
      If Mid(oCbo.Tag, 3) = "1" Or oCbo.Tag = "" Or oCbo.Tag = "C" Or oCbo.Tag = "F" Then
        If Trim$(oCbo.Text) <> "" And Trim$(oCbo.Text) <> "<Rien>" Then Sql = Sql & "'" & Trim$(Replace(oCbo.Text, "'", "''")) & "',"
      End If
    Next
End If

Sql = Mid(Sql, 1, Len(Sql) - 1)
Sql = Sql & ")"

If exe = 1 Then
    db.Execute Sql
    db.Execute "insert ZZ_QUERY (SQL) values ('" & Replace(Sql, "'", "''") & "')"
Else
    If Table = "GRID" Then
        Frm_grille_maj.TXTSQL.Text = Sql
    End If
    If Table = "GRID_COLONNE" Then
        Frm_grille_colonne_maj.TXTSQL.Text = Sql
    End If
    If Table = "CONTROLE" Then
        Frm_control_maj.TXTSQL.Text = Sql
    End If
End If
Exit Function

erreur:
MsgBox Err.Description
Insert_TB = False

End Function

Public Sub message_info(F As Form, message As String, Optional vbType = vbInformation)
Dim Sql As String
Dim rs As New Recordset
    
''on error goto erreur
    
Sql = "select * from message_info where fenetre='" & Replace(F.Caption, "'", "''") & "' and message_info='" & Replace(message, "'", "''") & "' "
rs.Open Sql, db, adOpenStatic, adLockOptimistic
If Not rs.EOF Then
    MsgBox rs("code") & " - " & rs("message_info"), vbType
    rs.Close
Else
    rs.Close
    Sql = "insert message_info (fenetre,message_info_system,message_info) values ('" & Replace(F.Caption, "'", "''") & "','" & Replace(message, "'", "''") & "','" & Replace(message, "'", "''") & "') "
    db.Execute Sql
    
    Sql = "select * from message_info where fenetre='" & Replace(F.Caption, "'", "''") & "' and message_info='" & Replace(message, "'", "''") & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    MsgBox rs("code") & " - " & message, vbType
    rs.Close
End If

Exit Sub

erreur:
MsgBox Err.Description

End Sub

