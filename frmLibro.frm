VERSION 5.00
Begin VB.Form frmLibro 
   Caption         =   "Agrega un libro"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9045
   LinkTopic       =   "Form2"
   ScaleHeight     =   8640
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPrestado 
      Caption         =   "Prestado actualmente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   5760
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   14
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox txtPrestadoA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   6480
      Width           =   4815
   End
   Begin VB.CheckBox chkRecomendado 
      Caption         =   "Recomendado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   5040
      Width           =   4815
   End
   Begin VB.CheckBox chkPorLeer 
      Caption         =   "Quiero leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   4320
      Width           =   4815
   End
   Begin VB.CheckBox chkLeido 
      Caption         =   "Ya leido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   3600
      Width           =   4815
   End
   Begin VB.TextBox txtCalificacion 
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.ComboBox cboGenero 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   2160
      Width           =   4815
   End
   Begin VB.TextBox txtAutor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtTitulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "Prestado a:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Calificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditandoID As Integer

Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()

End Sub
Private Sub chkPrestado_Click()
    If chkPrestado.Value = 1 Then
        txtPrestadoA.Enabled = True
    Else
        txtPrestadoA.Enabled = False
        txtPrestadoA.Text = ""
    End If
        
End Sub

Private Sub chkLeido_Click()
    If chkLeido.Value = 1 Then
       chkPorLeer.Value = 0
       txtCalificacion.Enabled = True
    Else
        txtCalificacion.Enabled = False
    End If
End Sub

Private Sub chkPorLeer_Click()
    If chkPorLeer.Value = 1 Then
        chkLeido.Value = 0
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If Trim(txtTitulo.Text) = "" Or Trim(txtAutor.Text) = "" Then
        MsgBox "El titulo y el autor son obligatorios", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    If cboGenero.ListIndex = -1 Then
        MsgBox "Seleccione un genero", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    If chkLeido.Value = 1 And Trim(txtCalificacion.Text) = "" Then
        MsgBox "Por favor ingrese una clasificacion del 1 al 5", vbExclamation
    End If
    
    'Calificacion
    Dim calif As Variant
    If Trim(txtCalificacion.Text) <> "" Then
        calif = Val(txtCalificacion.Text)
        If (calif < 1 Or calif > 5) Then
            MsgBox "Calificacion debe ser del 1 al 5", vbExclamation
            Exit Sub
        End If
    Else
        calif = "NULL"
    End If
    
    'Datos
    Dim titulo As String, autor As String, generoID As Long
    titulo = Replace(txtTitulo.Text, "'", "''")
    autor = Replace(txtAutor.Text, "'", "''")
    generoID = cboGenero.ItemData(cboGenero.ListIndex)
    
    Dim leido As Integer, porLeer As Integer, recom As Integer, prestado As Integer
    leido = IIf(chkLeido.Value = 1, 1, 0)
    porLeer = IIf(chkPorLeer.Value = 1, 1, 0)
    recom = IIf(chkRecomendado.Value = 1, 1, 0)
    prestado = IIf(chkPrestado.Value = 1, 1, 0)
    
    Dim prestadoA As String, fechaPrestamo As String
    If prestado = 1 Then
        prestadoA = Replace(txtAutor.Text, "'", "''")
        fechaPrestamo = Format$(Now, "yyyy-mm-dd")
    Else
        prestadoA = ""
        fechaPrestamo = ""
    End If
    
    On Error GoTo ErrSave
        
        Dim sqlInsert As String
        sqlInsert = "INSERT INTO Libros (Titulo, Autor, GeneroID, Calificacion, Leido, PorLeer, Recomendado, Prestado, PrestadoA, FechaPrestamo) VALUES ('" & titulo & "', '" & autor & "', '" & CStr(generoID) & "', "
        
        If calif = "NULL" Then
            sqlInsert = sqlInsert & "NULL"
        Else
            sqlInsert = sqlInsert & CStr(calif)
        End If
            
        sqlInsert = sqlInsert & ", " & CStr(leido) & ", " & CStr(porLeer) & ", " & CStr(recom) & ", " & CStr(prestado) & ", "
        
        If prestado = 1 Then
            sqlInsert = sqlInsert & "'" & prestadoA & "', '" & fechaPrestamo & "')"
        Else
            sqlInsert = sqlInsert & "NULL, NULL)"
        End If
        
        MsgBox sqlInsert, vbInformation
            
        conn.Execute sqlInsert
        MsgBox "Libro guardado con exito", vbInformation
        Exit Sub
    
ErrSave:
    MsgBox "Error al guardar: " & Err.Description, vbCritical
    
End Sub

Private Sub Form_Load()
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    rsG.Open "SELECT generoID, Nombre FROM Generos ORDER BY Nombre", conn, adOpenStatic, adLockReadOnly
    cboGenero.Clear
    Do Until rsG.EOF
        cboGenero.AddItem rsG!Nombre
        cboGenero.ItemData(cboGenero.NewIndex) = rsG!generoID
        rsG.MoveNext
    Loop
    
    rsG.Close: Set rsG = Nothing
    
    If EditandoID = 0 Then
        ' Modo agregar, limpiar campos
        txtTitulo.Text = ""
        txtAutor = ""
        cboGenero.ListIndex = -1 'No hay nada
        txtCalificacion = ""
        chkLeido.Value = 0
        txtPrestadoA.Enabled = False
        Me.Caption = "Agregar Libro"
        Me.Caption = "Titulo"
        Me.Caption = "Autor"
        Me.Caption = "Genero"
        Me.Caption = "Calificación"
        Me.Caption = "Prestado a:"
        Me.Caption = "Ya leido"
        Me.Caption = "Quiero leer"
        Me.Caption = "Recomendado"
        Me.Caption = "Prestado actualmente"
        Me.Caption = "Guardar"
        Me.Caption = "Cancelar"
    Else
        
    End If
    
End Sub

