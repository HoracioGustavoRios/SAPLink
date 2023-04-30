'------------------------------------------------------------------------------
' Clase para manejar ficheros INIs
' Permite leer secciones enteras y todas las secciones de un fichero INI
'
' �ltima revisi�n:                                                  (04/Abr/01)
' Para usar con Visual Basic.NET                                    (21/Jul/02)
'
' �Guillermo 'guille' Som, 1997-2002
'------------------------------------------------------------------------------
Option Strict On
Option Explicit On 

Public Class cIniArray

    Private sBuffer As String ' Para usarla en las funciones GetSection(s)

    '--- Declaraciones para leer ficheros INI ---
    ' Leer todas las secciones de un fichero INI, esto seguramente no funciona en Win95
    ' Esta funci�n no estaba en las declaraciones del API que se incluye con el VB
    Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    ' Leer una secci�n completa
    Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    ' Leer una clave de un fichero INI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    ' Escribir una clave de un fichero INI (tambi�n para borrar claves y secciones)
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer

    Public Sub IniDeleteKey(ByVal sIniFile As String, ByVal sSection As String, Optional ByVal sKey As String = "")
        '--------------------------------------------------------------------------
        ' Borrar una clave o entrada de un fichero INI                  (16/Feb/99)
        ' Si no se indica sKey, se borrar� la secci�n indicada en sSection
        ' En otro caso, se supone que es la entrada (clave) lo que se quiere borrar
        '
        ' Para borrar una secci�n se deber�a usar IniDeleteSection
        '
        If Len(sKey) = 0 Then
            ' Borrar una secci�n
            Call WritePrivateProfileString(sSection, 0, 0, sIniFile)
        Else
            ' Borrar una entrada
            Call WritePrivateProfileString(sSection, sKey, 0, sIniFile)
        End If
    End Sub

    Public Sub IniDeleteSection(ByVal sIniFile As String, ByVal sSection As String)
        '--------------------------------------------------------------------------
        ' Borrar una secci�n de un fichero INI                          (04/Abr/01)
        ' Borrar una secci�n
        Call WritePrivateProfileString(sSection, 0, 0, sIniFile)
    End Sub

    Public Function IniGet(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "") As String
        '--------------------------------------------------------------------------
        ' Devuelve el valor de una clave de un fichero INI
        ' Los par�metros son:
        '   sFileName   El fichero INI
        '   sSection    La secci�n de la que se quiere leer
        '   sKeyName    Clave
        '   sDefault    Valor opcional que devolver� si no se encuentra la clave
        '--------------------------------------------------------------------------
        Dim ret As Integer
        Dim sRetVal As String
        '
        sRetVal = New String(Chr(0), 255)
        '
        ret = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, Len(sRetVal), sFileName)
        If ret = 0 Then
            Return sDefault
        Else
            Return Left(sRetVal, ret)
        End If
    End Function

    Public Sub IniWrite(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, ByVal sValue As String)
        '--------------------------------------------------------------------------
        ' Guarda los datos de configuraci�n
        ' Los par�metros son los mismos que en LeerIni
        ' Siendo sValue el valor a guardar
        '
        Call WritePrivateProfileString(sSection, sKeyName, sValue, sFileName)
    End Sub

    Public Function IniGetSection(ByVal sFileName As String, ByVal sSection As String) As String()
        '--------------------------------------------------------------------------
        ' Lee una secci�n entera de un fichero INI                      (27/Feb/99)
        ' Adaptada para devolver un array de string                     (04/Abr/01)
        '
        ' Esta funci�n devolver� un array de �ndice cero
        ' con las claves y valores de la secci�n
        '
        ' Par�metros de entrada:
        '   sFileName   Nombre del fichero INI
        '   sSection    Nombre de la secci�n a leer
        ' Devuelve:
        '   Un array con el nombre de la clave y el valor
        '   Para leer los datos:
        '       For i = 0 To UBound(elArray) -1 Step 2
        '           sClave = elArray(i)
        '           sValor = elArray(i+1)
        '       Next
        '
        Dim aSeccion() As String
        Dim n As Integer
        '
        ReDim aSeccion(0)
        '
        ' El tama�o m�ximo para Windows 95
        sBuffer = New String(ChrW(0), 32767)
        '
        n = GetPrivateProfileSection(sSection, sBuffer, sBuffer.Length, sFileName)
        '
        If n > 0 Then
            '
            ' Cortar la cadena al n�mero de caracteres devueltos
            ' menos los dos �ltimos que indican el final de la cadena
            sBuffer = sBuffer.Substring(0, n - 2).TrimEnd()
            ' Cada elemento estar� separado por un Chr(0)
            ' y cada valor estar� en la forma: clave = valor
            aSeccion = sBuffer.Split(New Char() {ChrW(0), "="c})
        End If
        ' Devolver el array
        Return aSeccion
    End Function

    Public Function IniGetSections(ByVal sFileName As String) As String()
        '--------------------------------------------------------------------------
        ' Devuelve todas las secciones de un fichero INI                (27/Feb/99)
        ' Adaptada para devolver un array de string                     (04/Abr/01)
        '
        ' Esta funci�n devolver� un array con todas las secciones del fichero
        '
        ' Par�metros de entrada:
        '   sFileName   Nombre del fichero INI
        ' Devuelve:
        '   Un array con todos los nombres de las secciones
        '   La primera secci�n estar� en el elemento 1,
        '   por tanto, si el array contiene cero elementos es que no hay secciones
        '
        Dim n As Integer
        Dim aSections() As String
        '
        ReDim aSections(0)
        '
        ' El tama�o m�ximo para Windows 95
        sBuffer = New String(ChrW(0), 32767)
        '
        ' Esta funci�n del API no est� definida en el fichero TXT
        n = GetPrivateProfileSectionNames(sBuffer, sBuffer.Length, sFileName)
        '
        If n > 0 Then
            ' Cortar la cadena al n�mero de caracteres devueltos
            ' menos los dos �ltimos que indican el final de la cadena
            sBuffer = sBuffer.Substring(0, n - 2).TrimEnd()
            aSections = sBuffer.Split(ChrW(0))
        End If
        ' Devolver el array
        Return aSections
    End Function
    '
    Public Shared Function AppPath( _
            Optional ByVal backSlash As Boolean = False _
            ) As String
        ' System.Reflection.Assembly.GetExecutingAssembly...
        Dim s As String = IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.GetCallingAssembly.Location)
        ' si hay que a�adirle el backslash
        If backSlash Then
            s &= "\"
        End If
        Return s
    End Function
End Class
