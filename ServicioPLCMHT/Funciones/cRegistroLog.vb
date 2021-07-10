Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.IO

Public Class cRegistroLog

    Public Sub RegistroLog(ByVal Log As String)

        Dim ruta As String
        Dim fichero As String
        Dim carpeta As String
        Dim rutafinal As String
        Dim path As String

        Try
            ruta = Directory.GetCurrentDirectory()
            carpeta = "Registro Errores"
            fichero = "Log_Errores" & DateAndTime.Now.ToString("dd-MM-yyyy") + ".txt"
            path = ruta & "\" & carpeta
            rutafinal = ruta & "\" & carpeta & "\" & fichero

            Dim oSW As StreamWriter
            Dim linea As String

            If (Directory.Exists(path)) Then
                oSW = File.AppendText(rutafinal)
                linea = DateAndTime.Now.ToString() & ChrW(9) & Log
                oSW.WriteLine(linea)
                oSW.Flush()
                oSW.Close()
            Else
                Dim di As DirectoryInfo
                di = Directory.CreateDirectory(path)
                oSW = New StreamWriter(rutafinal)
                linea = DateAndTime.Now.ToString() & ChrW(9) & Log
                oSW.WriteLine(linea)
                oSW.Flush()
                oSW.Close()

            End If
        Catch ex As Exception

        End Try


    End Sub

End Class
