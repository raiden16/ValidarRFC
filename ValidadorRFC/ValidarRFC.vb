Imports System.Net

Public Class ValidarRFC

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
    End Sub


    Public Function GetStatusRFC(ByVal RFC As String)

        Dim URL As String
        Dim json As String
        Dim valididx, activeidx, typeidx, limit, dato, lenght As Integer
        Dim valid, active, status As String
        Dim webclient As WebClient
        'Dim reponsebyte As Byte()
        'Dim hola As NameValueCollection


        Try

            'Para meter registros mediante las API REST
            'hola = New NameValueCollection
            'hola.Add("nombre", "Pedro")
            'hola.Add("Apellido", "Perez")
            'reponsebyte = New WebClient().UploadValues(URL, "POST", hola)

            webclient = New WebClient()
            webclient.Headers.Add("X-API-Key", "0660bc2fc26592df72a377e5d81a961c")

            URL = "https://api.satws.com/rfc/validate/" & RFC
            json = webclient.DownloadString(URL)
            json = ArreglarTexto(json, """", " ").ToString.Trim
            valididx = json.IndexOf("valid :")
            activeidx = json.IndexOf(", active :")
            typeidx = json.IndexOf(", type :")
            valid = json.Substring(valididx + 7, activeidx - (valididx + 7)).Trim
            active = json.Substring(activeidx + 10, typeidx - (activeidx + 10)).Trim

            If valid = "true" And active = "true" Then

                cSBOApplication.MessageBox("✔ RFC Valido.")

            ElseIf valid = "false" Then

                cSBOApplication.MessageBox("✖ RFC Invalido.")

            ElseIf active = "false" Then

                cSBOApplication.MessageBox("✖ RFC Inactivo.")

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al validar el RFC. " & ex.Message)

        End Try

    End Function


    Public Function ArreglarTexto(ByVal TextoOriginal As String, ByVal QuitarCaracter As String, ByVal PonerCaracter As String)

        TextoOriginal = TextoOriginal.Replace(QuitarCaracter, PonerCaracter)
        Return TextoOriginal

    End Function


End Class
