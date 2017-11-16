Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI

'**
'Nom de la composante : cmdActiverMenu 
'
'''<summary>
''' Permet d'activer un menu pour afficher l'information des paramètres d'édition.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 15 Août 2017
'''</remarks>
'''
Public Class cmdActiverMenu
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private Shared _DockWindow As ESRI.ArcGIS.Framework.IDockableWindow
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
        Try
            'Définir les variables de travail
            Dim windowID As UID = New UIDClass

            'Vérifier si l'application est définie
            If Not Hook Is Nothing Then
                'Définir l'application
                gpApp = CType(Hook, IApplication)

                'Vérifier si l'application est ArcMap
                If TypeOf Hook Is IMxApplication Then
                    'Rendre active la commande
                    Enabled = True
                    'Définir le document
                    gpMxDoc = CType(gpApp.Document, IMxDocument)

                    'Créer un nouveau menu
                    windowID.Value = "BarreGeneralisation_dckMenuGeneralisation"
                    _DockWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(windowID)

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        Try
            'Sortir si le menu n'est pas créé
            If _DockWindow Is Nothing Then Return

            'Activer ou désactiver le menu
            _DockWindow.Show((Not _DockWindow.IsVisible()))
            Checked = _DockWindow.IsVisible()

        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Try
            Enabled = My.ArcMap.Application IsNot Nothing
        Catch erreur As Exception
            MsgBox(erreur.ToString)
        End Try
    End Sub
End Class
