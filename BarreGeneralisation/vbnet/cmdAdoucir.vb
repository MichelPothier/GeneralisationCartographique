Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry

'**
'Nom de la composante : cmdAdoucir.vb
'
'''<summary>
''' Commande qui permet d'adoucir les lignes et les limites des polygones pour les éléments sélectionnés selon une distance minimum.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''Date : 18 septembre 2017
'''</remarks>
''' 
Public Class cmdAdoucir
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Dim gpApp As IApplication = Nothing     'Interface ESRI contenant l'application ArcMap
    Dim gpMxDoc As IMxDocument = Nothing    'Interface ESRI contenant un document ArcMap

    Public Sub New()
        Try
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

                Else
                    'Rendre désactive la commande
                    Enabled = False
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            MsgBox(ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnClick()
        'Déclarer les variables de travail
        Dim pMouseCursor As IMouseCursor = Nothing  'Interface qui permet de changer l'image du curseur
        Dim pTrackCancel As ITrackCancel = Nothing  'Interface qui permet d'annuler la sélection avec la touche ESC.
        Dim dDateDebut As Date = Nothing            'Date de début du traitement.
        Dim dDateFin As Date = Nothing              'Date de fin du traitement.
        Dim dTempsTraitement As TimeSpan = Nothing  'Temps d'exécution du traitement.
        Dim iNbELem As Integer = 0                  'Contient le nombre d'éléments traités.

        Try
            'Changer le curseur en Sablier pour montrer qu'une tâche est en cours
            pMouseCursor = New MouseCursorClass
            pMouseCursor.SetCursor(2)


            'Si la référence spatiale est projeter
            If TypeOf (m_MxDocument.FocusMap.SpatialReference) Is IProjectedCoordinateSystem Then
                'Définir la date de début
                dDateDebut = System.DateTime.Now

                'Permettre d'annuler la sélection avec la touche ESC
                pTrackCancel = New CancelTracker
                pTrackCancel.CancelOnKeyPress = True
                pTrackCancel.CancelOnClick = False

                'Permettre l'affichage de la barre de progression
                iNbELem = m_MxDocument.FocusMap.SelectionCount
                m_Application.StatusBar.ProgressBar.Position = 0
                m_Application.StatusBar.ShowProgressBar("Adoucir les lignes et les limites des polygones ...", 0, iNbELem, 1, True)
                pTrackCancel.Progressor = m_Application.StatusBar.ProgressBar

                'Afficher le message de début du traitement
                m_MenuGeneralisation.tabEdition.SelectTab(2)
                m_MenuGeneralisation.rtbMessages.Text = "Adoucir les lignes et les limites des polygones ... " & vbCrLf & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Début du traitement : " & dDateDebut.ToString & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Précision : " & m_Precision.ToString & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Distance latérale : " & m_DistanceLaterale.ToString & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Distance minimale entre deux sommets : " & m_DistanceDensifier.ToString & vbCrLf

                'Adoucie les lignes et les limites des polygones pour les éléments sélectionnés
                modGeneralisation.Adoucir(m_Precision, m_DistanceLaterale, m_DistanceDensifier, pTrackCancel)

                'Définir la date de fin
                dDateFin = System.DateTime.Now
                'Afficher le message de fin du traitement
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Fin du traitement : " & dDateFin.ToString & vbCrLf & vbCrLf

                'Définir le temp de traitement
                dTempsTraitement = dDateFin.Subtract(dDateDebut).Add(New TimeSpan(5000000))
                'Afficher le message de temps du traitement
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Temps de traitement : " & dTempsTraitement.ToString.Substring(0, 8) & vbCrLf
                'Afficher le nombre d'éléments
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Nombre d'éléments traités : " & iNbELem.ToString & vbCrLf

                'Si la référence spatiale est géographique
            Else
                'Afficher le nombre d'erreurs
                m_MenuGeneralisation.tabEdition.SelectTab(2)
                m_MenuGeneralisation.rtbMessages.Text = "ERREUR : La référence spatiale ne doit pas être géographique!"
            End If

        Catch ex As Exception
            'Message d'erreur
            MsgBox(ex.ToString)
        Finally
            'Cacher la barre de progression
            pTrackCancel.Progressor.Hide()
            'Vider la mémoire
            pMouseCursor = Nothing
            pTrackCancel = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        'Déclarer les variables
        Dim pEditor As IEditor = Nothing    'Interface ESRI pour effectuer l'édition des éléments.

        Try
            'Définir la valeur par défaut
            Enabled = False

            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si au moins un élément est sélectionné et si on est en mode édition
            If gpMxDoc.FocusMap.SelectionCount > 0 And pEditor.EditState = esriEditState.esriStateEditing And m_MenuGeneralisation IsNot Nothing Then
                Enabled = True
            End If

        Catch ex As Exception
            'Message d'erreur
            MsgBox(ex.ToString)
        Finally
            'Vider la mémoire
            pEditor = Nothing
        End Try
    End Sub
End Class
