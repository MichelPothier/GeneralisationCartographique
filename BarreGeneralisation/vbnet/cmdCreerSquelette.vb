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
'Nom de la composante : cmdCreerSquelette.vb
'
'''<summary>
''' Commande qui permet de créer le squelette des éléments sélectionnés de type surface selon une longueur minimale et selon la méthode de Delaunay
''' à l'aide des outils de topologie tout en respectant les éléments en relation.
'''</summary>
'''
'''<remarks>
'''Ce traitement est utilisable en mode interactif à l'aide de ses interfaces graphiques et doit être utilisé dans ArcMap (ArcGisESRI).
'''
'''Auteur : Michel Pothier
'''Date : 05 septembre 2017
'''</remarks>
''' 
Public Class cmdCreerSquelette
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
        Dim pSpatialRefRes As ISpatialReferenceResolution = Nothing 'Interface qui permet d'initialiser la résolution XY.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.
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

                'Initialiser la résolution
                pSpatialRefRes = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceResolution)
                pSpatialRefRes.SetDefaultXYResolution()
                ''Interface pour définir la tolérance XY
                'pSpatialRefTol = CType(m_MxDocument.FocusMap.SpatialReference, ISpatialReferenceTolerance)
                'pSpatialRefTol.XYTolerance = pSpatialRefRes.XYResolution(True) * 2
                'pSpatialRefTol.XYTolerance = 0.001

                'Permettre l'affichage de la barre de progression
                iNbELem = m_MxDocument.FocusMap.SelectionCount
                m_Application.StatusBar.ProgressBar.Position = 0
                m_Application.StatusBar.ShowProgressBar("Création des squellettes ...", 0, iNbELem, 1, True)
                pTrackCancel.Progressor = m_Application.StatusBar.ProgressBar

                'Afficher le message de début du traitement
                m_MenuGeneralisation.tabEdition.SelectTab(2)
                m_MenuGeneralisation.rtbMessages.Text = "Création des squellettes ... " & vbCrLf & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Début du traitement : " & dDateDebut.ToString & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Précision : " & m_Precision.ToString & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Distance minimale entre 2 sommets : " & m_DistanceDensifier.ToString & vbCrLf
                m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Longueur de généralisation minimale : " & m_LongueurInterieure.ToString & vbCrLf
                If m_ClasseSquelette IsNot Nothing Then
                    m_MenuGeneralisation.rtbMessages.Text = m_MenuGeneralisation.rtbMessages.Text & "Classe du squelette : " & m_ClasseSquelette.FeatureClass.AliasName & vbCrLf
                End If

                'Créer le squelette des éléments sélectionnés
                modGeneralisation.CreerSquelettes(m_Precision, m_DistanceLaterale, m_DistanceDensifier, m_LongueurInterieure, m_ClasseSquelette, pTrackCancel)

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
            If pTrackCancel IsNot Nothing Then pTrackCancel.Progressor.Hide()
            'Vider la mémoire
            pMouseCursor = Nothing
            pTrackCancel = Nothing
            pSpatialRefRes = Nothing
            pSpatialRefTol = Nothing
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
            If gpMxDoc.FocusMap.SelectionCount > 0 And pEditor.EditState = esriEditState.esriStateEditing And m_MenuGeneralisation IsNot Nothing And m_ClasseSquelette IsNot Nothing Then
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
