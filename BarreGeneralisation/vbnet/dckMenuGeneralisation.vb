Imports System.Windows.Forms
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.DisplayUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Carto

'**
'Nom de la composante : dckMenuGeneralisation
'
'''<summary>
'''Menu qui permet d'afficher les paramètres et les valeurs des dimensions minimales.
'''
'''</summary>
''' 
'''<remarks>
'''Auteur : Michel Pothier
'''</remarks>
''' 
Public Class dckMenuGeneralisation

    Public Sub New(ByVal hook As Object)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Hook = hook


        'Définir l'application
        m_Application = CType(hook, IApplication)

        'Définir le document
        m_MxDocument = CType(m_Application.Document, IMxDocument)

        'Initialiser le menu
        Me.Init()

        'Conserver le menu en mémoire
        m_MenuGeneralisation = Me


    End Sub


    Private m_hook As Object
    ''' <summary>
    ''' Host object of the dockable window
    ''' </summary> 
    Public Property Hook() As Object
        Get
            Return m_hook
        End Get
        Set(ByVal value As Object)
            m_hook = value
        End Set
    End Property

    ''' <summary>
    ''' Implementation class of the dockable window add-in. It is responsible for
    ''' creating and disposing the user interface class for the dockable window.
    ''' </summary>
    Public Class AddinImpl
        Inherits ESRI.ArcGIS.Desktop.AddIns.DockableWindow

        Private m_windowUI As dckMenuGeneralisation

        Protected Overrides Function OnCreateChild() As System.IntPtr
            m_windowUI = New dckMenuGeneralisation(Me.Hook)
            Return m_windowUI.Handle
        End Function

        Protected Overrides Sub Dispose(ByVal Param As Boolean)
            If m_windowUI IsNot Nothing Then
                m_windowUI.Dispose(Param)
            End If

            MyBase.Dispose(Param)
        End Sub
    End Class

    Private Sub chkCreerFichierErreurs_CheckedChanged(sender As Object, e As EventArgs) Handles chkCreerFichierErreurs.CheckedChanged
        m_CreerFichierErreurs = chkCreerFichierErreurs.Checked
    End Sub

    Private Sub cboAttributsAdjacence_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAttributsAdjacence.SelectedIndexChanged, cboAttributsAdjacence.LostFocus
        m_AttributsAdjacence = cboAttributsAdjacence.Text
    End Sub

    Private Sub cboMethodeSquelette_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMethodeSquelette.SelectedIndexChanged, cboMethodeSquelette.LostFocus
        Try
            m_MethodeSquelette = CInt(cboMethodeSquelette.Text.Split(CChar(":"))(0))

        Catch ex As Exception
            If m_MethodeSquelette = 0 Then
                cboMethodeSquelette.Text = "0:Delaunay"
            Else
                cboMethodeSquelette.Text = "1:Voronoi"
            End If
        End Try
    End Sub

    Private Sub cboEchelle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEchelle.SelectedIndexChanged, cboEchelle.LostFocus
        Try
            m_Echelle = CInt(cboEchelle.Text)

            If m_Echelle = 50000 Then
                cboDistanceLaterale.Text = "1.5"
                cboDistanceDensifier.Text = "15.0"
                cboLongueurDroite.Text = "3.0"
                cboLongueurLigne.Text = "250.0"
                cboSuperficieExterieure.Text = "3600.0"
                cboSuperficieInterieure.Text = "1800.0"
                cboLargeurInterieure.Text = "25.0"
                cboLongueurInterieure.Text = "50.0"
                cboLargeurExterieure.Text = "25.0"
                cboLongueurExterieure.Text = "5.0"

            ElseIf m_Echelle = 250000 Then
                cboDistanceLaterale.Text = "7.5"
                cboDistanceDensifier.Text = "75.0"
                cboLongueurDroite.Text = "15.0"
                cboLongueurLigne.Text = "1500.0"
                cboSuperficieExterieure.Text = "70000.0"
                cboSuperficieInterieure.Text = "35000.0"
                cboLargeurInterieure.Text = "125.0"
                cboLongueurInterieure.Text = "250.0"
                cboLargeurExterieure.Text = "125.0"
                cboLongueurExterieure.Text = "25.0"
            End If

        Catch ex As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & ex.Message & vbCrLf & "--Source: " & ex.Source & vbCrLf & "--StackTrace: " & ex.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub cboPrecision_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPrecision.SelectedIndexChanged, cboPrecision.LostFocus
        Try
            m_Precision = CDbl(cboPrecision.Text)
            Call AjouterItem(cboPrecision)

        Catch ex As Exception
            cboPrecision.Text = m_Precision.ToString
        End Try
    End Sub

    Private Sub cboDistanceLaterale_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDistanceLaterale.SelectedIndexChanged, cboDistanceLaterale.LostFocus
        Try
            m_DistanceLaterale = CDbl(cboDistanceLaterale.Text)
            Call AjouterItem(cboDistanceLaterale)

        Catch ex As Exception
            cboDistanceLaterale.Text = m_DistanceLaterale.ToString
        End Try
    End Sub

    Private Sub cboDistanceDensifier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDistanceDensifier.SelectedIndexChanged, cboDistanceDensifier.LostFocus
        Try
            m_DistanceDensifier = CDbl(cboDistanceDensifier.Text)
            Call AjouterItem(cboDistanceDensifier)

        Catch ex As Exception
            cboDistanceDensifier.Text = m_DistanceDensifier.ToString
        End Try
    End Sub

    Private Sub cboLongueurDroite_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLongueurDroite.SelectedIndexChanged, cboLongueurDroite.LostFocus
        Try
            m_LongueurDroite = CDbl(cboLongueurDroite.Text)
            Call AjouterItem(cboLongueurDroite)

        Catch ex As Exception
            cboLongueurDroite.Text = m_LongueurDroite.ToString
        End Try
    End Sub

    Private Sub cboLongueurLigne_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLongueurLigne.SelectedIndexChanged, cboLongueurLigne.LostFocus
        Try
            m_LongueurLigne = CDbl(cboLongueurLigne.Text)
            Call AjouterItem(cboLongueurLigne)

        Catch ex As Exception
            cboLongueurLigne.Text = m_LongueurLigne.ToString
        End Try
    End Sub

    Private Sub cboSuperficieExterieure_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSuperficieExterieure.SelectedIndexChanged, cboSuperficieExterieure.LostFocus
        Try
            m_SuperficieExterieure = CDbl(cboSuperficieExterieure.Text)
            Call AjouterItem(cboSuperficieExterieure)

        Catch ex As Exception
            cboSuperficieExterieure.Text = m_SuperficieExterieure.ToString
        End Try
    End Sub

    Private Sub cboSuperficieInterieure_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSuperficieInterieure.SelectedIndexChanged, cboSuperficieInterieure.LostFocus
        Try
            m_SuperficieInterieure = CDbl(cboSuperficieInterieure.Text)
            Call AjouterItem(cboSuperficieInterieure)

        Catch ex As Exception
            cboSuperficieInterieure.Text = m_SuperficieInterieure.ToString
        End Try
    End Sub

    Private Sub cboLargeurInterieure_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLargeurInterieure.SelectedIndexChanged, cboLargeurInterieure.LostFocus
        Try
            m_LargeurInterieure = CDbl(cboLargeurInterieure.Text)
            Call AjouterItem(cboLargeurInterieure)

        Catch ex As Exception
            cboLargeurInterieure.Text = m_LargeurInterieure.ToString
        End Try
    End Sub

    Private Sub cboLongueurInterieure_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLongueurInterieure.SelectedIndexChanged, cboLongueurInterieure.LostFocus
        Try
            m_LongueurInterieure = CDbl(cboLongueurInterieure.Text)
            Call AjouterItem(cboLongueurInterieure)

        Catch ex As Exception
            cboLongueurInterieure.Text = m_LongueurInterieure.ToString
        End Try
    End Sub

    Private Sub cboLargeurExterieure_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLargeurExterieure.SelectedIndexChanged, cboLargeurExterieure.LostFocus
        Try
            m_LargeurExterieure = CDbl(cboLargeurExterieure.Text)
            Call AjouterItem(cboLargeurExterieure)

        Catch ex As Exception
            cboLargeurExterieure.Text = m_LargeurExterieure.ToString
        End Try
    End Sub

    Private Sub cboLongueurExterieure_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLongueurExterieure.SelectedIndexChanged, cboLongueurExterieure.LostFocus
        Try
            m_LongueurExterieure = CDbl(cboLongueurExterieure.Text)
            Call AjouterItem(cboLongueurExterieure)

        Catch ex As Exception
            cboLongueurExterieure.Text = m_LongueurExterieure.ToString
        End Try
    End Sub

    Private Sub btnInitialiser_Click(sender As Object, e As EventArgs) Handles btnInitialiser.Click
        Try
            Me.Init()

        Catch ex As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & ex.Message & vbCrLf & "--Source: " & ex.Source & vbCrLf & "--StackTrace: " & ex.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub cboClasseSquelette_GotFocus(sender As Object, e As EventArgs) Handles cboClasseSquelette.GotFocus
        Try
            'Remplir le ComboBox des FeatureLayer
            Call RemplirComboBox(cboClasseSquelette)

        Catch ex As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & ex.Message & vbCrLf & "--Source: " & ex.Source & vbCrLf & "--StackTrace: " & ex.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub cboClasseSquelette_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClasseSquelette.SelectedIndexChanged, cboClasseSquelette.LostFocus
        'Déclarer la variables de travail
        Dim oGererMapLayer As clsGererMapLayer = Nothing    'Objet utilisé pour extraire les FeatureLayer.
        Dim qFeatureLayerColl As Collection = Nothing       'Contient la liste des FeatureLayer visibles
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant une classe de données.

        'Définir le FeatureLayer de la classe du squelette par défaut
        m_ClasseSquelette = Nothing

        Try
            'Définir l'objet pour extraire les FeatureLayer
            oGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap, True)
            'Définir la liste des FeatureLayer de type ligne
            qFeatureLayerColl = oGererMapLayer.DefinirCollectionFeatureLayer(False, esriGeometryType.esriGeometryPolyline)
            'Vérifier si les FeatureLayers sont présents
            If qFeatureLayerColl IsNot Nothing Then
                'Traiter tous les FeatureLayer
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    'Vérifier si le nom du FeatureLayer correspond à celui sélectionné
                    If pFeatureLayer.Name = cboClasseSquelette.Text Then
                        'Définir le FeatureLayer de la classe du squelette
                        m_ClasseSquelette = pFeatureLayer
                    End If
                Next
            End If

        Catch ex As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & ex.Message & vbCrLf & "--Source: " & ex.Source & vbCrLf & "--StackTrace: " & ex.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oGererMapLayer = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
        End Try
    End Sub

    Private Sub cboClasseGeneraliser_GotFocus(sender As Object, e As EventArgs) Handles cboClasseGeneraliser.GotFocus
        Try
            'Remplir le ComboBox des FeatureLayer
            Call RemplirComboBox(cboClasseGeneraliser)

        Catch ex As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & ex.Message & vbCrLf & "--Source: " & ex.Source & vbCrLf & "--StackTrace: " & ex.StackTrace & vbCrLf)
        End Try
    End Sub

    Private Sub cboClasseGeneraliser_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClasseGeneraliser.SelectedIndexChanged, cboClasseGeneraliser.LostFocus
        'Déclarer la variables de travail
        Dim oGererMapLayer As clsGererMapLayer = Nothing    'Objet utilisé pour extraire les FeatureLayer.
        Dim qFeatureLayerColl As Collection = Nothing       'Contient la liste des FeatureLayer visibles
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant une classe de données.

        'Définir le FeatureLayer de la classe du squelette par défaut
        m_ClasseGeneraliser = Nothing

        Try
            'Définir l'objet pour extraire les FeatureLayer
            oGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap, True)
            'Définir la liste des FeatureLayer de type ligne
            qFeatureLayerColl = oGererMapLayer.DefinirCollectionFeatureLayer(False, esriGeometryType.esriGeometryPolyline)
            'Vérifier si les FeatureLayers sont présents
            If qFeatureLayerColl IsNot Nothing Then
                'Traiter tous les FeatureLayer
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    'Vérifier si le nom du FeatureLayer correspond à celui sélectionné
                    If pFeatureLayer.Name = cboClasseGeneraliser.Text Then
                        'Définir le FeatureLayer de la classe pour généraliser
                        m_ClasseGeneraliser = pFeatureLayer
                    End If
                Next
            End If

        Catch ex As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & ex.Message & vbCrLf & "--Source: " & ex.Source & vbCrLf & "--StackTrace: " & ex.StackTrace & vbCrLf)
        Finally
            'Vider la mémoire
            oGererMapLayer = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
        End Try
    End Sub

#Region "Routines et fonctions publiques"
    '''<summary>
    ''' Permet d'initialiser le menu.
    '''</summary>
    '''
    Public Sub Init()
        Try
            'Définir les valeurs par défaut
            tabEdition.SelectTab(0)
            cboEchelle.Text = "250000"
            cboPrecision.Text = "0.001"
            chkCreerFichierErreurs.Checked = False
            cboAttributsAdjacence.Text = ""
            cboClasseSquelette.Text = ""
            rtbMessages.Text = ""
            m_ClasseSquelette = Nothing
            m_ClasseGeneraliser = Nothing

            'Remplir le ComboBox des FeatureLayer
            Call RemplirComboBox(cboClasseSquelette)
            'Remplir le ComboBox des FeatureLayer
            Call RemplirComboBox(cboClasseGeneraliser)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    '''<summary>
    ''' Permet d'ajouter une valeur dans les items mais seulement si la valeur est absent.
    '''</summary>
    '''
    Private Sub AjouterItem(ByRef cboValeur As ComboBox)
        Try
            If cboValeur.FindString(cboValeur.Text) = -1 Then
                cboValeur.Items.Add(cboValeur.Text)
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de remplir le ComboBox à partir des noms de FeatureLayer contenus dans la Map active.
    '''</summary>
    ''' 
    Public Sub RemplirComboBox(ByRef cboClasse As ComboBox)
        'Déclarer la variables de travail
        Dim oGererMapLayer As clsGererMapLayer = Nothing    'Objet utilisé pour extraire les FeatureLayer.
        Dim qFeatureLayerColl As Collection = Nothing       'Contient la liste des FeatureLayer visibles
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant une classe de données
        Dim sNom As String = Nothing                        'Nom du FeatureLayer sélectionné
        Dim iDefaut As Integer = Nothing                    'Index par défaut
        Dim i As Integer = Nothing                          'Compteur

        Try
            'Initialiser le nom du FeatureLayer
            sNom = cboClasse.Text
            'Effacer tous les FeatureLayer existants
            cboClasse.Items.Clear()
            'Vider le nom de la classe du squelette
            cboClasse.Text = ""

            'Définir l'objet pour extraire les FeatureLayer
            oGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap, True)
            'Définir la liste des FeatureLayer de type ligne
            qFeatureLayerColl = oGererMapLayer.DefinirCollectionFeatureLayer(False, esriGeometryType.esriGeometryPolyline)
            'Vérifier si les FeatureLayers sont présents
            If qFeatureLayerColl IsNot Nothing Then
                'Traiter tous les FeatureLayer
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    'Ajouter le FeatureLayer
                    iDefaut = cboClasse.Items.Add(pFeatureLayer.Name)
                    'Vérifier si le nom du FeatureLayer correspond à celui sélectionné
                    If pFeatureLayer.Name = sNom Then
                        'Sélectionné la valeur par défaut
                        cboClasse.Select(iDefaut, 1)
                        cboClasse.Text = pFeatureLayer.Name
                    End If
                Next

                'Véfirier si on peut sélectionner une classe du squelette
                If cboClasse.Text.Length = 0 And qFeatureLayerColl.Count > 0 Then
                    'Sélectionné la valeur par défaut
                    cboClasse.Select(0, 1)
                    cboClasse.Text = pFeatureLayer.Name
                End If
            End If

            'Ajouter pour aucun FeatureLayer 
            cboClasse.Items.Add("")

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            oGererMapLayer = Nothing
            qFeatureLayerColl = Nothing
            pFeatureLayer = Nothing
            sNom = Nothing
            iDefaut = Nothing
            i = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de modifier les items du menu selon la dimension du menu.
    '''</summary>
    ''' 
    Private Sub dckMenuGeneralisation_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        'Déclarer les variables de travail
        Dim iDeltaHeight As Integer
        Dim iDeltaWidth As Integer

        Try
            'Calculer les deltas
            iDeltaHeight = Me.Height - m_Height
            iDeltaWidth = Me.Width - m_Width

            'Redimensionner les objets du menu
            tabEdition.Height = tabEdition.Height + iDeltaHeight
            tabEdition.Width = tabEdition.Width + iDeltaWidth
            rtbMessages.Height = rtbMessages.Height + iDeltaHeight
            rtbMessages.Width = rtbMessages.Width + iDeltaWidth
            btnInitialiser.Top = btnInitialiser.Top + iDeltaHeight

            'Initialiser les variables de dimension
            m_Height = Me.Height
            m_Width = Me.Width

        Catch ex As Exception
            'Afficher un message d'erreur
            MsgBox("--Message: " & ex.Message & vbCrLf & "--Source: " & ex.Source & vbCrLf & "--StackTrace: " & ex.StackTrace & vbCrLf)
        End Try
    End Sub
#End Region
End Class