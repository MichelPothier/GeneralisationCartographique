Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports System

'**
'Nom de la composante : modGeneralisation.vb 
'
'''<summary>
'''Librairie de routines et de fonctions utilisées pour effectuer la squelettisation et la généralisation cartographique.
'''</summary>
'''
'''<remarks>
'''Auteur : Michel Pothier
'''Date : 26 octobre 2017
'''</remarks>
'''
Module modGeneralisation
    'Liste des variables publiques utilisées
    '''<summary> Classe contenant le menu des paramètres de généralisation. </summary>
    Public m_MenuGeneralisation As dckMenuGeneralisation
    ''' <summary>Objet qui permet de gérer les ressources (textes et images).</summary>
    Public m_RessourceManager As System.Resources.ResourceManager = Nothing
    ''' <summary>Interface ESRI contenant l'application ArcMap.</summary>
    Public m_Application As IApplication = Nothing
    ''' <summary>Interface ESRI contenant le document ArcMap.</summary>
    Public m_MxDocument As IMxDocument = Nothing

    ''' <summary>Interface ESRI contenant les paramètres de la topologie courante.</summary>
    Public m_MapTopology As IMapTopology2 = Nothing
    ''' <summary>Interface ESRI contenant la topologie courante.</summary>
    Public m_TopologyGraph As ITopologyGraph4 = Nothing

    ''' <summary>Méthode pour créer le squelette.</summary>
    Public m_CreerFichierErreurs As Boolean = False
    ''' <summary>Méthode pour créer le squelette.</summary>
    Public m_MethodeSquelette As Integer = 0
    ''' <summary>Échele de représentation des données.</summary>
    Public m_Echelle As Integer = 250000
    ''' <summary>Précision des données.</summary>
    Public m_Precision As Double = 0.001
    ''' <summary>Distance latérale mimimum.</summary>
    Public m_DistanceLaterale As Double = 7.5
    ''' <summary>Distanice minimum pour densifier.</summary>
    Public m_DistanceDensifier As Double = 75.0
    ''' <summary>Longueur minimum des droites.</summary>
    Public m_LongueurDroite As Double = 15.0
    ''' <summary>Longueur minimum des lignes.</summary>
    Public m_LongueurLigne As Double = 1500.0
    ''' <summary>Superficie extérieure minimum des anneaux.</summary>
    Public m_SuperficieExterieure As Double = 70000.0
    ''' <summary>Superficie intérieure minimum des anneaux.</summary>
    Public m_SuperficieInterieure As Double = 35000.0
    ''' <summary>Largeur minimum de généralisation intérieure.</summary>
    Public m_LargeurInterieure As Double = 125.0
    ''' <summary>Longueur minimum de généralisation intérieure.</summary>
    Public m_LongueurInterieure As Double = 250.0
    ''' <summary>Largeur minimum de généralisation extérieure.</summary>
    Public m_LargeurExterieure As Double = 125.0
    ''' <summary>Longueur minimum de généralisation extérieure.</summary>
    Public m_LongueurExterieure As Double = 250.0
    ''' <summary>Nom des attributs d'adjacence.</summary>
    Public m_AttributsAdjacence As String = ""
    ''' <summary>Nom de la classe du squelette.</summary>
    Public m_ClasseSquelette As IFeatureLayer = Nothing
    ''' <summary>Nom de la classe pour généraliser.</summary>
    Public m_ClasseGeneraliser As IFeatureLayer = Nothing

    '''<summary>Valeur initiale de la dimension en hauteur du menu.</summary>
    Public m_Height As Integer = 300
    '''<summary>Valeur initiale de la dimension en largeur du menu.</summary>
    Public m_Width As Integer = 300

#Region "Routines et fonctions pour corriger la topologie des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la topologie des éléments sélectionnés selon une précision.
    ''' Les composantes vides des géométries des élément sont également éliminées. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la valeur de la précision utilisé pour simplifier.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerTopologie(ByVal dPrecision As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean,
                                 ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                'Interface ESRI pour effectuer l'édition des éléments.
        Dim pEnumFeature As IEnumFeature = Nothing      'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant la Featureclass d'erreur.
        Dim pFeature As IFeature = Nothing              'Interface ESRI contenant un élément en sélection.
        Dim pBagErreurs As IGeometryBag = Nothing       'Interface contenant les erreurs de géométrie des éléments.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface contenant le nombre de composantes.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer un Bag d'erreurs vide
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour ajouter les erreurs dans le Bag
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la topologie invalide (Précison=" & dPrecision.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Vérifier si une composante de la géométrie de l'élément est vide
                    Call CorrigerComposanteVideElement(bCorriger, pFeature, pBagErreurs)

                    'Corriger la topologie de l'élément
                    Call CorrigerTopologieElement(dPrecision, bCorriger, pFeature, pBagErreurs)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier si on doit corriger
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la topologie invalide")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Si on ne doit pas corriger
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de topologie invalide (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerTopologie_0", "Élément non simple, Topologie invalide : Précision=" & dPrecision.ToString, _
                                                                  m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryMultipoint, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface
            pEditor = Nothing

        Catch erreur As Exception
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pEnumFeature = Nothing
            pFeatureLayer = Nothing
            pFeature = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de corriger les composantes vides de la géométries d'un élément. 
    '''</summary>
    ''' 
    '''<param name="bCorriger"> Indique si on doit corriger l'élément.</param>
    '''<param name="pFeature"> Interface contenant l'élément à corriger.</param>
    '''<param name="pBagErreurs"> Interface contenant les erreurs de géométrie des éléments.</param>
    ''' 
    Public Function CorrigerComposanteVideElement(ByVal bCorriger As Boolean, ByRef pFeature As IFeature, ByRef pBagErreurs As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pWrite As IFeatureClassWrite = Nothing      'Interface qui permet d'écrire l'élément.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie de l'élément.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface contenant le nombre de composantes.

        'Définir la valeur par défaut
        CorrigerComposanteVideElement = False

        Try
            'Traite tous les éléments sélectionnés
            If pFeature IsNot Nothing Then
                'Interface contenant la référence spatiale
                pWrite = CType(pFeature.Class, IFeatureClassWrite)

                'Vérifier si la géométrie de l'élément est absent
                If pFeature.Shape Is Nothing Then
                    'Vérifier si la géométrie est un point
                    If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Définir la géométrie de l'élément
                        pGeometry = New Point
                        'Vérifier si la géométrie est un mulipoint
                    ElseIf pFeature.Shape.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                        'Définir la géométrie de l'élément
                        pGeometry = New Multipoint
                        'Vérifier si la géométrie est une polyligne
                    ElseIf pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                        'Définir la géométrie de l'élément
                        pGeometry = New Polyline
                        'Vérifier si la géométrie est un polygone
                    ElseIf pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                        'Définir la géométrie de l'élément
                        pGeometry = New Polygon
                    End If
                    'Définir la référence spatiale
                    pGeometry.SpatialReference = pBagErreurs.SpatialReference
                    'Interface pour ajouter les erreurs
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)
                    'Ajouter l'erreur
                    pGeomColl.AddGeometry(pGeometry)

                    'Indiquer qu'une composante vide a été trouvée
                    CorrigerComposanteVideElement = True

                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)
                        'Indiquer que l'élément a été détruit
                        pFeature = Nothing
                    End If

                    'Vérifier si la géométrie de l'élément est vide
                ElseIf pFeature.Shape.IsEmpty Then
                    'Définir la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy
                    'Projeter la géométrie
                    pGeometry.Project(pBagErreurs.SpatialReference)
                    'Interface pour ajouter les erreurs
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)
                    'Ajouter l'erreur
                    pGeomColl.AddGeometry(pGeometry)

                    'Indiquer qu'une composante vide a été trouvée
                    CorrigerComposanteVideElement = True

                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)
                        'Indiquer que l'élément a été détruit
                        pFeature = Nothing
                    End If

                    'Vérifier si la géométrie n'est pas un point
                ElseIf pFeature.Shape.GeometryType <> esriGeometryType.esriGeometryPoint Then
                    'Définir la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy
                    'Projeter la géométrie
                    pGeometry.Project(pBagErreurs.SpatialReference)

                    'Interface pour détruire les composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Traiter toutes les composantes de la géométrie
                    For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                        'Vérifier si la composante est vide
                        If pGeomColl.Geometry(i).IsEmpty Then
                            'Indiquer qu'une composante vide a été trouvée
                            CorrigerComposanteVideElement = True

                            'Détruire la composante de la géométrie
                            pGeomColl.RemoveGeometries(i, 1)
                        End If
                    Next

                    'Vérifier si une erreur est présente
                    If CorrigerComposanteVideElement Then
                        'Interface pour ajouter les erreurs
                        pGeomColl = CType(pBagErreurs, IGeometryCollection)
                        'Ajouter l'erreur
                        pGeomColl.AddGeometry(pGeometry)

                        'Vérifier si on doit corriger
                        If bCorriger Then
                            'Définir la géométrie de l'élément
                            pFeature.Shape = pGeometry

                            'Écrire l'élément
                            pWrite.WriteFeature(pFeature)
                        End If
                    End If
                End If
            End If

        Catch erreur As Exception
            Throw
        Finally
            'Vider la mémoire
            pWrite = Nothing
            pGeometry = Nothing
            pGeomColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de corriger la topologie d'un élément et retourner les erreurs. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Précision de la topologie.</param>
    '''<param name="bCorriger"> Indique si on doit corriger la topologie de l'élément.</param>
    '''<param name="pFeature"> Interface contenant l'élément à corriger la topologie.</param>
    '''<param name="pBagErreurs"> Interface contenant les erreurs de géométrie des éléments.</param>
    ''' 
    Public Sub CorrigerTopologieElement(ByVal dPrecision As Double, ByVal bCorriger As Boolean, ByRef pFeature As IFeature, ByRef pBagErreurs As IGeometryBag)
        'Déclarer les variables de travail
        Dim pDifference As IGeometry = Nothing          'Interface contenant les différences de la géométrie de l'élément.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface qui permet de simplifier la géométrie.
        Dim pWrite As IFeatureClassWrite = Nothing      'Interface qui permet d'écrire l'élément.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie de l'élément.
        Dim pGeometryOri As IGeometry = Nothing         'Interface contenant la géométrie originale de l'élément.
        Dim pPointColl As IPointCollection = Nothing    'Interface contenant le nombre de sommets.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface contenant le nombre de composantes.
        Dim dXYTol As Double = 0                        'Contient la résolution originale.
        Dim iNbSommets As Integer = 0                   'Contient le nombre de sommets.
        Dim iNbComposantes As Integer = 0               'Contient le noombre de composantes.

        Try
            'Traite tous les éléments sélectionnés
            If pFeature IsNot Nothing Then
                'Vérifier si la géométrie n'Est pas un point
                If pFeature.Shape.GeometryType <> esriGeometryType.esriGeometryPoint Then
                    'Définir la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy
                    'Projeter la géométrie
                    pGeometry.Project(pBagErreurs.SpatialReference)

                    'Conserver le nombre de sommets
                    pPointColl = CType(pGeometry, IPointCollection)
                    iNbSommets = pPointColl.PointCount
                    'Conserver le nombre de composantes
                    pGeomColl = CType(pGeometry, IGeometryCollection)
                    iNbComposantes = pGeomColl.GeometryCount

                    'Simplifier la géométrie de l'élément
                    pTopoOp = CType(pGeometry, ITopologicalOperator2)
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()

                    'Définir la géométrie originale
                    pGeometryOri = pFeature.ShapeCopy
                    'Projeter la géométrie
                    pGeometryOri.Project(pBagErreurs.SpatialReference)

                    Try
                        'Définir les différences de la géométrie
                        pDifference = pTopoOp.SymmetricDifference(pGeometryOri)
                    Catch ex As Exception
                        'Retourner la géométrie originale comme différence
                        pDifference = pGeometryOri
                    End Try

                    'Vérifier si aucune différence
                    If pDifference.IsEmpty Then
                        'Conserver le nombre de sommets
                        pPointColl = CType(pGeometry, IPointCollection)

                        'Conserver le nombre de composantes
                        pGeomColl = CType(pGeometry, IGeometryCollection)

                        'Vérifier si le nombre de sommets ou de composantes est différent
                        If iNbSommets <> pPointColl.PointCount Or iNbComposantes <> pGeomColl.GeometryCount Then
                            'Définir la différence
                            pDifference = pGeometryOri
                        End If
                    End If

                    'Interface contenant la référence spatiale
                    pWrite = CType(pFeature.Class, IFeatureClassWrite)

                    'Vérifier si la géométrie est vide
                    If pGeometry.IsEmpty Then
                        'Définir la différence
                        pDifference = pGeometryOri

                        'Vérifier si on doit corriger
                        If bCorriger Then
                            'Détruire l'élément
                            pWrite.RemoveFeature(pFeature)
                            'Indiquer que l'élément a été détruit
                            pFeature = Nothing
                        End If

                        'Si la géométrie n'est pas vide
                    Else
                        'Vérifier si aucune différence
                        If pDifference.IsEmpty Then
                            'Aucune différence n'est retournée
                            pDifference = Nothing
                        Else
                            'Simplifier la géométrie de l'élément
                            pTopoOp = CType(pGeometry, ITopologicalOperator2)
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()
                        End If

                        'Définir la géométrie de l'élément
                        pFeature.Shape = pGeometry

                        'Vérifier s'il faut corriger
                        If bCorriger Then
                            'Écrire l'élément
                            pWrite.WriteFeature(pFeature)
                        End If
                    End If

                    'Vérifier s'il y a une erreur
                    If pDifference IsNot Nothing Then
                        'Interface pour ajouter une erreur
                        pGeomColl = CType(pBagErreurs, IGeometryCollection)
                        'Ajouter l'erreur dans le Bag
                        pGeomColl.AddGeometry(pDifference)
                    End If
                End If
            End If

        Catch erreur As Exception
            Throw
        Finally
            'Vider la mémoire
            pDifference = Nothing
            pTopoOp = Nothing
            pWrite = Nothing
            pGeometry = Nothing
            pGeometryOri = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions pour corriger la proximité des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la proximité des éléments sélectionnés en utilisant la topologie. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la tolérance utilisée pour corriger la proximité des géométries des éléments sélectionnés.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger la segmentation manquante des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si un fichier d'erreurs doit être créé.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs des éléments.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerProximite(ByVal dPrecision As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean,
                                 ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les erreurs de la géométrie d'un élément.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface utilisé pour ajouter des erreurs dans le Bag.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer un Bag d'erreurs vide
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour ajouter les erreurs dans le Bag
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la proximité ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Corriger la proximité des éléments
                    Call ProximiteTopologieElement(pTopologyGraph, pFeature, bCorriger, pBagErreurs)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                ''Vérifieer si des erreurs sont présente
                'If pGeomColl.GeometryCount > 0 Then
                'Vérifier si on doit corriger
                If bCorriger Then
                    'Terminer l'opération UnDo
                    pEditor.StopOperation("Corriger la proximité")

                    'Enlever la sélection
                    m_MxDocument.FocusMap.ClearSelection()

                    'Sinon
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If

                'Vérifier si on doit créer le fichier d'erreurs
                If bCreerFichierErreurs Then
                    'Initialiser le message d'exécution
                    pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de proximité (NbErr=" & iNbErreurs.ToString & ") ..."
                    'Créer le FeatureLayer des erreurs
                    pFeatureLayer = CreerFeatureLayerErreurs("CorrigerProximite_0", "Erreur de proximité : Précision=" & dPrecision.ToString, _
                                                              m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryMultipoint, pBagErreurs)

                    'Ajouter le FeatureLayer d'erreurs dans la map active
                    m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                End If

                'Rafraîchir l'affichage
                m_MxDocument.ActiveView.Refresh()

                '    'Sinon
                'Else
                '    'Annuler l'opération UnDo
                '    pEditor.AbortOperation()
                'End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeatureLayer = Nothing
            pFeature = Nothing
            pBagErreurs = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pGeomColl = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger la proximité d'un élément selon la topologie des éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger.</param>
    '''<param name="pBagErreurs"> Interface contenant les erreurs de proximité des éléments.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Public Function ProximiteTopologieElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal bCorriger As Boolean,
                                              ByRef pBagErreurs As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie d'un élément.
        Dim pGeometryOri As IGeometry = Nothing         'Interface contenant la géométrie originale d'un élément.
        Dim pDifference As IGeometry = Nothing          'Interface contenant les différences de la géométrie de l'élément.
        Dim pRelOp As IRelationalOperator = Nothing     'Interface pour vérifier les différences des points.
        Dim pWrite As IFeatureClassWrite = Nothing      'Interface utilisé pour écrire un élément.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface qui permet de simplifier la géométrie.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface contenant le nombre de composantes.

        'Par défaut, aucune modification n'a été effectuée
        ProximiteTopologieElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Interface pour corriger les éléments sans affecter la topologie
                pWrite = CType(pFeature.Class, IFeatureClassWrite)

                'Conserver la géométrie originale
                pGeometryOri = pFeature.ShapeCopy
                'Projeter la géométrie originale
                pGeometryOri.Project(pBagErreurs.SpatialReference)

                'Définir la nouvelle géométrie de l'élément
                pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                'Vérifier si la géométrie est invalide
                If pGeometry Is Nothing Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)
                    End If

                    'Indiquer qu'il y a eu une modification
                    ProximiteTopologieElement = True
                    'Interface pour ajouter des erreurs dans le Bag
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)
                    'Ajouter l'erreur dans le Bag
                    pGeomColl.AddGeometry(pGeometryOri)

                    'Si la géométrie est vide
                ElseIf pGeometry.IsEmpty Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)
                    End If

                    'Indiquer qu'il y a eu une modification
                    ProximiteTopologieElement = True
                    'Interface pour ajouter des erreurs dans le Bag
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)
                    'Ajouter l'erreur dans le Bag
                    pGeomColl.AddGeometry(pGeometryOri)

                    'Si la géométrie est valide
                Else
                    'Vérifier si la géométrie n'est pas de type point
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Interface pour vérifier la relation spatiale
                        pRelOp = CType(pGeometry, IRelationalOperator)
                        'Vérifier si le point est différent
                        If pRelOp.Disjoint(pGeometryOri) Then
                            'Interface pour ajouter des erreurs dans le Bag
                            pGeomColl = CType(pBagErreurs, IGeometryCollection)
                            'Ajouter l'erreur dans le Bag
                            pGeomColl.AddGeometry(pGeometryOri)
                        End If

                        'Si la géométrie n'est pas de type point
                    Else
                        'Interface pour simplifier
                        pTopoOp = CType(pGeometryOri, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()

                        'Interface pour simplifier
                        pTopoOp = CType(pGeometry, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()

                        'Interface pour vérifier la relation spatiale
                        pRelOp = CType(pGeometry, IRelationalOperator)

                        'Vérifier s'il y a une différence
                        If Not pRelOp.Equals(pGeometryOri) Then
                            Try
                                'Définir les différences de la géométrie
                                pDifference = pTopoOp.SymmetricDifference(pGeometryOri)
                            Catch ex As Exception
                                'Définir les différences de la géométrie
                                pDifference = pGeometryOri
                            End Try

                            'Indiquer qu'il y a une erreur
                            ProximiteTopologieElement = True
                            'Interface pour ajouter des erreurs dans le Bag
                            pGeomColl = CType(pBagErreurs, IGeometryCollection)
                            'Ajouter l'erreur dans le Bag
                            pGeomColl.AddGeometry(pDifference)
                        End If
                    End If

                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Traiter le Z et le M
                        Call TraiterZ(pGeometry)
                        Call TraiterM(pGeometry)

                        'Corriger la géométrie de l'élément
                        pFeature.Shape = pGeometry

                        'Sauver la correction
                        pWrite.WriteFeature(pFeature)
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pGeometryOri = Nothing
            pDifference = Nothing
            pRelOp = Nothing
            pWrite = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
        End Try
    End Function
#End Region

#Region "Routines et fonctions pour corriger la duplication des éléments de type ligne"
    '''<summary>
    ''' Routine qui permet de corriger la dupplication des lignes pour les éléments de type ligne sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerDupplicationLignes(ByVal dPrecision As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean,
                                          ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.
        Dim iNbErr As Long = -1                             'Contient le nombre d'erreurs.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Traiter tant qu'on trouve des erreurs
                Do While iNbErreurs > iNbErr
                    'Définir le nombre d'erreurs
                    iNbErr = iNbErreurs

                    'Création de la topologie
                    pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                    pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                    'Initialiser la barre de progression
                    pTrackCancel.Progressor.Message = "Correction de la duplication des lignes (NbErr=" & iNbErreurs.ToString & ") ..."
                    InitBarreProgression(0, pEditor.SelectionCount, pTrackCancel)

                    'Interface pour extraire le premier élément de la sélection
                    pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                    'Réinitialiser la recherche des éléments
                    pEnumFeature.Reset()
                    'Extraire le premier élément de la sélection
                    pFeature = pEnumFeature.Next

                    'Traite tous les éléments sélectionnés
                    Do Until pFeature Is Nothing
                        'Corriger la dupplication de l'élément de type ligne
                        Call CorrigerDupplicationLigneElement(dPrecision, bCorriger, pTopologyGraph, pFeature, pBagErreurs)

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Exit Do

                        'Extraire le prochain élément de la sélection
                        pFeature = pEnumFeature.Next
                    Loop

                    'Retourner le nombre d'erreurs
                    iNbErreurs = pGeomColl.GeometryCount
                Loop

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la duplication des lignes")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()
                        'Sélectionner les éléments en erreur
                        'm_MxDocument.FocusMap.SelectByShape(pBagErreurs, Nothing, False)

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de duplication des lignes (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerDuplicationLignes_1", "Erreur de duplication de ligne : Précision=" & dPrecision.ToString,
                                                                 m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de corriger la dupplication de ligne pour un élément de type ligne. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie de l'élément.</param>
    '''<param name="pFeature"> Interface contenant l'élément.</param>
    '''<param name="pBagErreurs"> Contient les géométries d'erreurs.</param>
    ''' 
    Public Sub CorrigerDupplicationLigneElement(ByVal dPrecision As Double, ByVal bCorriger As Boolean, ByVal pTopologyGraph As ITopologyGraph,
                                                ByRef pFeature As IFeature, ByRef pBagErreurs As IGeometryBag)
        'Déclarer les variables de travail
        Dim pNewFeature As IFeature = Nothing               'Interface ESRI contenant un élément en sélection.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie à traiter.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant le super polygone à traiter.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant la polyligne d'erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie. 
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour enlever la partie de la géométrie en erreur.

        Try
            'Vérifier si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si la géométrie de l'élément est de type ligne
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Interface pour ajouter les lignes en erreur
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)

                    'Définir la ligne de la géométrie
                    pPolyline = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID), IPolyline)

                    'Vérifier si la polyligne est invalide
                    If pPolyline IsNot Nothing Then
                        'Interface pour extraire les Edges de l'élément dans la topologie
                        pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)
                        pEnumTopoEdge.Reset()

                        'Extraire la premier Edge
                        pTopoEdge = pEnumTopoEdge.Next

                        'Traiter tous les Edges
                        Do Until pTopoEdge Is Nothing
                            'Vérifier si le Edge a déjà été traité
                            If Not pTopoEdge.Visited Then
                                'Interface pour extraire le nombre d'intersections
                                pEnumTopoParent = pTopoEdge.Parents()

                                'Vérifier si plus d'un élément
                                If pEnumTopoParent.Count > 1 Then
                                    'Définir la ligne en erreur
                                    pPolylineErr = CType(pTopoEdge.Geometry, IPolyline)

                                    'Ajouter la ligne en erreur
                                    pGeomColl.AddGeometry(pTopoEdge.Geometry)

                                    'Interface pour corriger la ligne
                                    pTopoOp = CType(pPolyline, ITopologicalOperator2)
                                    pTopoOp.IsKnownSimple_2 = False
                                    pTopoOp.Simplify()

                                    'Corriger la ligne
                                    pPolyline = CType(pTopoOp.Difference(pPolylineErr), IPolyline)
                                End If

                                'Indiquer que le Edge a été traité
                                pTopoEdge.Visited = True
                            End If

                            'Extraire la premier Edge
                            pTopoEdge = pEnumTopoEdge.Next
                        Loop
                    End If

                    'Vérifier si doit corriger
                    If bCorriger Then
                        'Vérifier si l'élément doit être détruit
                        If pPolyline Is Nothing Then
                            'Détruire l'élément
                            pFeature.Delete()

                        ElseIf pPolyline.IsEmpty Then
                            'Détruire l'élément
                            pFeature.Delete()

                            'Si la géométrie doit être modifiée
                        Else
                            'Interface pour simplifier une géométrie.
                            pTopoOp = CType(pPolyline, ITopologicalOperator2)
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()
                            'Modifié la géométrie
                            pFeature.Shape = pPolyline
                            'Sauver la modification
                            pFeature.Store()
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pNewFeature = Nothing
            pPolyline = Nothing
            pPolylineErr = Nothing
            pGeomColl = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pTopoOp = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions pour corriger la segmentation manquante des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la segmentation manquante des éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour fusionner les géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger la segmentation manquante des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerSegmentationManquante(ByVal dPrecision As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, _
                                             ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les sommets de la segmantation manquante dans la géométrie d'un élément.
        Dim pPointErreur As IMultipoint = Nothing           'Interface contenant les points en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.
        Dim iNbErr As Long = -1                             'Contient le nombre d'erreurs.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                pEnvelope = EnveloppeElementsSelectionner(pEditor)

                'Traiter tant qu'il y a des erreurs
                Do While iNbErreurs > iNbErr
                    'Intialisation du nombre d'erreurs
                    iNbErr = iNbErreurs

                    'Création de la topologie
                    pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                    pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                    'Initialiser la barre de progression
                    pTrackCancel.Progressor.Message = "Correction de la segmentation manquante (NbErr=" & iNbErreurs.ToString & ")..."
                    InitBarreProgression(0, pEditor.SelectionCount, pTrackCancel)

                    'Interface pour extraire le premier élément de la sélection
                    pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                    'Réinitialiser la recherche des éléments
                    pEnumFeature.Reset()
                    'Extraire le premier élément de la sélection
                    pFeature = pEnumFeature.Next

                    'Traite tous les éléments sélectionnés
                    Do Until pFeature Is Nothing
                        'Vérifier si la segmentation manquante de la géométrie de l'élément a été corrigée
                        If CorrigerSegmentationManquanteElement(pTopologyGraph, pFeature, bCorriger, pPointErreur) Then
                            'Ajouter l'erreur dans le Bag
                            pGeomColl.AddGeometry(pPointErreur)
                        End If

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Exit Do

                        'Extraire le prochain élément de la sélection
                        pFeature = pEnumFeature.Next
                    Loop

                    'Retourner le noombre d'erreurs
                    iNbErreurs = pGeomColl.GeometryCount
                Loop

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la segmentation manquante")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de segmentation manquante (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerSegmentationManquante_0", "Segmentation manquante", _
                                                                  m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryMultipoint, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pPointErreur = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger la segmentation manquante d'un élément selon ses éléments en relation à l'aide des outils de topologie. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="bCorriger"> Permet d'indiquer si la correction de la segmentation manquante doit être effectuée.</param>
    '''<param name="pPointErreur"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerSegmentationManquanteElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, bCorriger As Boolean, _
                                                          ByRef pPointErreur As IMultipoint) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie. 
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la ligne de l'élément à écrire.
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface pour créer un nouvel élément.
        Dim pNewFeature As IFeature = Nothing               'Interface contenant un nouvel élément.
        Dim pExtremite As IMultipoint = Nothing             'Interface contenant les extrémités de la lignes à segmenter.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les points en erreur.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour extraire les points en erreurs.
        Dim pClosedPoint As IPoint = Nothing

        'Par défaut, aucune modification n'a été effectuée
        CorrigerSegmentationManquanteElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si l'élément est une ligne
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Créer le Bag vide des lignes en erreur
                    pPointErreur = New Multipoint
                    pPointErreur.SpatialReference = pFeature.Shape.SpatialReference
                    'Interface pour ajouter les points en erreur
                    pPointColl = CType(pPointErreur, IPointCollection)

                    'Interface pour extraire les Edges de l'élément dans la topologie
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)

                    'Vérifier s'il y a une erreur de la segmentation manquante
                    If pEnumTopoEdge.Count <> 1 Then
                        'Indiquer qu'il y a eu une modification
                        CorrigerSegmentationManquanteElement = True

                        'Définir l'élément à écrire
                        pNewFeature = pFeature

                        'Interface pour corriger les éléments
                        pFeatureClass = CType(pFeature.Class, IFeatureClass)
                        pWrite = CType(pFeature.Class, IFeatureClassWrite)

                        'Vérifier si la géométrie est invalide
                        If pEnumTopoEdge.Count = 0 Then
                            'Détruire l'élément
                            If bCorriger Then pWrite.RemoveFeature(pFeature)

                            'Si la géométrie est valide
                        Else
                            'Extraire la premier Edge
                            pTopoEdge = pEnumTopoEdge.Next

                            'Traiter tous les Edges
                            Do Until pTopoEdge Is Nothing
                                'Définir la géométrie du Edge
                                pGeometry = pTopoEdge.Geometry

                                'Définir la ligne de l'élément
                                pPolyline = CType(pGeometry, IPolyline)

                                If pPolyline.IsClosed Then
                                    'Conserver le point fermé
                                    pClosedPoint = pPolyline.FromPoint
                                    'Ajouter le point en erreur
                                    pPointColl.AddPoint(pClosedPoint)
                                Else
                                    'Ajouter les points en erreur
                                    pPointColl.AddPoint(pPolyline.FromPoint)
                                    pPointColl.AddPoint(pPolyline.ToPoint)
                                End If

                                'Vérifier si on doir corriger
                                If bCorriger Then
                                    'Vérifier si l'élément est valide
                                    If pNewFeature Is Nothing Then
                                        'Créer un nouvel élément
                                        pNewFeature = pFeatureClass.CreateFeature
                                        'Copier les valeurs d'attributs
                                        Call fbCopierValeurAttributElementIdentique(pNewFeature, pFeature)
                                    End If

                                    'Traiter le Z et le M
                                    Call TraiterZ(pGeometry)
                                    Call TraiterM(pGeometry)

                                    'Corriger la géométrie de l'élément
                                    pNewFeature.Shape = pGeometry

                                    'Sauver la correction
                                    pWrite.WriteFeature(pNewFeature)

                                    'Définir que l'élément a été traité
                                    pNewFeature = Nothing
                                End If

                                'Extraire les Edges suivants
                                pTopoEdge = pEnumTopoEdge.Next
                            Loop
                        End If

                        'Extraire la géométrie de l'élément de la topologie
                        pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)
                        'Interface pour extraire les extrémités de la géométrie de l'élément de la topologie
                        pTopoOp = CType(pGeometry, ITopologicalOperator2)
                        pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        'Simplifier les points en erreur
                        pTopoOp.Simplify()
                        'Extraire les extrémités de la géométrie de l'élément de la topologie
                        pExtremite = CType(pTopoOp, IMultipoint)

                        'Interface pour simplifier les points en erreur
                        pTopoOp = CType(pPointErreur, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        'Simplifier les points en erreur
                        pTopoOp.Simplify()

                        'Extraire les points en erreur
                        pPointErreur = CType(pTopoOp.Difference(pExtremite), IMultipoint)

                        'Si la géométrie est vide
                        If pPointErreur.IsEmpty Then
                            'Vérifier si un point fermé a été trouvé
                            If pClosedPoint IsNot Nothing Then
                                'Interface pour ajouter les points en erreur
                                pPointColl = CType(pPointErreur, IPointCollection)
                                'Ajouter le point en erreur
                                pPointColl.AddPoint(pClosedPoint)
                            Else
                                'Les points en erreur sont les extrémités
                                pPointErreur = pExtremite
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pWrite = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pPolyline = Nothing
            pFeatureClass = Nothing
            pNewFeature = Nothing
            pExtremite = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
            pClosedPoint = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de copier les valeurs d'attributs d'un élément à un autre. 
    ''' Les attributs OBJECTID, SHAPE, SHAPE_Length et SHAPE_Area ne sont pas traité. 
    ''' Le traitement n'est pas sauvé à la fin. 
    ''' Il faut effectuer un store à la suite de ce traitement si on veut que les valeurs d'attributs soient sauvées.
    '''</summary>
    '''
    '''<param name=" pFeatureTo "> Interface ESRI de l'élément avec les valeurs de destination.</param>
    '''<param name=" pFeatureFrom "> Interface ESRI de l'élément avec les valeurs d'origine.</param>
    '''
    '''<returns> La fonction va retourner un "Boolean". Si le traitement n'a pas réussi, le "Boolean" sera à "False".</returns>
    ''' 
    Public Function fbCopierValeurAttributElementIdentique(ByRef pFeatureTo As IFeature, ByVal pFeatureFrom As IFeature) As Boolean
        'Déclarer les variable des travail
        Dim pFields As IFields = Nothing       'Interface ESRI contenant les attributs d'un élément.
        Dim pField As IField = Nothing         'Interface ESRI contenat un attribut d'élément.
        Dim nPosition As Integer = Nothing     'Position de l'attribut recherché.

        Try
            'Interface pour trouver le nom de tous les attributs
            pFields = pFeatureTo.Fields

            'Traiter tous les attributs de l'élément
            For i = 0 To pFields.FieldCount - 1

                'Interface pour trouvé le nom des attributs
                pField = pFields.Field(i)

                'Vérifier si l'attribut est un attribut réservé
                If pField.Editable And pField.Type <> esriFieldType.esriFieldTypeGeometry Then
                    'Trouver la position de l'attribut correspondante
                    nPosition = pFeatureFrom.Fields.FindField(pField.Name)

                    'Vérifier la position a été trouvée
                    If nPosition >= 0 Then
                        'Copier la valeur d'attribut
                        pFeatureTo.Value(i) = pFeatureFrom.Value(nPosition)
                    End If
                End If
            Next i

            'Retourner le résultat
            fbCopierValeurAttributElementIdentique = True

        Catch erreur As Exception
            Throw erreur
        Finally
            'Vider la mémoire
            pFields = Nothing
            pField = Nothing
        End Try
    End Function
#End Region

#Region "Routines et fonctions pour corriger la segmentation en trop des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la segmentation en trop des éléments sélectionnés. 
    ''' Les éléments dont la segmentation en trop est présente seront fusionnés avec leurs éléments adjacents de même classe et même valeur d'attribut.
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour fusionner les géométries des éléments.</param>
    '''<param name="sListeAttributs"> Contient les noms des attributs dont les valeurs doivent être identiques pour enlever la segmentation.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger la segmentation manquante des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerSegmentationEnTrop(ByVal dPrecision As Double, ByVal sListeAttributs As String, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, _
                                          ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface ESRI contenant la classe d'un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les géométries des erreurs de segmentation en trop.
        Dim pErreur As IGeometry = Nothing                  'Interface contenant les géométries en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.
        Dim pDictLiens As New Dictionary(Of String, Integer) 'Dictionnaire contenant les liens entre les éléments traités.
        Dim pDictElements As New Dictionary(Of Integer, Integer) 'Dictionnaire contenant les éléments traités.
        Dim pPair As KeyValuePair(Of Integer, Integer)      'Contient les valeur d'un item du dictionnaire.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                pEnvelope = EnveloppeElementsSelectionner(pEditor)

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la segmentation en trop ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Vérifier si un élément est présent
                If pFeature IsNot Nothing Then
                    'Définir la classe de traitement
                    pFeatureClass = CType(pFeature.Class, IFeatureClass)
                End If

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Vérifier si la classe correspond
                    If pFeatureClass.Equals(pFeature.Class) Then
                        'Traiter la segmentation en trop de la géométrie de l'élément
                        Call CorrigerSegmentationEnTropElement(pTopologyGraph, pFeature, sListeAttributs, bCorriger, pDictLiens, pDictElements, pBagErreurs)
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le noombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Initialiser le message d'exécution
                    pTrackCancel.Progressor.Message = "Destruction des éléments fusionnés ..."

                    'Traiter tous les éléments à détruire
                    For Each pPair In pDictElements
                        'Vérifier si l'élément doit être détruit
                        If pPair.Key <> pPair.Value Then
                            'Extraire l'élément à détruire
                            pFeature = pFeatureClass.GetFeature(pPair.Key)

                            'Détruire l'élément
                            pFeature.Delete()
                        End If
                    Next

                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la segmentation en trop")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de segmentation en trop (NbErr=" & iNbErreurs.ToString & ") ..."

                        'Vérifier si la classe est de type polyline
                        If pFeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                            'Créer le FeatureLayer des erreurs
                            pFeatureLayer = CreerFeatureLayerErreurs("CorrigerSegmentationEnTrop_0", "Segmentation en trop", _
                                                                      m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPoint, pBagErreurs)

                            'Ajouter le FeatureLayer d'erreurs dans la map active
                            m_MxDocument.FocusMap.AddLayer(pFeatureLayer)

                            'Si la classe est de type polygon
                        ElseIf pFeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                            'Créer le FeatureLayer des erreurs
                            pFeatureLayer = CreerFeatureLayerErreurs("CorrigerSegmentationEnTrop_1", "Segmentation en trop", _
                                                                      m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                            'Ajouter le FeatureLayer d'erreurs dans la map active
                            m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                        End If
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pErreur = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
            pDictLiens = Nothing
            pDictElements = Nothing
            pPair = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger la segmentation en trop d'un élément selon ses éléments en relation à l'aide des outils de topologie. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="sListeAttributs"> Liste des noms d'attributs à comparer.</param>
    '''<param name="bCorriger"> Permet d'indiquer si la correction de la segmentation manquante doit être effectuée.</param>
    '''<param name="pDictLiens"> Interface contenant le dictionnaire des liens entre les éléments traités.</param>
    '''<param name="pDictElements"> Interface contenant le dictionnaire des éléments traités.</param>
    '''<param name="pBagErreurs"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerSegmentationEnTropElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal sListeAttributs As String, bCorriger As Boolean, _
                                                       ByRef pDictLiens As Dictionary(Of String, Integer), ByRef pDictElements As Dictionary(Of Integer, Integer), _
                                                       ByRef pBagErreurs As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pGeometryAdj As IGeometry = Nothing             'Interface contenant la géométrie de l'élément adjacent.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément.
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface contenant la classe d'un élément
        Dim pFeatureTmp As IFeature = Nothing               'Interface contenant un élément temporaire.
        Dim pFeatureAdj As IFeature = Nothing               'Interface contenant l'élément adjacent.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utiliser pour fusionner les géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes en erreur.
        Dim sLien As String = ""                            'Contient la valeur d'un lien.
        Dim sMessage As String = ""                         'Contient les différences d'attribut.
        Dim sListeOid As String = "/"                       'Contient la liste des OID fusionnés.
        Dim iOid As Integer = 0                             'Contient le ObjectId précédent.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerSegmentationEnTropElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Debug.Print("-")
                'Debug.Print("-" & pFeature.OID.ToString)
                'Conserver le Oid précédent
                iOid = pFeature.OID
                sListeOid = "/" & iOid.ToString & "/"
                pFeatureTmp = pFeature

                'Interface pour ajouter les lignes en erreur
                pGeomColl = CType(pBagErreurs, IGeometryCollection)

                'Vérifier si le cas a déjà été traité
                If pDictElements.ContainsKey(iOid) Then
                    'Extraire la FeatureClass
                    pFeatureClass = CType(pFeature.Class, IFeatureClass)
                    'Extraire l'élément
                    pFeatureTmp = pFeatureClass.GetFeature(pDictElements.Item(iOid))

                    'Traiter tant que le OID est différent de l'élément
                    Do While iOid <> pFeatureTmp.OID
                        'Debug.Print("-" & pFeatureTmp.OID.ToString)

                        'Vérifier si l'élément
                        If sListeOid.Contains("/" & pDictElements.Item(iOid).ToString & "/") Then
                            'Debug.Print(sListeOid)
                            'Extraire l'élément
                            pFeatureAdj = pFeatureClass.GetFeature(iOid)
                            'Extraire la géométrie
                            pGeometry = pFeatureTmp.ShapeCopy
                            pGeometry.Project(pBagErreurs.SpatialReference)

                            'Interface pour fusionner les géométries
                            pTopoOp = CType(pGeometry, ITopologicalOperator2)
                            'Définir la géométrie de l'élément adjacent
                            pGeometryAdj = pFeatureAdj.ShapeCopy
                            pGeometryAdj.Project(pBagErreurs.SpatialReference)
                            'Fusionner les géométries
                            pGeometry = pTopoOp.Union(pGeometryAdj)

                            'Redéfinir la géométrie de l'élément
                            pFeatureTmp.Shape = pGeometry

                            'Indiquer l'élément à conserver
                            pDictElements.Item(pFeatureTmp.OID) = pFeatureTmp.OID
                            'Définir le dernier OID à traiter
                            iOid = pFeatureTmp.OID

                        Else
                            'Conserver le Oid précédent
                            iOid = pFeatureTmp.OID
                            'Définir la liste des oids
                            sListeOid = sListeOid & iOid.ToString & "/"
                            'Extraire l'élément
                            pFeatureTmp = pFeatureClass.GetFeature(pDictElements.Item(iOid))
                        End If
                    Loop
                End If

                'Définir la géométrie de l'élément
                pGeometry = pFeatureTmp.ShapeCopy
                pGeometry.Project(pBagErreurs.SpatialReference)

                'Vérifier si l'élément est une ligne
                If pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Corriger la segmentation en trop pour les lignes
                    CorrigerSegmentationEnTropElement = CorrigerSegmentationEnTropLigne(pTopologyGraph, pFeature, sListeAttributs, _
                                                                                        pFeatureTmp, pGeometry, pDictLiens, pDictElements, pBagErreurs)

                    'Vérifier si l'élément est une surface
                ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Corriger la segmentation en trop pour les surface
                    CorrigerSegmentationEnTropElement = CorrigerSegmentationEnTropSurface(pTopologyGraph, pFeature, sListeAttributs, _
                                                                                          pFeatureTmp, pGeometry, pDictLiens, pDictElements, pBagErreurs)
                End If

                'Vérifier si on doit corriger et qu'une erreur est présente
                If bCorriger And CorrigerSegmentationEnTropElement Then
                    'Vérifier si l'élément n'a pas été traité
                    If Not pDictElements.ContainsKey(pFeature.OID) Then
                        'Ajouter l'élément dans le dictionnaire
                        pDictElements.Add(pFeature.OID, pFeature.OID)
                    End If

                    'Interface pour corriger les éléments
                    pFeatureClass = CType(pFeature.Class, IFeatureClass)
                    pWrite = CType(pFeature.Class, IFeatureClassWrite)

                    'Traiter le Z et le M
                    Call TraiterZ(pGeometry)
                    Call TraiterM(pGeometry)

                    'Corriger la géométrie de l'élément
                    pFeatureTmp.Shape = pGeometry

                    'Sauver la correction
                    pWrite.WriteFeature(pFeatureTmp)

                    'Debug.Print(">" & pFeatureTmp.OID.ToString)

                    'Si pas de correction
                Else
                    'Par défaut, aucune modification n'a été effectuée
                    CorrigerSegmentationEnTropElement = False
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pWrite = Nothing
            pFeatureClass = Nothing
            pFeatureTmp = Nothing
            pFeatureAdj = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de corriger la segmentation en trop d'un élément de type ligne selon ses éléments en relation à l'aide des outils de topologie. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="sListeAttributs"> Liste des noms d'attributs à comparer.</param>
    '''<param name="pFeatureTmp"> Contient l'élément temporaire à traiter.</param>
    '''<param name="pGeometry"> Contient la géométrie de l'élément temporaire corrigée.</param>
    '''<param name="pDictLiens"> Interface contenant le dictionnaire des liens entre les éléments traités.</param>
    '''<param name="pDictElements"> Interface contenant le dictionnaire des éléments traités.</param>
    '''<param name="pBagErreurs"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerSegmentationEnTropLigne(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal sListeAttributs As String, _
                                                     ByRef pFeatureTmp As IFeature, ByRef pGeometry As IGeometry, ByRef pDictLiens As Dictionary(Of String, Integer), _
                                                     ByRef pDictElements As Dictionary(Of Integer, Integer), ByRef pBagErreurs As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pGeometryAdj As IGeometry = Nothing             'Interface contenant la géométrie de l'élément adjacent.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface utilisé pour extraire les nodes de la topologie.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node de la topologie. 
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie. 
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface contenant la classe d'un élément
        Dim pFeatureAdj As IFeature = Nothing               'Interface contenant l'élément adjacent.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utiliser pour fusionner les géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes en erreur.
        Dim pPolyline As IPolyline = Nothing                'Interface pour vérifier si la ligne est fermée.
        Dim sLien As String = ""                            'Contient la valeur d'un lien.
        Dim sMessage As String = ""                         'Contient les différences d'attribut.
        Dim iOid As Integer = 0                             'Contient le ObjectId précédent.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerSegmentationEnTropLigne = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si l'élément est une ligne
                If pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Interface pour vérifier si la ligne est fermée
                    pPolyline = CType(pGeometry, IPolyline)
                    pPolyline.Project(pBagErreurs.SpatialReference)

                    'Vérifier si la ligne est fermée
                    If Not pPolyline.IsClosed Then
                        'Interface pour ajouter les lignes en erreur
                        pGeomColl = CType(pBagErreurs, IGeometryCollection)

                        'Interface pour extraire les composantes
                        pEnumTopoNode = pTopologyGraph.GetParentNodes(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                        'Extraire la première composante
                        pTopoNode = pEnumTopoNode.Next

                        'Traiter toutes les composantes
                        Do Until pTopoNode Is Nothing
                            'Interface pour extraire le nombre d'intersections
                            pEnumTopoParent = pTopoNode.Parents()

                            'Si le nombre d'éléments dans le Edge est 2, on peut peut-être fusionner
                            If pEnumTopoParent.Count = 2 Then
                                'Initialiser l'extraction
                                pEnumTopoParent.Reset()
                                'Extraire le premier élément
                                pEsriTopoParent = pEnumTopoParent.Next()
                                'Extraire la FeatureClass
                                pFeatureClass = pEsriTopoParent.m_pFC
                                'Extraire l'élément
                                pFeatureAdj = pFeatureClass.GetFeature(pEsriTopoParent.m_FID)
                                'Vérifier si l'élément adjacent est l'élément traité
                                If pFeature.OID = pFeatureAdj.OID And pFeature.Class.AliasName = pFeatureAdj.Class.AliasName Then
                                    'Extraire le prochain élément
                                    pEsriTopoParent = pEnumTopoParent.Next()
                                    'Extraire la FeatureClass
                                    pFeatureClass = pEsriTopoParent.m_pFC
                                    'Extraire l'élément
                                    pFeatureAdj = pFeatureClass.GetFeature(pEsriTopoParent.m_FID)
                                End If
                                'Debug.Print("!" & pFeature.OID.ToString & "-" & pFeatureAdj.OID.ToString)

                                'Vérifier si la classe est la même entre les deux éléments
                                If pFeature.Class.AliasName = pFeatureAdj.Class.AliasName Then
                                    'Interface pour vérifier si la ligne est fermée
                                    pPolyline = CType(pFeatureAdj.Shape, IPolyline)
                                    pPolyline.Project(pBagErreurs.SpatialReference)

                                    'Vérifier si la ligne est fermée
                                    If Not pPolyline.IsClosed Then
                                        'Vérifier si le oid de l'élément est plus petit 
                                        If pFeature.OID < pFeatureAdj.OID Then
                                            'Définir la valeur du lien
                                            sLien = pFeature.OID.ToString & "_" & pFeatureAdj.OID.ToString
                                            'S'il est plus grand
                                        Else
                                            'Définir la valeur du lien
                                            sLien = pFeatureAdj.OID.ToString & "_" & pFeature.OID.ToString
                                        End If

                                        'Vérifier si le cas a déjà été traité
                                        If Not (pDictLiens.ContainsKey(sLien)) Then
                                            'Ajouter le lien de l'élément dans le dictionnaire
                                            pDictLiens.Add(sLien, pFeatureTmp.OID)

                                            'Comparer les valeurs d'attribut entre les 2 éléments
                                            sMessage = ComparerAttributElement(pFeature, pFeatureAdj, sListeAttributs & "/")

                                            'Vérifier s'il y a une différence
                                            If sMessage.Length = 0 Then
                                                'Vérifier si le cas a déjà été traité
                                                If pDictElements.ContainsKey(pFeatureAdj.OID) Then
                                                    'Extraire la FeatureClass
                                                    pFeatureClass = CType(pFeature.Class, IFeatureClass)
                                                    'Trouver l'élément non-détruit correspondant
                                                    Do
                                                        'Debug.Print("=" & pFeatureAdj.OID.ToString)
                                                        'Conserver le Oid précédent
                                                        iOid = pFeatureAdj.OID
                                                        'Extraire l'élément
                                                        pFeatureAdj = pFeatureClass.GetFeature(pDictElements.Item(pFeatureAdj.OID))
                                                        'Traiter tant que le OID est différent de l'élément
                                                    Loop While iOid <> pFeatureAdj.OID

                                                    'Vérifier si l'élément n'est pas un élément à détruire
                                                    If pDictElements.Item(pFeatureAdj.OID) = pFeatureAdj.OID Then
                                                        'Indiquer que l'élément est à détruire
                                                        pDictElements.Item(pFeatureAdj.OID) = pFeature.OID
                                                    End If

                                                    'si le cas n'a pas été traité
                                                Else
                                                    'Ajouter l'élément dans le dictionnaire
                                                    pDictElements.Add(pFeatureAdj.OID, pFeatureTmp.OID)
                                                End If

                                                'Indiquer qu'il y a eu une modification
                                                CorrigerSegmentationEnTropLigne = True

                                                'Ajouter le point en erreur
                                                pGeomColl.AddGeometry(pTopoNode.Geometry)

                                                'Interface pour fusionner les géométries
                                                pTopoOp = CType(pGeometry, ITopologicalOperator2)

                                                'Définir la géométrie de l'élément adjacent
                                                pGeometryAdj = pFeatureAdj.ShapeCopy
                                                pGeometryAdj.Project(pBagErreurs.SpatialReference)

                                                'Fusionner les géométries
                                                pGeometry = pTopoOp.Union(pGeometryAdj)

                                                'Debug.Print("+" & pFeatureTmp.OID.ToString & "-" & pFeatureAdj.OID.ToString)
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            'Extraire la prochaine composante
                            pTopoNode = pEnumTopoNode.Next
                        Loop
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pEnumTopoParent = Nothing
            pEsriTopoParent = Nothing
            pFeatureClass = Nothing
            pFeatureAdj = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pPolyline = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de corriger la segmentation en trop d'un élément de type surface selon ses éléments en relation à l'aide des outils de topologie. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="sListeAttributs"> Liste des noms d'attributs à comparer.</param>
    '''<param name="pFeatureTmp"> Contient l'élément temporaire à traiter.</param>
    '''<param name="pGeometry"> Contient la géométrie de l'élément temporaire corrigée.</param>
    '''<param name="pDictLiens"> Interface contenant le dictionnaire des liens entre les éléments traités.</param>
    '''<param name="pDictElements"> Interface contenant le dictionnaire des éléments traités.</param>
    '''<param name="pBagErreurs"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerSegmentationEnTropSurface(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal sListeAttributs As String, _
                                                       ByRef pFeatureTmp As IFeature, ByRef pGeometry As IGeometry, ByRef pDictLiens As Dictionary(Of String, Integer), _
                                                       ByRef pDictElements As Dictionary(Of Integer, Integer), ByRef pBagErreurs As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pGeometryAdj As IGeometry = Nothing             'Interface contenant la géométrie de l'élément adjacent.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface utilisé pour extraire les nodes de la topologie.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un Node de la topologie. 
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie. 
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface contenant les parents du EDGE traité.
        Dim pEsriTopoParent As esriTopologyParent = Nothing 'Interface contenant la structure d'information du parent sélectionné.
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface contenant la classe d'un élément
        Dim pFeatureAdj As IFeature = Nothing               'Interface contenant l'élément adjacent.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utiliser pour fusionner les géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes en erreur.
        Dim sLien As String = ""                            'Contient la valeur d'un lien.
        Dim sMessage As String = ""                         'Contient les différences d'attribut.
        Dim iOid As Integer = 0                             'Contient le ObjectId précédent.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerSegmentationEnTropSurface = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                If pFeature.OID = 536 Then
                    Debug.Print(pFeature.OID.ToString)
                End If
                'Définir la géométrie de l'élément
                pGeometry = pFeatureTmp.ShapeCopy
                pGeometry.Project(pBagErreurs.SpatialReference)

                'Vérifier si l'élément est une ligne
                If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Interface pour ajouter les lignes en erreur
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)

                    'Interface pour extraire les composantes
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, ESRI.ArcGIS.Geodatabase.IFeatureClass), pFeature.OID)

                    'Extraire la première composante
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter toutes les composantes
                    Do Until pTopoEdge Is Nothing
                        'Interface pour extraire le nombre d'intersections
                        pEnumTopoParent = pTopoEdge.Parents()

                        'Si le nombre d'éléments dans le Edge est 2, on peut peut-être fusionner
                        If pEnumTopoParent.Count = 2 Then
                            'Initialiser l'extraction
                            pEnumTopoParent.Reset()
                            'Extraire le premier élément
                            pEsriTopoParent = pEnumTopoParent.Next()
                            'Extraire la FeatureClass
                            pFeatureClass = pEsriTopoParent.m_pFC
                            'Extraire l'élément
                            pFeatureAdj = pFeatureClass.GetFeature(pEsriTopoParent.m_FID)
                            'Vérifier si l'élément adjacent est l'élément traité
                            If pFeature.OID = pFeatureAdj.OID And pFeature.Class.AliasName = pFeatureAdj.Class.AliasName Then
                                'Extraire le prochain élément
                                pEsriTopoParent = pEnumTopoParent.Next()
                                'Extraire la FeatureClass
                                pFeatureClass = pEsriTopoParent.m_pFC
                                'Extraire l'élément
                                pFeatureAdj = pFeatureClass.GetFeature(pEsriTopoParent.m_FID)
                            End If
                            'Debug.Print("!" & pFeature.OID.ToString & "-" & pFeatureAdj.OID.ToString)

                            'Vérifier si la classe est la même entre les deux éléments
                            If pFeature.Class.AliasName = pFeatureAdj.Class.AliasName Then
                                'Vérifier si le oid de l'élément est plus petit 
                                If pFeature.OID < pFeatureAdj.OID Then
                                    'Définir la valeur du lien
                                    sLien = pFeature.OID.ToString & "_" & pFeatureAdj.OID.ToString
                                    'S'il est plus grand
                                Else
                                    'Définir la valeur du lien
                                    sLien = pFeatureAdj.OID.ToString & "_" & pFeature.OID.ToString
                                End If

                                'Vérifier si le cas a déjà été traité
                                If Not (pDictLiens.ContainsKey(sLien)) Then
                                    'Ajouter le lien de l'élément dans le dictionnaire
                                    pDictLiens.Add(sLien, pFeatureTmp.OID)

                                    'Comparer les valeurs d'attribut entre les 2 éléments
                                    sMessage = ComparerAttributElement(pFeature, pFeatureAdj, sListeAttributs & "/")

                                    'Vérifier s'il y a une différence
                                    If sMessage.Length = 0 Then
                                        'Vérifier si le cas a déjà été traité
                                        If pDictElements.ContainsKey(pFeatureAdj.OID) Then
                                            'Trouver l'élément non-détruit correspondant
                                            Do
                                                'Debug.Print("=" & pFeatureAdj.OID.ToString)
                                                'Conserver le Oid précédent
                                                iOid = pFeatureAdj.OID
                                                'Extraire la FeatureClass
                                                pFeatureClass = CType(pFeature.Class, IFeatureClass)
                                                'Extraire l'élément
                                                pFeatureAdj = pFeatureClass.GetFeature(pDictElements.Item(pFeatureAdj.OID))

                                                'Vérifier si l'élément adjacent est le même que l'élément traité
                                                If pDictElements.Item(pFeatureAdj.OID) = pFeature.OID Then
                                                    'Définir le OID pour arrêter la boucle
                                                    iOid = pFeatureAdj.OID
                                                End If

                                                'Traiter tant que le OID est différent de l'élément
                                            Loop While iOid <> pFeatureAdj.OID

                                            'Vérifier si l'élément n'est pas un élément à détruire
                                            If pDictElements.Item(pFeatureAdj.OID) = pFeatureAdj.OID Then
                                                'Indiquer que l'élément est à détruire
                                                pDictElements.Item(pFeatureAdj.OID) = pFeature.OID
                                            End If

                                            'si le cas n'a pas été traité
                                        Else
                                            'Ajouter l'élément dans le dictionnaire
                                            pDictElements.Add(pFeatureAdj.OID, pFeatureTmp.OID)
                                        End If

                                        'Indiquer qu'il y a eu une modification
                                        CorrigerSegmentationEnTropSurface = True

                                        'Ajouter la ligne en erreur
                                        pGeomColl.AddGeometry(pTopoEdge.Geometry)

                                        'Interface pour fusionner les géométries
                                        pTopoOp = CType(pGeometry, ITopologicalOperator2)

                                        'Définir la géométrie de l'élément adjacent
                                        pGeometryAdj = pFeatureAdj.ShapeCopy
                                        pGeometryAdj.Project(pBagErreurs.SpatialReference)

                                        'Fusionner les géométries
                                        pGeometry = pTopoOp.Union(pGeometryAdj)

                                        'Debug.Print("+" & pFeatureTmp.OID.ToString & "-" & pFeatureAdj.OID.ToString)
                                    End If
                                End If
                            End If
                        End If

                        'Extraire la prochaine composante
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometryAdj = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pEnumTopoParent = Nothing
            pEsriTopoParent = Nothing
            pFeatureClass = Nothing
            pFeatureAdj = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de comparer tous les attributs de la liste entre deux éléments.
    ''' 
    ''' Les attributs non étitables et les géométries sont toujours excluts de la comparaison.
    '''</summary>
    ''' 
    '''<param name="pFeature"> Interface contenant l'élément à traiter.</param>
    '''<param name="pFeatureRel"> Interface contenant l'élément à comparer.</param>
    '''<param name="sListeAttribut"> Liste des noms d'attributs à comparer.</param>
    ''' 
    ''' <return> String contenant les différences entre les 2 éléments.</return>
    '''
    Private Function ComparerAttributElement(ByVal pFeature As IFeature, ByVal pFeatureRel As IFeature, ByVal sListeAttribut As String) As String
        'Déclarer les variables de travail
        Dim pFields As IFields = Nothing        'Interface contenant les attribut des éléments à traiter.
        Dim pFieldsRel As IFields = Nothing     'Interface contenant les attribut des éléments à comparer.
        Dim iPosAtt As Integer = -1             'Contient la position de l'attribut en relation.

        'Définir la valeur par défaut
        ComparerAttributElement = ""

        Try
            'Interface pour traiter les attributs
            pFields = pFeature.Fields
            pFieldsRel = pFeatureRel.Fields

            'Traiter tous les attributs
            For i = 0 To pFields.FieldCount - 1
                'Vérifier si l'attribut est à comparer
                If sListeAttribut.Contains(pFields.Field(i).Name & "/") Then
                    'Vérifier si l'attribut est éditable et n'est pas de type "Geometry"
                    If pFields.Field(i).Editable And Not pFields.Field(i).Type = esriFieldType.esriFieldTypeGeometry Then
                        'Définir la position de l'attribut de l'élément en relation
                        iPosAtt = pFeatureRel.Fields.FindField(pFields.Field(i).Name)

                        'Vérifier si l'attribut est présent de la FeatureClass en relation
                        If iPosAtt >= 0 Then
                            'Vérifier si la valeur de l'attribut est différent
                            If pFeature.Value(i).ToString() <> pFeatureRel.Value(iPosAtt).ToString() Then
                                'Ajouter la différence
                                ComparerAttributElement = ComparerAttributElement & "#" & pFields.Field(i).Name & "=" & pFeature.Value(i).ToString _
                                                                                  & "/" & pFeatureRel.Value(iPosAtt).ToString
                                'Sortir de la fonction
                                Exit Function
                            End If

                            'si l'attribut est absent de la FeatureClass en relation
                        Else
                            'Ajouter la différence
                            ComparerAttributElement = ComparerAttributElement & "#" & pFields.Field(i).Name & "/Aucun attribut correspond"
                            'Sortir de la fonction
                            Exit Function
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pFields = Nothing
            pFieldsRel = Nothing
        End Try
    End Function
#End Region

#Region "Routines et fonctions pour corriger le filtrage des sommets des éléments"
    '''<summary>
    ''' Routine qui permet de corriger le filtrage des éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour fusionner les géométries des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale utilisée pour filtrer les sommets en trop des géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerFiltrage(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, _
                                ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les points en erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les points en erreur dans la géométrie d'un élément.
        Dim bModif As Boolean = False                       'Indique s'il y a eu une modification.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer un Bag d'erreurs vide
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire les points en erreur
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                pEnvelope = EnveloppeElementsSelectionner(pEditor)

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la distance latérale des sommets (Distance Latérale = " & dDistLat.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Vérifier si la géométrie de l'élément a été filtrée selon une distance latérale
                    If CorrigerFiltrageElement(pTopologyGraph, pFeature, dDistLat, bCorriger, pBagErreurs) Then
                        'Indiquer qu'une modification a été effectuée
                        bModif = True
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Interface pour extraire les points en erreur
                pGeomColl = CType(pBagErreurs, IGeometryCollection)

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier la présecense d'une erreur et si on doit créer le fichier d'erreurs
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier la présecense d'une modification et si on doit corriger
                    If bModif And bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la distance latérale des sommets")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier la présecense d'une erreur et si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de distance latérale des sommets (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerFiltrage_0", "Erreur de distance latérale des sommets  : " & dDistLat.ToString, _
                                                                  m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryMultipoint, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pGeomColl = Nothing
            pBagErreurs = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger le filtrage latérale d'un élément selon Douglass-Peuker 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion entre les éléments.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="dDistLat"> Contient la distance latérale utilisée pour filtrer les sommets en trop des géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les éléments.</param>
    '''<param name="pBagErreurs"> Interface contenant les points en erreur.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Public Function CorrigerFiltrageElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal dDistLat As Double, _
                                            ByVal bCorriger As Boolean, ByRef pBagErreurs As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément.
        Dim pGeometryOri As IGeometry = Nothing             'Interface contenant la géométrie originale d'un élément.
        Dim pDifference As IGeometry = Nothing              'Interface contenant les différences.
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un edge de la topologie. 
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour ajouter les points en erreur d'une géométrie.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne. 
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire le nombre de sommets.
        Dim pMultipointAvant As IMultipoint = Nothing       'Interface contenant les points de la géométrie de l'élément avant le filtrage.
        Dim pMultipointApres As IMultipoint = Nothing       'Interface contenant les points de la géométrie de l'élément après le filtrage.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire les points en erreur.
        Dim iNbPoints As Integer = 0                        'Contient le nombre de sommets.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerFiltrageElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Interface pour corriger les éléments sans affecter la topologie
                pWrite = CType(pFeature.Class, IFeatureClassWrite)

                'Définir la géométrie
                pGeometryOri = pFeature.ShapeCopy
                'Projeter la géométrie
                pGeometryOri.Project(pBagErreurs.SpatialReference)

                'Définir la nouvelle géométrie de l'élément
                pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                'Vérifier si la géométrie est invalide
                If pGeometry Is Nothing Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)
                    End If

                    'Indiquer qu'il y a eu une modification
                    CorrigerFiltrageElement = True
                    'Interface pour ajouter l'erreur dans le Bag
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)
                    'Ajouter l'erreur dans le Bag
                    pGeomColl.AddGeometry(pGeometryOri)

                    'Si la géométrie est vide
                ElseIf pGeometry.IsEmpty Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)
                    End If

                    'Indiquer qu'il y a eu une modification
                    CorrigerFiltrageElement = True
                    'Interface pour ajouter l'erreur dans le Bag
                    pGeomColl = CType(pBagErreurs, IGeometryCollection)
                    'Ajouter l'erreur dans le Bag
                    pGeomColl.AddGeometry(pGeometryOri)

                    'Si la géométrie est valide
                Else
                    'Vérifier si la géométrie est de type polyline ou polygon
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Or pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                        'Interface pour extraire le nombre de sommets
                        pPointColl = CType(pGeometryOri, IPointCollection)
                        'Conserver le nombre de sommets
                        iNbPoints = pPointColl.PointCount

                        'Créer un multipoint vide
                        pMultipointAvant = New Multipoint
                        pMultipointAvant.SpatialReference = pGeometry.SpatialReference
                        'Interface pour ajouter les points de la géométrie
                        pPointColl = CType(pMultipointAvant, IPointCollection)
                        'Ajouter les points de la géométrie originale
                        pPointColl.AddPointCollection(CType(pGeometryOri, IPointCollection))
                        'Simplifier la géométrie
                        pTopoOp = CType(pMultipointAvant, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()

                        'Interface pour extraire le nombre de sommets
                        pPointColl = CType(pGeometry, IPointCollection)
                        'Vérifier si le nombre de sommets est différent
                        If iNbPoints <> pPointColl.PointCount Then
                            'Indiquer qu'une erreur de filtrage est présente
                            CorrigerFiltrageElement = True
                        End If

                        'Interface pour extraire les composantes
                        pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)

                        'Extraire la première composante
                        pTopoEdge = pEnumTopoEdge.Next

                        'Traiter toutes les composantes
                        Do Until pTopoEdge Is Nothing
                            'Vérifier si le Edge est une Polyline
                            If TypeOf (pTopoEdge.Geometry) Is IPolyline Then
                                'Interface pour extraire la ligne à traiter
                                pGeomColl = CType(pTopoEdge.Geometry, IGeometryCollection)

                                'Extraire la ligne à traiter
                                pPath = CType(pGeomColl.Geometry(0), IPath)
                            Else
                                'Définir la ligne à traiter
                                pPath = CType(pTopoEdge.Geometry, IPath)
                            End If

                            'Interface pour extraire le nombre de sommets
                            pPointColl = CType(pPath, IPointCollection)
                            'Conserver le nombre de sommets
                            iNbPoints = pPointColl.PointCount

                            'Généraliser la géométrie du TopologyEdge selon la distance latérale
                            pPath.Generalize(dDistLat)

                            'Interface pour extraire le nombre de sommets
                            pPointColl = CType(pPath, IPointCollection)
                            'Vérifier si une modification est présente
                            If iNbPoints <> pPointColl.PointCount Then
                                'Indiquer qu'il y a eu une modification
                                CorrigerFiltrageElement = True

                                'Mettre à jour la géométrie dans la topologie
                                pTopologyGraph.SetEdgeGeometry(pTopoEdge, pPath)
                            End If

                            'Extraire la première composante
                            pTopoEdge = pEnumTopoEdge.Next
                        Loop

                        'Définir la nouvelle géométrie filtrée de l'élément
                        pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                        'Simplifier la géométrie filtrée
                        pTopoOp = CType(pGeometry, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()

                        'Créer un multipoint vide
                        pMultipointApres = New Multipoint
                        pMultipointApres.SpatialReference = pGeometry.SpatialReference
                        'Interface pour ajouter les points de la géométrie
                        pPointColl = CType(pMultipointApres, IPointCollection)
                        'Ajouter les points de la géométrie
                        pPointColl.AddPointCollection(CType(pGeometry, IPointCollection))
                        'Simplifier la géométrie
                        pTopoOp = CType(pMultipointApres, ITopologicalOperator2)
                        pTopoOp.IsKnownSimple_2 = False
                        pTopoOp.Simplify()

                        'Interface pour extraire les points en erreur
                        pTopoOp = CType(pMultipointAvant, ITopologicalOperator2)
                        'Définir les points en erreur
                        pDifference = CType(pTopoOp.SymmetricDifference(pMultipointApres), IMultipoint)

                        'Vérifier la présence des points en erreur
                        If Not pDifference.IsEmpty Then
                            'Interface pour ajouter les points en erreur
                            pGeomCollAdd = CType(pBagErreurs, IGeometryCollection)
                            'Ajouter les points en erreur dans le Bag
                            pGeomCollAdd.AddGeometry(pDifference)
                        End If
                    End If

                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Traiter le Z et le M
                        Call TraiterZ(pGeometry)
                        Call TraiterM(pGeometry)

                        'Corriger la géométrie de l'élément
                        pFeature.Shape = pGeometry

                        'Sauver la correction
                        pWrite.WriteFeature(pFeature)
                    End If
                End If
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pGeometryOri = Nothing
            pDifference = Nothing
            pWrite = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pPointColl = Nothing
            pMultipointAvant = Nothing
            pMultipointApres = Nothing
            pTopoOp = Nothing
        End Try
    End Function
#End Region

#Region "Routines et fonctions pour corriger la longueur des droites des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la longueur des droites des éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour fusionner les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour filtrer les droites et sommets en trop des géométries des éléments.</param>
    '''<param name="bCorrigerPoints"> Permet d'indiquer si on doit corriger les sommets en trop des géométries des éléments.</param>
    '''<param name="bCorrigerLignes"> Permet d'indiquer si on doit corriger les droites et lignes en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerLongueurDroites(ByVal dPrecision As Double, ByVal dLongMin As Double, ByVal bCorrigerPoints As Boolean, ByVal bCorrigerLignes As Boolean, _
                                       ByVal bCreerFichierErreurs As Boolean, ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagPointsErreur As IGeometryBag = Nothing      'Interface contenant les points en trop dans la géométrie d'un élément.
        Dim pBagLignesErreur As IGeometryBag = Nothing      'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pLigne As IPolyline = Nothing                   'Interface contenant une ligne en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreur.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour extraire les points en erreur.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.
        Dim bModif As Boolean = False                       'Indique s'il y a eu une modification.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                pEnvelope = EnveloppeElementsSelectionner(pEditor)

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la longueur des droites (Longueur = " & dLongMin.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Vérifier si les droites de la géométrie de l'élément ont été corrigées
                    If CorrigerLongueurDroitesElement(pTopologyGraph, pFeature, dLongMin, bCorrigerPoints, pBagPointsErreur, pBagLignesErreur) Then
                        'Indiquer qu'une modification a été effectuée
                        bModif = True
                    End If

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Interface pour ajuter les points en erreur
                pGeomCollAdd = CType(pBagPointsErreur, IGeometryCollection)

                'Interface pour extraire les lignes en erreurs
                pGeomColl = CType(pBagLignesErreur, IGeometryCollection)

                'Vérifier si on doit corriger les lignes en erreur
                If pGeomColl.GeometryCount > 0 And bCorrigerLignes Then
                    'Définir la map active
                    pMap = m_MxDocument.ActiveView.FocusMap

                    'Traiter toutes les lignes en erreur
                    For i = 0 To pGeomColl.GeometryCount - 1
                        'Définir la ligne en erreur
                        pLigne = CType(pGeomColl.Geometry(i), IPolyline)

                        'Sélectionner les éléments au point de début de la ligne
                        pMap.SelectByShape(pLigne, Nothing, False)

                        'Déplacer tous les sommets du début de la ligne vers la fin de ligne
                        Call DeplacerSommetElements(pLigne.FromPoint, pLigne.ToPoint, dPrecision)

                        'Ajouter le point déplacé en erreur
                        pGeomCollAdd.AddGeometry(pLigne.FromPoint)
                    Next

                    'Indiquer qu'une modification a été effectuée
                    bModif = True
                End If

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomCollAdd.GeometryCount

                'Vérifier la présense d'une erreur
                If pGeomCollAdd.GeometryCount > 0 Then
                    'Vérifier la présecense d'une modification
                    If bModif Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la longueur des droites")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier la présense d'une erreur et si on doit créer le fichier d'erreurs
                    If pGeomCollAdd.GeometryCount > 0 And bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de longueur des droites (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerLongueurDroites_0", "Longueur de la droite plus petite que la dimension minimale : " & dLongMin.ToString, _
                                                                  m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPoint, pBagPointsErreur)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Sinon
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pLigne = Nothing
            pBagPointsErreur = Nothing
            pBagLignesErreur = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger la longueur des droites d'un élément selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="bCorrigerPoints"> Permet d'indiquer si la correction des sommets en trop doit être effectuée.</param>
    '''<param name="pBagPointsErreur"> Interface contenant les points en trop dans la géométrie de l'élément.</param>
    '''<param name="pBagLignesErreur"> Interface contenant les lignes et les droites en trop dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Public Function CorrigerLongueurDroitesElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal dLongMin As Double, bCorrigerPoints As Boolean, _
                                                   ByRef pBagPointsErreur As IGeometryBag, ByRef pBagLignesErreur As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerLongueurDroitesElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si le Bag des lignes en erreur est valide
                If pBagLignesErreur Is Nothing Then
                    'Créer le Bag vide des lignes en erreur
                    pBagLignesErreur = New GeometryBag
                    pBagLignesErreur.SpatialReference = pFeature.Shape.SpatialReference
                End If

                'Vérifier si le Bag des points en erreur est valide
                If pBagPointsErreur Is Nothing Then
                    'Créer le Bag vide des points en erreur
                    pBagPointsErreur = New GeometryBag
                    pBagPointsErreur.SpatialReference = pFeature.Shape.SpatialReference
                End If

                'Interface pour extraire les Edges de l'élément dans la topologie
                pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)

                'Corriger la longueur des droites de la géométrie de l'élément
                CorrigerLongueurDroitesElement = CorrigerLongueurDroitesGeometrie(pTopologyGraph, pEnumTopoEdge, dLongMin, pBagPointsErreur, pBagLignesErreur)

                'Vérifier si on doit corriger les sommets en trop
                If bCorrigerPoints Then
                    'Définir la nouvelle géométrie de l'élément
                    pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                    'Interface pour corriger les éléments sans affecter la topologie
                    pWrite = CType(pFeature.Class, IFeatureClassWrite)

                    'Vérifier si la géométrie est invalide
                    If pGeometry Is Nothing Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)

                        'Indiquer qu'il y a eu une modification
                        CorrigerLongueurDroitesElement = True

                        'Si la géométrie est valide
                    ElseIf CorrigerLongueurDroitesElement Then
                        'Traiter le Z et le M
                        Call TraiterZ(pGeometry)
                        Call TraiterM(pGeometry)

                        'Corriger la géométrie de l'élément
                        pFeature.Shape = pGeometry

                        'Sauver la correction
                        pWrite.WriteFeature(pFeature)
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pWrite = Nothing
            pEnumTopoEdge = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de corriger la longueur des droites d'une géométrie selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pEnumTopoEdge"> Interface contenant les Edges de la topologie de l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="pBagPointsErreur"> Interface contenant les points en trop dans la géométrie de l'élément.</param>
    '''<param name="pBagLignesErreur"> Interface contenant les lignes et les droites en trop dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Public Function CorrigerLongueurDroitesGeometrie(ByVal pTopologyGraph As ITopologyGraph, ByVal pEnumTopoEdge As IEnumTopologyEdge, ByVal dLongMin As Double, _
                                                     ByRef pBagPointsErreur As IGeometryBag, ByRef pBagLignesErreur As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un edge de la topologie. 
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour ajouter des géométries dans un Bag.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne. 
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire le nombre de sommets.
        Dim pProxOp As IProximityOperator = Nothing         'Interface pour calculer la la distance entre deux points.
        Dim pLigne As IPolyline = Nothing                   'Interface contenant une ligne ou une droite en erreur.
        Dim iNbPoints As Integer = 0                        'Contient le nombre de sommets.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerLongueurDroitesGeometrie = False

        Try
            'Si les Edges de l'élément sont valident
            If pEnumTopoEdge IsNot Nothing Then
                'Extraire la premier Edge
                pTopoEdge = pEnumTopoEdge.Next

                'Traiter tous les Edges
                Do Until pTopoEdge Is Nothing
                    'Vérifier si le Edge est une Polyline
                    If TypeOf (pTopoEdge.Geometry) Is IPolyline Then
                        'Interface pour extraire la ligne à traiter
                        pGeomColl = CType(pTopoEdge.Geometry, IGeometryCollection)

                        'Extraire la ligne à traiter
                        pPath = CType(pGeomColl.Geometry(0), IPath)
                    Else
                        'Définir la ligne à traiter
                        pPath = CType(pTopoEdge.Geometry, IPath)
                    End If

                    'Vérifier si le Edge est plus grand que la longueur minimale
                    If pPath.Length > dLongMin Then
                        'Interface pour extraire le nombre de sommets
                        pPointColl = CType(pPath, IPointCollection)
                        'Conserver le nombre de sommets
                        iNbPoints = pPointColl.PointCount

                        'Détruire les droites en trop
                        For i = pPointColl.PointCount - 1 To 1 Step -1
                            'Interface pour calculer la distance
                            pProxOp = CType(pPointColl.Point(i), IProximityOperator)

                            'Vérifier la longueur de la droite
                            If pProxOp.ReturnDistance(pPointColl.Point(i - 1)) < dLongMin Then
                                'Indiquer qu'il y a eu une modification
                                CorrigerLongueurDroitesGeometrie = True

                                'Si ce n'est pas le premier sommet
                                If i > 1 Then
                                    'Interface pour ajouter un point dans le Bag d'erreur
                                    pGeomCollAdd = CType(pBagPointsErreur, IGeometryCollection)
                                    'Ajouter le point en erreur dans le Bag
                                    pGeomCollAdd.AddGeometry(pPointColl.Point(i - 1))

                                    'Détruire le sommet précédent
                                    pPointColl.RemovePoints(i - 1, 1)

                                    'Si c'est le premier sommet et plus de 2 sommets
                                ElseIf pPointColl.PointCount > 2 Then
                                    'Interface pour ajouter un point dans le Bag d'erreur
                                    pGeomCollAdd = CType(pBagPointsErreur, IGeometryCollection)
                                    'Ajouter le point en erreur dans le Bag
                                    pGeomCollAdd.AddGeometry(pPointColl.Point(i))

                                    'Détruire le sommet courant
                                    pPointColl.RemovePoints(i, 1)
                                End If
                            End If
                        Next

                        'Vérifier si une modification est présente
                        If iNbPoints <> pPointColl.PointCount Then
                            'Indiquer qu'il y a eu une modification
                            CorrigerLongueurDroitesGeometrie = True

                            'Mettre à jour la géométrie dans la topologie
                            pTopologyGraph.SetEdgeGeometry(pTopoEdge, pPath)
                        End If

                        'Si le Edge est plus petit ou égal à la longueur minimale
                    Else
                        'Créer une polyligne vide d'erreur
                        pLigne = New Polyline
                        pLigne.SpatialReference = pPath.SpatialReference
                        'Interface pour ajouter une ligne dans la polyligne d'erreur
                        pGeomCollAdd = CType(pLigne, IGeometryCollection)
                        'Ajouter la ligne en erreur dans la polyligne
                        pGeomCollAdd.AddGeometry(pPath)

                        'Interface pour ajouter une ligne dans le Bag d'erreur
                        pGeomCollAdd = CType(pBagLignesErreur, IGeometryCollection)
                        'Ajouter la ligne en erreur dans le Bag
                        pGeomCollAdd.AddGeometry(pLigne)
                    End If

                    'Extraire les Edges suivants
                    pTopoEdge = pEnumTopoEdge.Next
                Loop
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pTopoEdge = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPath = Nothing
            pPointColl = Nothing
            pProxOp = Nothing
            pLigne = Nothing
        End Try
    End Function
    '''<summary>
    ''' Routine qui permet de déplacer un sommet sur tous les éléments sélectionnable de la fenêtre graphique active selon la tolérence active. 
    '''</summary>
    '''
    '''<param name="pPointA">Interface ESRI contenant le point de départ.</param>
    '''<param name="pPointB">Interface ESRI contenant le point d'arrivé.</param>
    '''<param name="dTol">Contient la tolérence de recherche.</param>
    ''' 
    Public Sub DeplacerSommetElements(ByVal pPointA As IPoint, ByVal pPointB As IPoint, Optional ByVal dTol As Double = 5.0)
        'Déclarer les variables de travail
        Dim pMap As IMap = Nothing                      'Interface ESRI contenant la Map active.
        Dim pEnumFeature As IEnumFeature = Nothing      'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing              'Interface ESRI contenant un élément en sélection.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie de l'élément.
        Dim pGeodataset As IGeoDataset = Nothing        'Interface contenant la référence spatial de la classe
        Dim pGeometryDef As IGeometryDef = Nothing      'Interface ESRI qui permet de vérifier la présence du Z et du M.
        Dim pEnv As ISelectionEnvironment = Nothing     'Interface contenant l'environnement de recherche
        Dim pPolygon As IPolygon = Nothing              'Interface pour calculer la longueur d'une surface.
        Dim pPolyline As IPolyline = Nothing            'Interface pour calculer la longueur d'une ligne.
        Dim bModif As Boolean = False                   'Indique si une modification a été effectuée.

        Try
            'Définir la Map courante
            pMap = m_MxDocument.FocusMap

            'Vérifier si des éléments sont sélectionnés
            If pMap.SelectionCount = 0 Then
                'Définir un environnement de recherche
                pEnv = New SelectionEnvironment
                pEnv.CombinationMethod = esriSelectionResultEnum.esriSelectionResultNew
                'Sélectionner les éléments qui intersecte le point selon la distance active
                pMap.SelectByShape(fpBufferGeometrie(pPointA, dTol), pEnv, False)
            End If

            'Interface pour extraire le premier élément de la sélection
            pEnumFeature = CType(pMap.FeatureSelection, IEnumFeature)

            'Extraire le premier élément de la sélection
            pFeature = pEnumFeature.Next

            'Traite tous les éléments sélectionnés
            Do Until pFeature Is Nothing
                'Interface contenant la référence spatiale
                pGeodataset = CType(pFeature.Class, IGeoDataset)
                'Définir la géométrie de l'élément
                pGeometry = pFeature.ShapeCopy
                'Déplacer et vérifier si le sommet a été déplacé
                If DeplacerSommetGeometrie(pGeometry, pPointA, pPointB, dTol, False, True) Then
                    'Indiquer qu'il y a eu une modification
                    bModif = True
                    'Projet la géométrie
                    pGeometry.Project(pGeodataset.SpatialReference)
                    'Interface pour vérifier la présence du Z et du M
                    pGeometryDef = RetournerGeometryDef(CType(pFeature.Class, IFeatureClass))
                    'Vérifier la présence du Z
                    If pGeometryDef.HasZ Then
                        'Traiter le Z
                        Call TraiterZ(pGeometry)
                    End If
                    'Vérifier la présence du M
                    If pGeometryDef.HasM Then
                        'Traiter le M
                        Call TraiterM(pGeometry)
                    End If
                    'Changer la géométrie de l'élément
                    pFeature.Shape = pGeometry

                    'Si la géométrie est une surface
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                        'Interface pour calculer la longueur d'une surface
                        pPolygon = CType(pGeometry, IPolygon)
                        'Vérifier si la longueur est nulle
                        If pPolygon.Length = 0 Then
                            'Détruire l'élément
                            pFeature.Delete()
                        Else
                            'Sauver la modification
                            pFeature.Store()
                        End If

                        'Si la géométrie est une ligne
                    ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                        'Interface pour calculer la longueur d'une ligne
                        pPolyline = CType(pGeometry, IPolyline)
                        'Vérifier si la longueur est nulle
                        If pPolyline.Length = 0 Then
                            'Détruire l'élément
                            pFeature.Delete()
                        Else
                            'Sauver la modification
                            pFeature.Store()
                        End If

                        'Sinon
                    Else
                        'Sauver la modification
                        pFeature.Store()
                    End If
                End If

                'Extraire le prochain élément de la sélection
                pFeature = pEnumFeature.Next
            Loop

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            pMap = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnv = Nothing
            pGeodataset = Nothing
            pGeometryDef = Nothing
            pEnv = Nothing
            pPolygon = Nothing
            pPolyline = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de déplacer un sommet existant sur une géométrie en fonction d'un point, d'une tolérance de recherche
    ''' et d'un point de déplacement.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI contenant la géométrie à traiter.</param>
    '''<param name="pPointA">Interface ESRI contenant le sommet à rechercher sur la géométrie selon une tolérance.</param> 
    '''<param name="pPointB">Interface ESRI contenant le sommet résultant du sommet trouvé.</param> 
    '''<param name="dTolerance">Contient la tolérance de recherche du sommet à rechercher sur la géométrie.</param>
    '''<param name="bSimplifier">Indiquer si on doit simplifier la géométrie après avoir déplacé le sommet.</param>
    '''
    '''<returns>TRUE pour indiquer si le sommet a été déplacé, FALSE sinon.</returns>
    '''
    Public Function DeplacerSommetGeometrie(ByRef pGeometry As IGeometry, ByVal pPointA As IPoint, ByVal pPointB As IPoint, _
    Optional ByVal dTolerance As Double = 0, Optional ByVal bSeulementUn As Boolean = False, _
    Optional ByVal bSimplifier As Boolean = True) As Boolean
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface utilisée pour extraire la limite d'une géométrie
        Dim pPath As IPath = Nothing                    'Interface ESRI utilisé pour vérifier si la partie de la géométrie est fermée
        Dim pClone As IClone = Nothing                  'Interface ESRI utilisé pour cloner une géométrie
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface ESRI utilisé pour accéder aux parties de la géométrie
        Dim pPointColl As IPointCollection = Nothing    'Interface ESRI utilisé pour détruire un sommet
        Dim pHitTest As IHitTest = Nothing              'Interface pour tester la présence du sommet recherché
        Dim pNewPoint As IPoint = Nothing               'Interface contenant le sommet trouvé
        Dim pProxOp As IProximityOperator = Nothing     'Interface qui permet de calculer la distance
        Dim dDistance As Double = Nothing               'Interface contenant la distance calculée entre le point de recherche et le sommet trouvé
        Dim nNumeroPartie As Integer = Nothing          'Numéro de partie trouvée
        Dim nNumeroSommet As Integer = Nothing          'Numéro de sommet de la partie trouvée
        Dim bCoteDroit As Boolean = Nothing             'Indiquer si le point trouvé est du côté droit de la géométrie
        Dim j As Integer = Nothing                      'Compteur

        Try
            'Écrire une trace de début
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Debut")

            'Initialiser la valeur de retour
            DeplacerSommetGeometrie = False

            'Interface pour extraire le sommet et le numéro de sommet de proximité
            pHitTest = CType(pGeometry, IHitTest)

            'Vérifier s'il s'agit d'un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Interface pour calculer la distance
                pProxOp = CType(pPointA, IProximityOperator)

                'Vérifier la distance
                If pProxOp.ReturnDistance(pGeometry) <= dTolerance Then
                    'Interface pour cloner la géométrie
                    pClone = CType(pPointB, IClone)

                    'Redéfinir la géométrie
                    pGeometry = CType(pClone.Clone, IGeometry)

                    'Indiquer que le sommet a été déplacé
                    DeplacerSommetGeometrie = True
                End If

                'Vérifier s'il s'agit d'un multipoint
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryMultipoint Then
                'Rechercher le point par rapport à chaque sommet de la géométrie
                If pHitTest.HitTest(pPointA, dTolerance, esriGeometryHitPartType.esriGeometryPartVertex, _
                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Interface pour détruire le sommet de proximité
                    pPointColl = CType(pGeometry, IPointCollection)

                    'Déplacer le sommet
                    pPointColl.UpdatePoint(nNumeroSommet, pPointB)

                    'Indiquer que le sommet a été déplacé
                    DeplacerSommetGeometrie = True
                End If

                'Vérifier s'il s'agit d'une ligne ou d'une surface
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Or _
            pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Rechercher le point par rapport à chaque sommet de la géométrie
                If pHitTest.HitTest(pPointA, dTolerance, esriGeometryHitPartType.esriGeometryPartVertex, _
                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Interface pour accéder aux parties de la géométrie
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Interface pour vérifier si la partie est fermée
                    pPath = CType(pGeomColl.Geometry(nNumeroPartie), IPath)

                    'Interface pour déplacer le sommet
                    pPointColl = CType(pPath, IPointCollection)

                    'Déplacer le sommet trouvé
                    pPointColl.UpdatePoint(nNumeroSommet, pPointB)

                    'Vérifier si seulement un sommet doit être traité
                    If bSeulementUn Then
                        'Vérifier si la partie de la ligne est fermé et que le numéro de sommet est le premier
                        If pPath.IsClosed And nNumeroSommet = 0 Then
                            'Mettre à jour le dernier sommet
                            pPointColl.UpdatePoint(pPointColl.PointCount - 1, pPointB)
                        End If

                        'Si tous les sommets doivent être traité
                    Else
                        'Interface pour calculer la distance
                        pProxOp = CType(pPointA, IProximityOperator)
                        'Redéfinir le numéro de partie au début
                        nNumeroPartie = 0
                        'Traiter tous les Path
                        Do Until (nNumeroPartie > pGeomColl.GeometryCount - 1)
                            'Interface pour traiter chaque Path
                            pPath = CType(pGeomColl.Geometry(nNumeroPartie), IPath)
                            'Interface pour déplacer le sommet
                            pPointColl = CType(pPath, IPointCollection)

                            'Traiter tous les sommets du Path
                            For j = 0 To pPointColl.PointCount - 1
                                'Vérifier la distance
                                If pProxOp.ReturnDistance(pPointColl.Point(j)) <= dTolerance Then
                                    'Déplacer le sommet trouvé
                                    pPointColl.UpdatePoint(j, pPointB)
                                End If
                            Next

                            'Vérifier la longueur du Path
                            If pPath.Length = 0 Then
                                'Détruire le path
                                pGeomColl.RemoveGeometries(nNumeroPartie, 1)
                            Else
                                'Changer de numéro de partie
                                nNumeroPartie = nNumeroPartie + 1
                            End If
                        Loop
                    End If

                    'Indiquer que le sommet a été détruit
                    DeplacerSommetGeometrie = True

                    'Sinon c'est l'absence d'un sommet
                ElseIf (pHitTest.HitTest(pPointA, dTolerance, esriGeometryHitPartType.esriGeometryPartBoundary, _
                pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit)) Then
                    'Interface pour accéder aux parties de la géométrie
                    pGeomColl = CType(pGeometry, IGeometryCollection)

                    'Interface pour vérifier si la partie est fermée
                    pPath = CType(pGeomColl.Geometry(nNumeroPartie), IPath)

                    'Interface pour déplacer le sommet
                    pPointColl = CType(pPath, IPointCollection)

                    'Insérer un nouveau sommet
                    pPointColl.InsertPoints(nNumeroSommet + 1, 1, pPointB)

                    'Indiquer que le sommet a été détruit
                    DeplacerSommetGeometrie = True
                End If
            End If

            'Vérifier si un déplacement a eu lieu et que ce n'est pas un point
            If DeplacerSommetGeometrie And pGeometry.GeometryType <> esriGeometryType.esriGeometryPoint And bSimplifier Then
                'Simplifier la géométrie
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
            End If

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pClone = Nothing
            pPath = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pHitTest = Nothing
            pNewPoint = Nothing
            pProxOp = Nothing
            dDistance = Nothing
            nNumeroPartie = Nothing
            nNumeroSommet = Nothing
            bCoteDroit = Nothing
            'Écrire une trace de fin
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Fin")
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de définir les segments utilisées dans un FeedBack de déplacement de sommets pour plusieurs éléments. 
    '''</summary>
    '''
    '''<param name="pPoint">Interface ESRI contenant le point de recherche.</param>
    '''<param name="pVertexFeedback">Interface ESRI contenant les segments utilisés pour voir les déplacement.</param>
    '''<param name="dTol">Contient la tolérence de recherche.</param>
    ''' 
    Public Sub DefinirSegment(ByVal pPoint As IPoint, ByRef pVertexFeedback As IVertexFeedback, Optional ByVal dTol As Double = 5.0)
        'Déclarer les variables de travail
        Dim pDatasetEdit As IDatasetEdit = Nothing      'Interface qui permet d'indiquer si l'élément est en mode édition
        Dim pMap As IMap = Nothing                      'Interface ESRI contenant la Map active.
        Dim pEnumFeature As IEnumFeature = Nothing      'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing              'Interface ESRI contenant un élément en sélection.
        Dim pGeometry As IGeometry = Nothing            'Interface contenant la géométrie de l'élément.
        Dim pPolyline As IPolyline = Nothing            'Interface contenant une ligne.
        Dim pSegmentColl As ISegmentCollection = Nothing 'Interface contenant la géométrie de l'élément.
        Dim pGeometryColl As IGeometryCollection = Nothing 'Interface contenant la géométrie de l'élément.
        Dim pPointColl As IPointCollection = Nothing    'Interface ESRI utilisé pour détruire un sommet
        Dim pEnv As ISelectionEnvironment = Nothing     'Interface contenant l'environnement de recherche
        Dim pHitTest As IHitTest = Nothing              'Interface pour extraire les numéros de segment
        Dim nNoVertex As Integer = Nothing              'Contient le numéro de sommet en commencant par 0
        Dim nNoPartie As Integer = Nothing              'Contient le numéro de partie
        Dim bCoteDroit As Boolean = Nothing             'Indiquer si le point trouvé est du côté droit de la géométrie
        Dim bHitTest As Boolean = Nothing               'Indique si un vertex est présent dans la tolérance

        Try
            'Écrire une trace de début
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Debut")

            'Définir la Map courante
            pMap = m_MxDocument.FocusMap

            'Vérifier si des éléments sont sélectionnés
            If pMap.SelectionCount = 0 Then
                'Définir un environnement de recherche
                pEnv = New SelectionEnvironment
                pEnv.CombinationMethod = esriSelectionResultEnum.esriSelectionResultNew
                'Sélectionner les éléments qui intersecte le point selon la distance active
                pMap.SelectByShape(fpBufferGeometrie(pPoint, dTol), pEnv, False)
            End If

            'Interface pour extraire le premier élément de la sélection
            pEnumFeature = CType(pMap.FeatureSelection, IEnumFeature)

            'Extraire le premier élément de la sélection
            pFeature = pEnumFeature.Next

            'Traite tous les éléments sélectionnés
            Do Until pFeature Is Nothing
                'Interface qui permet d'indiquer si l'élément est en mode édition
                pDatasetEdit = CType(pFeature.Class, IDatasetEdit)
                'Vérifier si l'élément est en mode édition
                If pDatasetEdit.IsBeingEdited Then
                    'Définir la géométrie de l'élément
                    pGeometry = pFeature.ShapeCopy

                    'Vérifier si la géométrie est une ligne ou une surface
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline _
                    Or pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                        'Interface pour extraire la partie de la géométrie
                        pGeometryColl = CType(pGeometry, IGeometryCollection)
                        'Interface pour trouver les numéro de segment
                        pHitTest = CType(pGeometry, IHitTest)
                        'Définir si le point intersecte un sommet de la géométrie à l'intérieur d'une tolérance
                        bHitTest = pHitTest.HitTest(pPoint, dTol, esriGeometryHitPartType.esriGeometryPartVertex, Nothing, 0, _
                                   nNoPartie, nNoVertex, bCoteDroit)
                        'Vérifier si le point intersecte un sommet de la géométrie à l'intérieur d'une tolérance
                        If Not bHitTest Then
                            'Définir si le point intersecte un sommet de la géométrie à l'intérieur d'une tolérance
                            bHitTest = pHitTest.HitTest(pPoint, dTol, esriGeometryHitPartType.esriGeometryPartBoundary, Nothing, 0, _
                                       nNoPartie, nNoVertex, bCoteDroit)
                            'Définir si le point intersecte une droite de la géométrie à l'intérieur d'une tolérance
                            If bHitTest Then
                                'Interface pour extraire la partie de la géométrie
                                pPointColl = CType(pGeometryColl.Geometry(nNoPartie), IPointCollection)
                                'Insérer un nouveau sommet
                                pPointColl.InsertPoints(nNoVertex + 1, 1, pPoint)
                                'Définir si le point intersecte un sommet de la géométrie à l'intérieur d'une tolérance
                                bHitTest = pHitTest.HitTest(pPoint, dTol, esriGeometryHitPartType.esriGeometryPartVertex, Nothing, 0, _
                                           nNoPartie, nNoVertex, bCoteDroit)
                            End If
                        End If

                        'Vérifier si le point intersecte un sommet de la géométrie à l'intérieur d'une tolérance
                        If bHitTest Then
                            'Interface pour extraire le segment dans la partie de la géométrie
                            pSegmentColl = CType(pGeometryColl.Geometry(nNoPartie), ISegmentCollection)
                            'Vérifier si la géométrie est une ligne
                            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolyline Then
                                'Interface pour vérifier si la ligne est fermé
                                pPolyline = CType(pGeometry, IPolyline)
                                'Vérifier si le numéro de segment est le premier
                                If nNoVertex = 0 Then
                                    'Ajouter le premier segment en indiquant que c'est le FromPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(0), False)
                                    'Vérifier si la ligne est fermé
                                    If pPolyline.IsClosed Then
                                        'Ajouter le dernier segment en indiquant que c'est le toPoint qui doit être déplacé
                                        pVertexFeedback.AddSegment(pSegmentColl.Segment(pSegmentColl.SegmentCount - 1), True)
                                    End If

                                    'Si le numéro de segment est le dernier
                                ElseIf nNoVertex = pSegmentColl.SegmentCount Then
                                    'Ajouter le dernier segment en indiquant que c'est le toPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(pSegmentColl.SegmentCount - 1), True)
                                    'Vérifier si la ligne est fermé
                                    If pPolyline.IsClosed Then
                                        'Ajouter le premier segment en indiquant que c'est le FromPoint qui doit être déplacé
                                        pVertexFeedback.AddSegment(pSegmentColl.Segment(0), False)
                                    End If

                                    'Si le numéro de sommet n'est ni le premier ni le dernier
                                Else
                                    'Ajouter le segment trouvé en indiquant que c'est le FromPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(nNoVertex - 1), True)
                                    'Ajouter le dernier segment en indiquant que c'est le ToPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(nNoVertex), False)
                                End If

                                'Si la géométrie est une surface
                            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                                'Ajouter le segment trouvé en indiquant que c'est le FromPoint qui doit être déplacé
                                pVertexFeedback.AddSegment(pSegmentColl.Segment(nNoVertex - 1), True)
                                'Vérifier si le numéro de segment est le premier et si la ligne est fermée
                                If nNoVertex = 0 Then
                                    'Ajouter le premier segment en indiquant que c'est le FromPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(0), True)
                                    'Ajouter le dernier segment en indiquant que c'est le ToPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(pSegmentColl.SegmentCount - 1), False)

                                    'Si le numéro de segment est le dernier
                                ElseIf nNoVertex = pSegmentColl.SegmentCount Then
                                    'Ajouter le dernier segment en indiquant que c'est le toPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(pSegmentColl.SegmentCount - 1), False)
                                    'Ajouter le premier segment en indiquant que c'est le FromPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(0), True)

                                    'Si le numéro de sommet n'est ni le premier ni le dernier
                                Else
                                    'Ajouter le segment trouvé en indiquant que c'est le FromPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(nNoVertex - 1), True)
                                    'Ajouter le dernier segment en indiquant que c'est le ToPoint qui doit être déplacé
                                    pVertexFeedback.AddSegment(pSegmentColl.Segment(nNoVertex), False)
                                End If
                            End If
                        End If
                    End If
                End If

                'Extraire le prochain élément de la sélection
                pFeature = pEnumFeature.Next
            Loop

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pDatasetEdit = Nothing
            pMap = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pGeometry = Nothing
            pPolyline = Nothing
            pGeometryColl = Nothing
            pSegmentColl = Nothing
            pPointColl = Nothing
            pEnv = Nothing
        End Try
    End Sub

    ''' <summary> 
    ''' Cette fonction permet de retourner la définition de la géométrie à partir de la classe afin de vérifier la présence du Z et M.
    ''' </summary>
    ''' 
    ''' <param name="pFeatureClass"></param>
    ''' 
    ''' <returns>IGeometryDef</returns>
    ''' 
    Public Function RetournerGeometryDef(ByVal pFeatureClass As IFeatureClass) As IGeometryDef
        Dim shapeFieldName As String = pFeatureClass.ShapeFieldName
        Dim fields As IFields = pFeatureClass.Fields
        Dim geometryIndex As Integer = fields.FindField(shapeFieldName)
        Dim field As IField = fields.Field(geometryIndex)
        Dim geometryDef As IGeometryDef = field.GeometryDef
        Return geometryDef
    End Function
#End Region

#Region "Routines et fonctions pour corriger la longueur des lignes des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la longueur des lignes des éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour filtrer les lignes en trop des géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerLongueurLignes(ByVal dPrecision As Double, ByVal dLongMin As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, _
                                      ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pLigneErreur As IPolyline = Nothing             'Interface contenant une ligne en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.
        Dim iNbErr As Long = -1                             'Contient le nombre d'erreurs.
        Dim dLongTmp As Double = 0                          'Contient la longueur minimale temporaire.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Traiter qu'il y a des erreurs ou que la longueur temporaire n'est pas égale à la longueur minimale
                Do While iNbErreurs > iNbErr Or dLongTmp < dLongMin
                    'Intialisation
                    iNbErr = iNbErreurs

                    'Vérifier si la longueur minimal de traitement n'est pas spécifiée
                    If dLongTmp = 0 Then
                        'Définir la longueur minimale temporaire de début
                        dLongTmp = dLongMin / 2
                    Else
                        'Définir la longueur minimale
                        dLongTmp = dLongMin
                    End If

                    'Création de la topologie
                    pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                    pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                    'Initialiser la barre de progression
                    pTrackCancel.Progressor.Message = "Correction de la longueur des lignes (NbErr=" & iNbErreurs.ToString & ", LongMin=" & dLongTmp.ToString & ") ..."
                    InitBarreProgression(0, pEditor.SelectionCount, pTrackCancel)

                    'Interface pour extraire le premier élément de la sélection
                    pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                    'Réinitialiser la recherche des éléments
                    pEnumFeature.Reset()
                    'Extraire le premier élément de la sélection
                    pFeature = pEnumFeature.Next

                    'Traite tous les éléments sélectionnés
                    Do Until pFeature Is Nothing
                        'Vérifier si les lignes de la géométrie de l'élément ont été corrigées
                        If CorrigerLongueurLignesElement(pTopologyGraph, pFeature, dLongTmp, bCorriger, pLigneErreur) Then
                            'Ajouter l'erreur dans le Bag
                            pGeomColl.AddGeometry(pLigneErreur)
                        End If

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Exit Do

                        'Extraire le prochain élément de la sélection
                        pFeature = pEnumFeature.Next
                    Loop

                    'Retourner le nombre d'erreurs
                    iNbErreurs = pGeomColl.GeometryCount
                Loop

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la longueur des lignes")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de longueur des lignes (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerLongueurLignes_1", "Longueur de la ligne plus petite que la dimension minimale : " & dLongMin.ToString, _
                                                                  m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pLigneErreur = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger la longueur des lignes d'un élément selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="bCorriger"> Permet d'indiquer si la correction des lignes en trop doit être effectuée.</param>
    '''<param name="pLigneErreur"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerLongueurLignesElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal dLongMin As Double, bCorriger As Boolean, _
                                                   ByRef pLigneErreur As IPolyline) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour enlever la partie de la géométrie en erreur.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerLongueurLignesElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si l'élément est une ligne
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Créer le Bag vide des lignes en erreur
                    pLigneErreur = New Polyline
                    pLigneErreur.SpatialReference = pFeature.Shape.SpatialReference

                    'Interface pour extraire les Edges de l'élément dans la topologie
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)
                    pEnumTopoEdge.Reset()

                    'Corriger la longueur des lignes de la géométrie de l'élément
                    CorrigerLongueurLignesElement = CorrigerLongueurLignesGeometrie(pTopologyGraph, pEnumTopoEdge, dLongMin, pLigneErreur)

                    'Vérifier si on doit corriger les sommets en trop
                    If bCorriger And CorrigerLongueurLignesElement Then
                        'Interface pour extraire les Edges de l'élément dans la topologie
                        pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)
                        'Définir la nouvelle géométrie de l'élément
                        pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                        'Interface pour corriger les éléments sans affecter la topologie
                        pWrite = CType(pFeature.Class, IFeatureClassWrite)

                        'Vérifier si la géométrie est invalide
                        If pGeometry Is Nothing Or pEnumTopoEdge.Count = 0 Then
                            'Détruire l'élément
                            pWrite.RemoveFeature(pFeature)
                            'Indiquer qu'il y a eu une modification
                            CorrigerLongueurLignesElement = True

                            'Si la géométrie est valide
                        Else
                            'Indiquer qu'il y a eu une modification
                            CorrigerLongueurLignesElement = True

                            'Interface pour enlever la partie de la géométrie en erreur
                            pTopoOp = CType(pGeometry, ITopologicalOperator2)
                            'Enlever la partie de la géométrie en erreur
                            pGeometry = pTopoOp.Difference(pLigneErreur)

                            'Si la géométrie est vide
                            If pGeometry.IsEmpty Then
                                'Détruire l'élément
                                pWrite.RemoveFeature(pFeature)

                                'Sinon
                            Else
                                'Traiter le Z et le M
                                Call TraiterZ(pGeometry)
                                Call TraiterM(pGeometry)

                                'Corriger la géométrie de l'élément
                                pFeature.Shape = pGeometry

                                'Sauver la correction
                                pWrite.WriteFeature(pFeature)
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pWrite = Nothing
            pEnumTopoEdge = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de corriger la longueur des lignes d'une géométrie selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pEnumTopoEdge"> Interface contenant les Edges de la topologie de l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="pLigneErreur"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerLongueurLignesGeometrie(ByVal pTopologyGraph As ITopologyGraph, ByVal pEnumTopoEdge As IEnumTopologyEdge, ByVal dLongMin As Double, _
                                                     ByRef pLigneErreur As IPolyline) As Boolean
        'Déclarer les variables de travail
        Dim pTopoNodeInt As ITopologyNode = Nothing         'Interface contenant un noeud d'intersection. 
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie. 
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans un Bag.
        Dim pLigneEdges As IPolyline = Nothing              'Interface contenant la ligne des Edges continus.
        Dim pCollEdges As Collection = Nothing              'Objet contenant une collection de Edges.

        'Par défaut, une modification a été effectuée
        CorrigerLongueurLignesGeometrie = False

        Try
            'Si les Edges de l'élément sont valident
            If pEnumTopoEdge IsNot Nothing Then
                'Si au moins un Edge est présent
                If pEnumTopoEdge.Count > 0 Then
                    'Interface pour ajouter les lignes en erreur
                    pGeomColl = CType(pLigneErreur, IGeometryCollection)

                    'Extraire la premier Edge
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter tous les Edges
                    Do Until pTopoEdge Is Nothing
                        'Vérifier si le Edge n'est pas traité
                        If Not pTopoEdge.Visited Then
                            'Vérifier si le Edge contient une fin de ligne 
                            If pTopoEdge.FromNode.Edges(True).Count = 1 Or pTopoEdge.ToNode.Edges(True).Count = 1 Then
                                'Si le noeud de début est une extrémité
                                If pTopoEdge.FromNode.Edges(True).Count = 1 Then
                                    'Définir la ligne des Edges continus
                                    Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoEdge, pLigneEdges, pCollEdges, pTopoNodeInt)

                                    'Si le noeud de fin est une extrémité
                                Else
                                    'Définir la ligne des Edges continus
                                    Call IdentifierLigneEdge(pTopoEdge.ToNode, pTopoEdge, pLigneEdges, pCollEdges, pTopoNodeInt)
                                End If

                                'Vérifier si la ligne est fermée ou si aucun noeud d'intersection n'est présent (ligne isolée)
                                If pLigneEdges.IsClosed Or pTopoNodeInt Is Nothing Then
                                    'Vérifier si la ligne Edge est plus petite que la longueur minimale
                                    If pLigneEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                                    End If

                                    'Si la ligne n'est pas fermée
                                Else
                                    'Traiter le noeud d'intersection
                                    Call TraiterNoeudIntersection(dLongMin, pTopologyGraph, pTopoNodeInt, pLigneEdges, pCollEdges)
                                End If

                                'Si le Edge ne contient pas une fin de ligne 
                            Else
                                'Définir la ligne des Edges continus
                                Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoEdge, pLigneEdges, pCollEdges, pTopoNodeInt, False)

                                'Vérifier si la ligne est fermée ou  si aucun noeud d'intersection n'est présent (ligne isolée)
                                If pLigneEdges.IsClosed Then
                                    'Vérifier si la ligne Edge est plus petite que la longueur minimale
                                    If pLigneEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                                    End If
                                End If
                            End If
                        End If

                        'Vérifier si on doit détruire le Edge
                        If pTopoEdge.IsSelected Then
                            'Détruire le Edge dans la topologie
                            pTopologyGraph.DeleteEdge(pTopoEdge)

                            'Ajouter la ligne du Edge en erreur
                            pGeomColl.AddGeometryCollection(CType(pTopoEdge.Geometry, IGeometryCollection))

                            'Indiquer qu'une correction a été effectuée
                            CorrigerLongueurLignesGeometrie = True
                        End If

                        'Extraire les Edges suivants
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pTopoNodeInt = Nothing
            pTopoEdge = Nothing
            pGeomColl = Nothing
            pLigneEdges = Nothing
            pCollEdges = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter un noeud d'intersection afin de conserver l'extrémité la plus longue. 
    '''</summary>
    ''' 
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les droites en trop dans la géométrie d'un élément.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pTopoNode"> Interface contenant le noeud d'intersection.</param>
    '''<param name="pLigneEdges"> Interface contenant la ligne des Edges continus.</param>
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    ''' 
    Private Sub TraiterNoeudIntersection(ByVal dLongMin As Double, ByVal pTopologyGraph As ITopologyGraph, ByVal pTopoNode As ITopologyNode, _
                                         ByVal pLigneEdges As IPolyline, ByVal pCollEdges As Collection)
        'Déclarer les variables de travail
        Dim pTopoNodeEdge As ITopologyEdge = Nothing        'Interface contenant un Edge d'un noeud de la topologie. 
        Dim pEnumNodeEdges As IEnumNodeEdge = Nothing       'Interface contenant la ligne des Edges continus du noeud d'intersection.
        Dim pLigneNodeEdges As IPolyline = Nothing          'Interface contenant les Edges d'un noeud.
        Dim pCollNodeEdges As Collection = Nothing          'Collection des Edges de la ligne continue du noeud d'intersection.
        Dim pTopoNodeLigne As ITopologyNode = Nothing       'Interface contenant un noeud de la topologie. 
        Dim iNbExt As Integer = 0                           'Contient le nombre d'extémités.
        Dim iNbEdge As Integer = 0                          'Contient le nombre d'éléments.

        Try
            'Vérifier si le noeud possède plus de 2 Edges
            If Not pTopoNode.Visited Then
                'Indiquer que le noeud a été visité
                pTopoNode.Visited = True

                'Vérifier si le noeud possède plus de 2 Edges
                If pTopoNode.Edges(True).Count > 2 Then
                    'Interface pour extraire les Edges du noeud
                    pEnumNodeEdges = pTopoNode.Edges(True)

                    'Extraire la premier Edge
                    pEnumNodeEdges.Next(pTopoNodeEdge, True)

                    'Traiter tous les Edges du noeud
                    Do Until pTopoNodeEdge Is Nothing
                        'Vérifier si le Edge est déjà traité
                        If Not pTopoNodeEdge.Visited Then
                            'Compter le nombre de Edges
                            iNbEdge = iNbEdge + 1

                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoNode, pTopoNodeEdge, pLigneNodeEdges, pCollNodeEdges, pTopoNodeLigne, False)

                            'Vérifier si aucun noeud d'intersection autre que celui de début n'a été trouvé
                            'Si la ligne est une extrémité
                            If pTopoNodeLigne Is Nothing Then
                                'Compter le nombre d'extrémités trouvées
                                iNbExt = iNbExt + 1

                                'Initialiser les noeuds visités
                                Call InitEdgesVisiter(pCollEdges, True)

                                'Vérifier si la ligne continue de départ est la plus longue
                                If pLigneEdges.Length > pLigneNodeEdges.Length Then
                                    'Vérifier si la ligne est inférieure à la logueur minimale
                                    If pLigneNodeEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollNodeEdges)
                                    End If

                                    'Si la ligne continue de départ n'est plus la plus longue
                                Else
                                    'Vérifier si la ligne est inférieure à la logueur minimale
                                    If pLigneEdges.Length <= dLongMin Then
                                        'Sélectionner les Edges de la topologie à détruire
                                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                                    End If

                                    'Redéfinir la ligne continue la plus longue 
                                    pLigneEdges = pLigneNodeEdges
                                    'Redéfinir la collection des edges de la ligne continue la plus longue 
                                    pCollEdges = pCollNodeEdges
                                End If

                                'Si la ligne n'est pas une extrémité
                            Else
                                'Mettre l'attribut Visited=False pour tous les Edges
                                Call InitEdgesVisiter(pCollNodeEdges)
                            End If
                        End If

                        'Extraire le prochain Edge
                        pEnumNodeEdges.Next(pTopoNodeEdge, True)
                    Loop

                    'Vérifier si aucune autre extrémité n'a été trouvée
                    If iNbEdge > 0 And iNbExt = 0 And pLigneEdges.Length <= dLongMin Then
                        'Sélectionner les Edges de la topologie à détruire
                        Call SelectionnerEdges(pTopologyGraph, pCollEdges)
                    End If
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEnumNodeEdges = Nothing
            pTopoNodeEdge = Nothing
            pLigneNodeEdges = Nothing
            pCollNodeEdges = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'identifier tous les Edges d'une lignes continue. 
    '''</summary>
    ''' 
    '''<param name="pTopoNodeDebut"> Interface contenant le noeud d'intersection de début.</param>
    '''<param name="pTopoEdge"> Interface contenant le Edge à traiter.</param>
    '''<param name="pLigneEdges"> Interface contenant la ligne des Edges continus.</param>
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    '''<param name="pTopoNodeFin"> Interface contenant le noeud d'intersection de fin.</param>
    '''<param name="bVisiter"> Indiquer si on doit indiquer si le Edge a été visité.</param>
    '''<param name="bInitialiser"> Indiquer si on doit indiquer si le Edge a été visité.</param>
    ''' 
    Private Sub IdentifierLigneEdge(ByVal pTopoNodeDebut As ITopologyNode, ByVal pTopoEdge As ITopologyEdge, _
                                    ByRef pLigneEdges As IPolyline, ByRef pCollEdges As Collection, ByRef pTopoNodeFin As ITopologyNode,
                                    Optional ByVal bVisiter As Boolean = True, Optional ByVal bInitialiser As Boolean = True)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des lignes.
        Dim pTopoNodeEdge As ITopologyEdge = Nothing        'Interface contenant un Edge d'un noeud de la topologie. 
        Dim pEnumNodeEdges As IEnumNodeEdge = Nothing       'Interface contenant les Edges d'un noeud.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier la ligne continue.

        Try
            'Indiquer que le Edge a été traité
            If bVisiter Then pTopoEdge.Visited = True

            'Vérifier si on doit initialiser
            If bInitialiser Then
                pLigneEdges = Nothing
                pCollEdges = Nothing
                pTopoNodeFin = Nothing
            End If

            'Si la polyligne des Edges est invalide
            If pLigneEdges Is Nothing Then
                'Créer une polyligne vide
                pLigneEdges = New Polyline
                pLigneEdges.SpatialReference = pTopoEdge.Geometry.SpatialReference
            End If

            'Si la collection des Edges est invalide
            If pCollEdges Is Nothing Then
                'Créer une collection vide
                pCollEdges = New Collection
            End If

            'Ajouter le Edge dans la collection
            pCollEdges.Add(pTopoEdge)

            'Interface pour ajouter la ligne du Edge dans la polyligne
            pGeomColl = CType(pLigneEdges, IGeometryCollection)
            'Ajouter la ligne du Edge dans la polyligne
            pGeomColl.AddGeometryCollection(CType(pTopoEdge.Geometry, IGeometryCollection))

            'Vérifier si plusieurs composantes sont présentes dans la ligne
            If pGeomColl.GeometryCount > 1 Then
                'Simplifier la géométrie
                pTopoOp = CType(pLigneEdges, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
            End If

            'Vérifier si le Edge n'est pas fermé
            If Not pLigneEdges.IsClosed Then
                'Si le noeud de début est différent de celui de départ
                If Not pTopoNodeDebut.Equals(pTopoEdge.FromNode) Then
                    'Si le début du Edge se continu
                    If pTopoEdge.FromNode.Edges(True).Count = 2 Then
                        'Interface pour extraire les Edges du noeud
                        pEnumNodeEdges = pTopoEdge.FromNode.Edges(True)
                        pEnumNodeEdges.Reset()

                        'Extraire la premier Edge
                        pEnumNodeEdges.Next(pTopoNodeEdge, True)

                        'Vérifier si le Edge est le même que celui de départ
                        If pTopoEdge.Equals(pTopoNodeEdge) Then
                            'Extraire la prochain Edge
                            pEnumNodeEdges.Next(pTopoNodeEdge, True)
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)

                            'Si le Edge n'est pas le même que celui de départ
                        Else
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.FromNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)
                        End If

                        'Si le noeud contient une intersection
                    ElseIf pTopoEdge.FromNode.Edges(True).Count > 2 Then
                        'Définir le noeud d'intersection
                        pTopoNodeFin = pTopoEdge.FromNode
                    End If
                End If

                'Si le noeud de fin est différent de celui de départ
                If Not pTopoNodeDebut.Equals(pTopoEdge.ToNode) Then
                    'Si la fin du Edge se continu
                    If pTopoEdge.ToNode.Edges(True).Count = 2 Then
                        'Interface pour extraire les Edges du noeud
                        pEnumNodeEdges = pTopoEdge.ToNode.Edges(True)
                        pEnumNodeEdges.Reset()

                        'Extraire la premier Edge
                        pEnumNodeEdges.Next(pTopoNodeEdge, True)

                        'Vérifier si le Edge est le même que celui de départ
                        If pTopoEdge.Equals(pTopoNodeEdge) Then
                            'Extraire la premier Edge
                            pEnumNodeEdges.Next(pTopoNodeEdge, True)
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.ToNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)

                            'Si le Edge n'est pas le même que celui de départ
                        Else
                            'Définir la ligne des Edges continus
                            Call IdentifierLigneEdge(pTopoEdge.ToNode, pTopoNodeEdge, pLigneEdges, pCollEdges, pTopoNodeFin, bVisiter, False)
                        End If

                        'Si le noeud contient une intersection
                    ElseIf pTopoEdge.ToNode.Edges(True).Count > 2 Then
                        'Définir le noeud d'intersection
                        pTopoNodeFin = pTopoEdge.ToNode
                    End If
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pEnumNodeEdges = Nothing
            pTopoNodeEdge = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de sélectionner tous les Edges d'une lignes continue. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Interface contenant la topologie.</param>
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    ''' 
    Private Sub SelectionnerEdges(ByVal pTopologyGraph As ITopologyGraph, ByVal pCollEdges As Collection)
        'Déclarer les variables de travail
        Dim pEdge As ITopologyEdge = Nothing        'Interface contenant un Edge de la topologie

        Try
            'Traiter tous les Edges de la collection
            For Each pEdge In pCollEdges
                'Sélectionner le Edge
                pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pEdge)
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEdge = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'initialiser l'attribut "Visited=False" pour tous les Edges. 
    '''</summary>
    ''' 
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    '''<param name="bValeur"> Contient la valeur d'initialisation, Faux par défaut.</param>
    ''' 
    Private Sub InitEdgesVisiter(ByVal pCollEdges As Collection, Optional ByVal bValeur As Boolean = False)
        'Déclarer les variables de travail
        Dim pEdge As ITopologyEdge = Nothing        'Interface contenant un Edge de la topologie

        Try
            'Traiter tous les Edges de la collection
            For Each pEdge In pCollEdges
                'Mettre l'attribut Visited=False pour le Edge
                pEdge.Visited = bValeur
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEdge = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions pour corriger la densité des lignes des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la densité des lignes fermées pour les éléments sélectionnés selon une longueur minimale. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour corriger la densité des lignes en trop des géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerDensiteLignes(ByVal dPrecision As Double, ByVal dLongMin As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, _
                                     ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pLigneErreur As IPolyline = Nothing             'Interface contenant une ligne en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.
        Dim iNbErr As Long = -1                             'Contient le nombre d'erreurs.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Traiter qu'il y a des erreurs
                Do While iNbErreurs > iNbErr
                    'Intialisation
                    iNbErr = iNbErreurs

                    'Création de la topologie
                    pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                    pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                    'Initialiser la barre de progression
                    pTrackCancel.Progressor.Message = "Correction de la densité des lignes fermées (NbErr=" & iNbErreurs.ToString & ", LongMin=" & dLongMin.ToString & ") ..."
                    InitBarreProgression(0, pEditor.SelectionCount, pTrackCancel)

                    'Interface pour extraire le premier élément de la sélection
                    pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                    'Réinitialiser la recherche des éléments
                    pEnumFeature.Reset()
                    'Extraire le premier élément de la sélection
                    pFeature = pEnumFeature.Next

                    'Traite tous les éléments sélectionnés
                    Do Until pFeature Is Nothing
                        'Vérifier si les lignes de la géométrie de l'élément ont été corrigées
                        If CorrigerDensiteLignesElement(pTopologyGraph, pFeature, dLongMin, bCorriger, pLigneErreur) Then
                            'Ajouter l'erreur dans le Bag
                            pGeomColl.AddGeometry(pLigneErreur)
                        End If

                        'Vérifier si un Cancel a été effectué
                        If pTrackCancel.Continue = False Then Exit Do

                        'Extraire le prochain élément de la sélection
                        pFeature = pEnumFeature.Next
                    Loop

                    'Retourner le nombre d'erreurs
                    iNbErreurs = pGeomColl.GeometryCount
                Loop

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la densité des lignes fermées")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de densité des lignes fermées (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerDensiteLignes_1", "Longueur de la ligne fermée plus petite que la dimension minimale : " & dLongMin.ToString, _
                                                                  m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pLigneErreur = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger la densité des lignes fermées d'un élément selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger les lignes fermées.</param>
    '''<param name="bCorriger"> Permet d'indiquer si la correction des lignes fermées en trop doit être effectuée.</param>
    '''<param name="pLigneErreur"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerDensiteLignesElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal dLongMin As Double, bCorriger As Boolean, _
                                                  ByRef pLigneErreur As IPolyline) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour enlever la partie de la géométrie en erreur.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerDensiteLignesElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si l'élément est une ligne
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Debug.Print("-")
                    'Debug.Print("OID=" & pFeature.OID)
                    'Créer le Bag vide des lignes en erreur
                    pLigneErreur = New Polyline
                    pLigneErreur.SpatialReference = pFeature.Shape.SpatialReference

                    'Interface pour extraire les Edges de l'élément dans la topologie
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)
                    pEnumTopoEdge.Reset()

                    'Corriger la densité des lignes fermées de la géométrie de l'élément
                    CorrigerDensiteLignesElement = CorrigerDensiteLignesGeometrie(pTopologyGraph, pEnumTopoEdge, dLongMin, pLigneErreur)

                    'Vérifier si on doit corriger les sommets en trop
                    If bCorriger And CorrigerDensiteLignesElement Then
                        'Interface pour extraire les Edges de l'élément dans la topologie
                        pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)
                        'Définir la nouvelle géométrie de l'élément
                        pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                        'Interface pour corriger les éléments sans affecter la topologie
                        pWrite = CType(pFeature.Class, IFeatureClassWrite)

                        'Vérifier si la géométrie est invalide
                        If pGeometry Is Nothing Or pEnumTopoEdge.Count = 0 Then
                            'Détruire l'élément
                            pWrite.RemoveFeature(pFeature)
                            'Indiquer qu'il y a eu une modification
                            CorrigerDensiteLignesElement = True

                            'Si la géométrie est valide
                        Else
                            'Indiquer qu'il y a eu une modification
                            CorrigerDensiteLignesElement = True

                            'Interface pour enlever la partie de la géométrie en erreur
                            pTopoOp = CType(pGeometry, ITopologicalOperator2)
                            'Enlever la partie de la géométrie en erreur
                            pGeometry = pTopoOp.Difference(pLigneErreur)

                            'Si la géométrie est vide
                            If pGeometry.IsEmpty Then
                                'Détruire l'élément
                                pWrite.RemoveFeature(pFeature)

                                'Sinon
                            Else
                                'Traiter le Z et le M
                                Call TraiterZ(pGeometry)
                                Call TraiterM(pGeometry)

                                'Corriger la géométrie de l'élément
                                pFeature.Shape = pGeometry

                                'Sauver la correction
                                pWrite.WriteFeature(pFeature)
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pWrite = Nothing
            pEnumTopoEdge = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de corriger la densité des lignes fermées d'une géométrie selon une longueur minimale 
    ''' et selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion avec l'élément.</param>
    '''<param name="pEnumTopoEdge"> Interface contenant les Edges de la topologie de l'élément à traiter.</param>
    '''<param name="dLongMin"> Contient la longueur minimale utilisée pour identifier et corriger la densité des lignes en trop dans la géométrie d'un élément.</param>
    '''<param name="pLigneErreur"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerDensiteLignesGeometrie(ByVal pTopologyGraph As ITopologyGraph, ByVal pEnumTopoEdge As IEnumTopologyEdge, ByVal dLongMin As Double, _
                                                    ByRef pLigneErreur As IPolyline) As Boolean
        'Déclarer les variables de travail
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un Edge de la topologie. 
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans un Bag.
        Dim pLigneFermee As IPolyline = Nothing             'Interface contenant la ligne fermée des Edges.
        Dim pCollEdges As Collection = Nothing              'Objet contenant une collection de Edges.
        Dim pTopoEdgeLong As ITopologyEdge = Nothing        'Interface contenant le Edge le plus long.

        'Par défaut, une modification a été effectuée
        CorrigerDensiteLignesGeometrie = False

        Try
            'Si les Edges de l'élément sont valident
            If pEnumTopoEdge IsNot Nothing Then
                'Si au moins un Edge est présent
                If pEnumTopoEdge.Count > 0 Then
                    'Interface pour ajouter les lignes en erreur
                    pGeomColl = CType(pLigneErreur, IGeometryCollection)

                    'Extraire la premier Edge
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter tous les Edges
                    Do Until pTopoEdge Is Nothing
                        'Vérifier si le Edge n'est pas sélectionné
                        If Not pTopoEdge.Visited Then
                            'Définir que le Edge a été traité
                            pTopoEdge.Visited = True

                            'Définir la ligne du Edge
                            pLigneFermee = CType(pTopoEdge.Geometry, IPolyline)

                            'Vérifier si la ligne est plus petite que la longueur minimale
                            If pLigneFermee.Length <= dLongMin Then
                                'Vérifier si la ligne est fermée
                                If pLigneFermee.IsClosed Then
                                    'Sélectionner le Edge de la topologie à détruire
                                    pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoEdge)

                                    'Si la ligne n'est pas fermée
                                Else
                                    'Debug.Print(" ")
                                    'Debug.Print("Horaire")
                                    'Debug.Print("Length=" & pLigneFermee.Length)
                                    'Définir la ligne fermée des Edges horaire
                                    Call IdentifierLigneFermee(True, dLongMin, pTopoEdge, pLigneFermee, pCollEdges, pTopoEdgeLong)
                                    'Debug.Print("TotalLength=" & pLigneFermee.Length)

                                    'Vérifier si la ligne fermée est plus petite que la longueur minimale
                                    If pLigneFermee.IsClosed And pLigneFermee.Length <= dLongMin Then
                                        'Indiquer le Edge visité
                                        pTopoEdgeLong.Visited = True
                                        'Sélectionner le Edge de la topologie à détruire
                                        pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoEdgeLong)

                                        'Sinon
                                    Else
                                        'Debug.Print("AntiHoraire")
                                        'Définir la ligne du Edge
                                        pLigneFermee = CType(pTopoEdge.Geometry, IPolyline)
                                        'Debug.Print("Length=" & pLigneFermee.Length)
                                        'Définir la ligne fermée des Edges antihoraire
                                        Call IdentifierLigneFermee(False, dLongMin, pTopoEdge, pLigneFermee, pCollEdges, pTopoEdgeLong)
                                        'Debug.Print("TotalLength=" & pLigneFermee.Length)

                                        'Vérifier si la ligne fermée est plus petite que la longueur minimale
                                        If pLigneFermee.IsClosed And pLigneFermee.Length <= dLongMin Then
                                            'Indiquer le Edge visité
                                            pTopoEdgeLong.Visited = True
                                            'Sélectionner le Edge de la topologie à détruire
                                            pTopologyGraph.Select(esriTopologySelectionResultEnum.esriTopologySelectionResultAdd, pTopoEdgeLong)
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        'Vérifier si on doit détruire le Edge
                        If pTopoEdge.IsSelected Then
                            'Détruire le Edge dans la topologie
                            pTopologyGraph.DeleteEdge(pTopoEdge)

                            'Ajouter la ligne du Edge en erreur
                            pGeomColl.AddGeometryCollection(CType(pTopoEdge.Geometry, IGeometryCollection))

                            'Indiquer qu'une correction a été effectuée
                            CorrigerDensiteLignesGeometrie = True
                        End If

                        'Extraire les Edges suivants
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pTopoEdge = Nothing
            pGeomColl = Nothing
            pLigneFermee = Nothing
            pCollEdges = Nothing
            pTopoEdgeLong = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'identifier tous les Edges d'une lignes fermée à partir d'un Edge de départ. 
    '''</summary>
    ''' 
    '''<param name="bHoraire"> Indique le sens horaire ou antihoraire pour identifier une ligne fermée.</param>
    '''<param name="dLongMin"> Contient la longueur minimale d'une ligne fermée.</param>
    '''<param name="pTopoEdgeDepart"> Interface contenant le Edge de départ.</param>
    '''<param name="pLigneFermee"> Interface contenant la ligne des Edges continus.</param>
    '''<param name="pCollEdges"> Collection des Edges de la ligne continue.</param>
    '''<param name="pTopoEdgeLong"> Interface contenant le Edge d'une ligne le plus long.</param>
    ''' 
    Private Sub IdentifierLigneFermee(ByVal bHoraire As Boolean, ByVal dLongMin As Double, ByVal pTopoEdgeDepart As ITopologyEdge,
                                      ByRef pLigneFermee As IPolyline, ByRef pCollEdges As Collection, ByRef pTopoEdgeLong As ITopologyEdge)
        'Déclarer les variables de travail
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le Edge traité.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant le noeud traité.
        Dim pTopoNodeDepart As ITopologyNode = Nothing      'Interface contenant le noeud de départ.
        Dim pTopoEdgeSuivant As ITopologyEdge = Nothing     'Interface contenant le Edge suivant.
        Dim pTopoNodeSuivant As ITopologyNode = Nothing     'Interface contenant le noeud suivant.
        Dim pLigne As IPolyline = Nothing                   'Interface contenant la ligne du Edge.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des lignes.
        Dim pEnumNodeEdges As IEnumNodeEdge = Nothing       'Interface contenant les Edges d'un noeud.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier la ligne.
        Dim dLongMax As Double = 0                          'Contient la longueur du Edge le plus long.
        Dim dAngle As Double = 0                            'Contient l'angle de la droite traitée. 

        Try
            'Définir le noeud de départ
            pTopoNodeDepart = pTopoEdgeDepart.ToNode
            'Définir le Edge le plus long par défaut
            pTopoEdgeLong = pTopoEdgeDepart
            'Définir la longueur la plus longue
            dLongMax = pLigneFermee.Length

            'Créer une collection vide
            pCollEdges = New Collection
            'Ajouter le Edge dans la collection
            pCollEdges.Add(pTopoEdgeDepart)

            'Définir la ligne du Edge
            pLigneFermee = New Polyline
            pLigneFermee.SpatialReference = pTopoEdgeDepart.Geometry.SpatialReference
            'Interface pour ajouter la ligne du Edge dans la polyligne
            pGeomColl = CType(pLigneFermee, IGeometryCollection)
            'Ajouter la ligne du Edge
            pGeomColl.AddGeometryCollection(CType(pTopoEdgeDepart.Geometry, IGeometryCollection))

            'Définir le Edge et noeud à traiter
            pTopoEdge = pTopoEdgeDepart
            pTopoNode = pTopoNodeDepart

            'Traiter 
            Do
                'Interface pour extraire les Edges du noeud
                Call ExtraireEdgeSuivant(bHoraire, pTopoEdge, pTopoNode, pTopoEdgeSuivant, pTopoNodeSuivant)

                'Sortir si le Edge est déjà traité
                If pTopoEdgeSuivant.Visited Then Exit Do

                'Définir la ligne du Edge
                pLigne = CType(pTopoEdgeSuivant.Geometry, IPolyline)

                'Sortir si la ligne suivante est fermée
                If pLigne.IsClosed Then Exit Do

                'Vérifier si la ligne du Edge est le plus long
                If pLigne.Length > dLongMax Then
                    'Vérifier si aucun des élément du Edge est de type polygon
                    If Not EdgeFeatureGeometryType(pTopoEdgeSuivant, esriGeometryType.esriGeometryPolygon) Then
                        'Définir le Edge le plus long
                        pTopoEdgeLong = pTopoEdgeSuivant
                        'définir la longueur la plus longue
                        dLongMax = pLigne.Length
                    End If
                End If

                'Ajouter le Edge dans la collection
                pCollEdges.Add(pTopoEdgeSuivant)

                'Ajouter la ligne du Edge dans la polyligne
                pGeomColl.AddGeometryCollection(CType(pTopoEdgeSuivant.Geometry, IGeometryCollection))

                'Simplifier la géométrie
                pTopoOp = CType(pLigneFermee, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()

                'Définir le Edge et noeud à traiter
                pTopoEdge = pTopoEdgeSuivant
                pTopoNode = pTopoNodeSuivant
                'Debug.Print("  Length=" & pLigne.Length)

                'Continuer tant que c'est possible
            Loop Until pLigneFermee.IsClosed Or pTopoNodeSuivant.Degree = 1 Or pLigneFermee.Length > dLongMin

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pLigne = Nothing
            pEnumNodeEdges = Nothing
            pTopoNodeDepart = Nothing
            pTopoEdgeSuivant = Nothing
            pTopoNodeSuivant = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet d'indiquer si le type de géométrie d'un des élément du Edge correspond au type spécifié. 
    '''</summary>
    ''' 
    '''<param name="pTopoEdge"> Interface contenant le Edge de la topologie.</param>
    '''<param name="pEsriGeometryType"> Interface contenant le type de géométrie à vérifier.</param>
    ''' 
    '''<returns>Boolean qui indique si le type de géométrie d'un des élément du Edge correspond au type spécifié.</returns>
    ''' 
    Private Function EdgeFeatureGeometryType(ByVal pTopoEdge As ITopologyEdge, ByVal pEsriGeometryType As esriGeometryType) As Boolean
        'Déclarer les variables de travail
        Dim pEnumTopoParent As IEnumTopologyParent = Nothing 'Interface pour extraire les éléments des Edges sélectionnés
        Dim pTopologyParent As esriTopologyParent = Nothing 'Interface contenant les parents des edges
        Dim pFeatureClass As IFeatureClass = Nothing        'Interface contenant la classe de données
        Dim pFeature As IFeature = Nothing                  'Interface contenant un élément

        'Par défaut, le type de la géométrie de l'élément parent ne correspond pas
        EdgeFeatureGeometryType = False

        Try
            'Interface pour extraire les éléments parent du Edge sélectionné
            pEnumTopoParent = pTopoEdge.Parents
            'Initialiser la recherche
            pEnumTopoParent.Reset()

            'Traiter tous les parents du Edge sélectionné
            For j = 1 To pEnumTopoParent.Count
                'Interface pour extraire la classe et le OID d'un élément parent
                pTopologyParent = pEnumTopoParent.Next()

                'Extraire la classe de données
                pFeatureClass = pTopologyParent.m_pFC

                'Extraire l'élément du Edge sélectionné
                pFeature = pFeatureClass.GetFeature(pTopologyParent.m_FID)

                'Vérifier si le type de la géométrie de l'élément parent correspond
                If pFeature.Shape.GeometryType = pEsriGeometryType Then
                    'Indiquer que le type de la géométrie de l'élément parent correspond
                    EdgeFeatureGeometryType = True

                    'Sortir de la boucle
                    Exit For
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEnumTopoParent = Nothing
            pTopologyParent = Nothing
            pFeatureClass = Nothing
            pFeature = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'identifier le Edge suivant à partir d'un sens de traitement et d'un Edge et noeud de début. 
    '''</summary>
    ''' 
    '''<param name="bHoraire"> Indique le sens horaire ou antihoraire pour identifier une ligne fermée.</param>
    '''<param name="pTopoEdgeDebut"> Interface contenant le Edge de début.</param>
    '''<param name="pTopoNodeDebut"> Interface contenant le noeud de début.</param>
    '''<param name="pTopoEdgeSuivant"> Interface contenant le Edge suivant.</param>
    '''<param name="pTopoNodeSuivant"> Interface contenant le noeud suivant.</param>
    ''' 
    Private Sub ExtraireEdgeSuivant(ByVal bHoraire As Boolean, ByVal pTopoEdgeDebut As ITopologyEdge, ByVal pTopoNodeDebut As ITopologyNode,
                                    ByRef pTopoEdgeSuivant As ITopologyEdge, ByRef pTopoNodeSuivant As ITopologyNode)
        'Déclarer les variables de travail
        Dim pLigne As IPolyline = Nothing                   'Interface contenant la ligne du Edge.
        Dim pPointColl As IPointCollection = Nothing        'Interface contenant les points de la ligne.
        Dim pEnumNodeEdges As IEnumNodeEdge = Nothing       'Interface contenant les Edges d'un noeud.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant le prochaine Edge.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant le noeud de départ.
        Dim dAngleDepart As Double = 0                      'Contient l'angle de la droite de départ. 
        Dim dAngleDiff As Double = 0                        'Contient l'angle de différence entre l'angle de la droite de départ et celle traitée. 
        Dim dAngleMin As Double = 360                       'Contient l'angle de différence la plus petite. 
        Dim dAngle As Double = 0                            'Contient l'angle de la droite traitée. 
        Dim bFrom As Boolean = False                        'Indique si le Edge trouvé contient un FromNode.

        Try
            'Par défaut, on retourne le Edge et node suivant sont les mêmes que deux de début
            pTopoEdgeSuivant = pTopoEdgeDebut
            pTopoNodeSuivant = pTopoNodeDebut

            'Interface pour extraire la longueur de la ligne
            pLigne = CType(pTopoEdgeDebut.Geometry, IPolyline)
            'Interface pour extraire les points de la ligne
            pPointColl = CType(pLigne, IPointCollection)
            'Vérifier si le FromNode du Edge est incident au noeud traité.
            If pTopoEdgeDebut.FromNode.Equals(pTopoNodeDebut) Then
                'Définir l'angle de la ligne de début
                dAngleDepart = clsGeneraliserGeometrie.Angle(pPointColl.Point(0), pPointColl.Point(1))
            Else
                'Définir l'angle de la ligne de début
                dAngleDepart = clsGeneraliserGeometrie.Angle(pPointColl.Point(pPointColl.PointCount - 1), pPointColl.Point(pPointColl.PointCount - 2))
            End If
            'Debug.Print("Length=" & pLigne.Length & ", AngleDepart=" & dAngleDepart)

            'Interface pour extraire les Edges du noeud de début
            pEnumNodeEdges = pTopoNodeDebut.Edges(bHoraire)
            'Initialiser la recherche
            pEnumNodeEdges.Reset()
            'Extraire la premier Edge
            pEnumNodeEdges.Next(pTopoEdge, bFrom)

            'Traiter tous les Edges du Noeud de début
            Do Until pTopoEdge Is Nothing
                'Interface pour extraire la longueur de la ligne
                pLigne = CType(pTopoEdge.Geometry, IPolyline)
                'Interface pour extraire les points de la ligne
                pPointColl = CType(pLigne, IPointCollection)

                'Vérifier si le FromNode du Edge est incident au noeud traité.
                If bFrom Then
                    'Définir l'angle de la ligne de début
                    dAngle = clsGeneraliserGeometrie.Angle(pPointColl.Point(0), pPointColl.Point(1))
                Else
                    'Définir l'angle de la ligne de début
                    dAngle = clsGeneraliserGeometrie.Angle(pPointColl.Point(pPointColl.PointCount - 1), pPointColl.Point(pPointColl.PointCount - 2))
                End If

                'Vérifier si c'est le même angle
                If dAngle <> dAngleDepart Then
                    'Vérifier si on extrait dans le sens horaire
                    If bHoraire Then
                        'Vérifier si l'angle est plus grande que celle de début
                        If dAngle > dAngleDepart Then
                            'Définir le nouvel angle
                            dAngle = dAngle - 360
                        End If

                        'Définir la différence d'angle
                        dAngleDiff = dAngleDepart - dAngle

                        'Si on extrait dans le sens antihoraire
                    Else
                        'Vérifier si l'angle est plus grande que celle de début
                        If dAngle < dAngleDepart Then
                            'Définir le nouvel angle
                            dAngle = dAngle + 360
                        End If

                        'Définir la différence d'angle
                        dAngleDiff = dAngle - dAngleDepart
                    End If
                    'Debug.Print("  Length=" & pLigne.Length & ", Angle=" & dAngle & ", AngleDiff=" & dAngleDiff)

                    'Vérifier si la différence d'angle est la plus petite
                    If dAngleDiff < dAngleMin Then
                        'Conserver la différence d'angle minimum
                        dAngleMin = dAngleDiff

                        'Definir le Edge suivant
                        pTopoEdgeSuivant = pTopoEdge

                        'Vérifer si le From Node est incident au noeud traité
                        If bFrom Then
                            'Définir le noeud suivant
                            pTopoNodeSuivant = pTopoEdgeSuivant.ToNode

                            'Sinon
                        Else
                            'Définir le noeud suivant
                            pTopoNodeSuivant = pTopoEdgeSuivant.FromNode
                        End If
                    End If
                End If

                'Extraire la prochain Edge
                pEnumNodeEdges.Next(pTopoEdge, bFrom)
            Loop

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLigne = Nothing
            pPointColl = Nothing
            pEnumNodeEdges = Nothing
            pTopoEdge = Nothing
            pTopoNode = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions pour corriger la superficie des surfaces des éléments"
    '''<summary>
    ''' Routine qui permet de corriger la superficie des surfaces des éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dSupExtMin"> Contient la superficie minimale utilisée pour filtrer les anneaus extérieures et intérieures en trop des géométries des éléments.</param>
    '''<param name="dSupIntMin"> Contient la superficie minimale utilisée pour filtrer les anneaus intérieures et intérieures en trop des géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les surfaces en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerSuperficieSurfaces(ByVal dSupExtMin As Double, ByVal dSupIntMin As Double, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, _
                                          ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les anneaux en trop dans la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la surperficie des surfaces (SupExtMin=" & dSupExtMin.ToString & "/SupIntMin=" & dSupIntMin.ToString & ") ..."

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Vérifier si les anneaux de la surface de l'élément ont été corrigées
                    Call CorrigerSuperficieSurfacesElement(pFeature, dSupExtMin, dSupIntMin, bCorriger, pBagErreurs)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la superficie des surfaces")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de superficie des surfaces (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerSuperficieSurface_2", "Superficie plus petite que la dimension minimale : " & dSupExtMin.ToString & "/" & dSupIntMin.ToString, _
                                                                  m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolygon, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pFeatureLayer = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de corriger la superficie des surfaces d'un élément selon une superficie minimale.
    ''' Les anneaux des surfaces dont leur superficie sont inférieure à la superficie minimale sont éliminés. 
    '''</summary>
    ''' 
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="dSupExtMin"> Contient la superficie minimale utilisée pour identifier et corriger les anneaux extérieures en trop dans la géométrie d'un élément.</param>
    '''<param name="dSupIntMin"> Contient la superficie minimale utilisée pour identifier et corriger les anneaux intérieures en trop dans la géométrie d'un élément.</param>
    '''<param name="bCorriger"> Permet d'indiquer si la correction des anneaux en trop doit être effectuée.</param>
    '''<param name="pBagErreurs"> Interface contenant les erreurs dans la géométrie de l'élément.</param>
    ''' 
    '''<returns>Boolean qui indique si une modification a été effectuée sur la géométrie de l'élément.</returns>
    ''' 
    Private Function CorrigerSuperficieSurfacesElement(ByVal pFeature As IFeature, ByVal dSupExtMin As Double, ByVal dSupIntMin As Double, _
                                                       ByVal bCorriger As Boolean, ByRef pBagErreurs As IGeometryBag) As Boolean
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la nouvelle géométrie de l'élément.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface utilisé pour extraire les anneaux extérieurs et intérieurs.
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux extérieures d'une surface.
        Dim pGeomCollInt As IGeometryCollection = Nothing   'Interface pour extraire les anneaux intérieures d'une surface.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour ajouter les les lignes des anneaux en erreur.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.
        Dim pRingInt As IRing = Nothing                     'Interface contenant l'anneau intérieur.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie d'un anneau.
        Dim pPolygonErreur As IPolygon = Nothing            'Interface contenant une ligne d'un anneau en erreur.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier la géométrie.

        'Par défaut, aucune modification n'a été effectuée
        CorrigerSuperficieSurfacesElement = False

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si l'élément est une surface
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Créer un nouveau polygone vide
                    pGeometry = New Polygon
                    pGeometry.SpatialReference = pBagErreurs.SpatialReference

                    'Interface contenant la géométrie de l'élément
                    pPolygon = CType(pFeature.ShapeCopy, IPolygon4)
                    'Projeter la géométrie
                    pPolygon.Project(pBagErreurs.SpatialReference)
                    'Interface pour extraire les anneaux de la surface
                    pGeomCollExt = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                    'Traiter tous les anneaux de la surface
                    For i = 0 To pGeomCollExt.GeometryCount - 1
                        'Définir l'anneau extérieur
                        pRingExt = CType(pGeomCollExt.Geometry(i), IRing)

                        'Interface pour extraire la superficie de l'anneau
                        pArea = CType(pRingExt, IArea)

                        'Vérifier la superficie est inférieure à la superficie minimale
                        If Math.Abs(pArea.Area) < dSupExtMin Then
                            'Créer un polygone d'erreur vide
                            pPolygonErreur = New Polygon
                            pPolygonErreur.SpatialReference = pFeature.Shape.SpatialReference
                            'Interface pour extraire les anneaux de la surface
                            pGeomCollAdd = CType(pPolygonErreur, IGeometryCollection)
                            'Ajouter l'anneau dans le polygone en erreur
                            pGeomCollAdd.AddGeometry(pRingExt)
                            'Interface pour ajouter les les lignes des anneaux en erreur
                            pGeomCollAdd = CType(pBagErreurs, IGeometryCollection)
                            'Ajouter le polygone en erreur dans le Bag
                            pGeomCollAdd.AddGeometry(pPolygonErreur)

                            'Indiquer qu'une erreur est présente
                            CorrigerSuperficieSurfacesElement = True

                            'Si aucune erreur
                        Else
                            'Interface pour extraire les anneaux de la surface
                            pGeomCollAdd = CType(pGeometry, IGeometryCollection)
                            'Ajouter l'anneau dans le polygone en erreur
                            pGeomCollAdd.AddGeometry(pRingExt)
                        End If

                        'Extraire tous les anneaux intérieurs
                        pGeomCollInt = CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection)

                        'Traiter tous les anneaux intérieurs
                        For j = pGeomCollInt.GeometryCount - 1 To 0 Step -1
                            'Définir l'anneau intérieur
                            pRingInt = CType(pGeomCollInt.Geometry(j), IRing)

                            'Interface pour extraire la superficie
                            pArea = CType(pRingInt, IArea)

                            'Vérifier la superficie est inférieure ;a la superficie minimale
                            If Math.Abs(pArea.Area) < dSupIntMin Then
                                'Créer un polygone d'erreur vide
                                pPolygonErreur = New Polygon
                                pPolygonErreur.SpatialReference = pBagErreurs.SpatialReference
                                'Interface pour extraire les anneaux de la surface
                                pGeomCollAdd = CType(pPolygonErreur, IGeometryCollection)
                                'Ajouter l'anneau dans le polygone en erreur
                                pGeomCollAdd.AddGeometry(pRingInt)
                                'Interface pour ajouter les les lignes des anneaux en erreur
                                pGeomCollAdd = CType(pBagErreurs, IGeometryCollection)
                                'Ajouter le polygone en erreur dans le Bag
                                pGeomCollAdd.AddGeometry(pPolygonErreur)

                                'Indiquer qu'une erreur est présente
                                CorrigerSuperficieSurfacesElement = True

                                'Si aucune erreur
                            Else
                                'Interface pour extraire les anneaux de la surface
                                pGeomCollAdd = CType(pGeometry, IGeometryCollection)
                                'Ajouter l'anneau dans le polygone en erreur
                                pGeomCollAdd.AddGeometry(pRingInt)
                            End If
                        Next
                    Next

                    'Vérifier si on doit corriger les sommets en trop
                    If bCorriger And CorrigerSuperficieSurfacesElement Then
                        'Interface pour corriger les éléments sans affecter la topologie
                        pWrite = CType(pFeature.Class, IFeatureClassWrite)

                        'Vérifier si la géométrie est invalide
                        If pGeometry Is Nothing Then
                            'Détruire l'élément
                            pWrite.RemoveFeature(pFeature)
                            'Indiquer qu'il y a eu une modification
                            CorrigerSuperficieSurfacesElement = True

                            'Si la géométrie est valide
                        Else
                            'Indiquer qu'il y a eu une modification
                            CorrigerSuperficieSurfacesElement = True

                            'Si la géométrie est vide
                            If pPolygon.IsEmpty Then
                                'Détruire l'élément
                                pWrite.RemoveFeature(pFeature)

                                'Sinon
                            Else
                                'Simplifier la géométrie
                                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                                pTopoOp.IsKnownSimple_2 = False
                                pTopoOp.Simplify()

                                'Traiter le Z et le M
                                Call TraiterZ(pGeometry)
                                Call TraiterM(pGeometry)

                                'Corriger la géométrie de l'élément
                                pFeature.Shape = pGeometry

                                'Sauver la correction
                                pWrite.WriteFeature(pFeature)
                            End If
                        End If

                    Else
                        'Indiquer qu'il n'y a pas eu une modification
                        CorrigerSuperficieSurfacesElement = False
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pPolygon = Nothing
            pWrite = Nothing
            pGeomCollExt = Nothing
            pGeomCollInt = Nothing
            pGeomCollAdd = Nothing
            pRingExt = Nothing
            pRingInt = Nothing
            pArea = Nothing
            pPolygonErreur = Nothing
            pTopoOp = Nothing
        End Try
    End Function
#End Region

#Region "Routines et fonctions pour corriger l'adoucissement des éléments"
    '''<summary>
    ''' Routine qui permet de'adoucir les lignes et les limites des polygones pour les éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale entre les sommets des géométries des éléments.</param>
    '''<param name="dDistMin"> Contient la distance minimale entre 2 sommets des géométries des éléments.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub Adoucir(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dDistMin As Double, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Adoucir les lignes et les limites des polygones (DistMin=" & dDistMin.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Adoucir l'élément selon une distance minimum entre 2 sommets
                    Call AdoucirElement(pTopologyGraph, pFeature, dDistLat, dDistMin)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Terminer l'opération UnDo
                pEditor.StopOperation("Adoucir les lignes et les limites des polygones")

                'Enlever la sélection
                m_MxDocument.FocusMap.ClearSelection()

                'Sélectionner les éléments en erreur
                'm_MxDocument.FocusMap.SelectByShape(pBagErreurs, Nothing, False)

                'Rafraîchir l'affichage
                m_MxDocument.ActiveView.Refresh()

                'Sinon
            Else
                'Annuler l'opération UnDo
                pEditor.AbortOperation()
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet d'adoucir la ligne ou la limite du polygone d'une élément selon la topologie des éléments afin de conserver la connexion entre les éléments. 
    '''</summary>
    ''' 
    '''<param name="pTopologyGraph"> Contient la topologie des éléments à traiter afin de conserver la connexion entre les éléments.</param>
    '''<param name="pFeature"> Contient l'élément à traiter.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale entre les sommets des géométries des éléments.</param>
    '''<param name="dDistMin"> Contient la distance minimum utilisée pour adoucir les géométries des éléments.</param>
    ''' 
    Public Sub AdoucirElement(ByVal pTopologyGraph As ITopologyGraph, ByVal pFeature As IFeature, ByVal dDistLat As Double, ByVal dDistMin As Double)
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une polyligne.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie d'un élément
        Dim pWrite As IFeatureClassWrite = Nothing          'Interface utilisé pour écrire un élément.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing    'Interface utilisé pour extraire les edges de la topologie.
        Dim pTopoEdge As ITopologyEdge = Nothing            'Interface contenant un edge de la topologie. 
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne. 
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire les points en erreur.

        Try
            'Si l'élément est valide
            If pFeature IsNot Nothing Then
                'Vérifier si la géométrie n'est pas de type point
                If pFeature.Shape.GeometryType <> esriGeometryType.esriGeometryPoint Then
                    'Interface pour extraire les composantes
                    pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Table, IFeatureClass), pFeature.OID)

                    'Extraire la première composante
                    pTopoEdge = pEnumTopoEdge.Next

                    'Traiter toutes les composantes
                    Do Until pTopoEdge Is Nothing
                        'Vérifier si le Edge est une Polyline
                        If TypeOf (pTopoEdge.Geometry) Is IPolyline Then
                            'Interface pour extraire la ligne à traiter
                            pGeomColl = CType(pTopoEdge.Geometry, IGeometryCollection)

                            'Extraire la ligne à traiter
                            pPath = CType(pGeomColl.Geometry(0), IPath)
                        Else
                            'Définir la ligne à traiter
                            pPath = CType(pTopoEdge.Geometry, IPath)
                        End If

                        'Adoucir la géométrie du TopologyEdge selon la distance minimum
                        pPath = TraiterCentreDroites(pPath, dDistMin)

                        'Créer une nouvelle polyligne vide
                        pPolyline = New Polyline
                        pPolyline.SpatialReference = pTopoEdge.Geometry.SpatialReference
                        'Interface pour ajouter la ligne
                        pGeomColl = CType(pPolyline, IGeometryCollection)
                        'Ajouter la ligne
                        pGeomColl.AddGeometry(pPath)
                        'Filtrer les sommets selon la distance latérale
                        pPolyline.Generalize(dDistLat)
                        'Redefinir la ligne 
                        pPath = CType(pGeomColl.Geometry(0), IPath)

                        'Mettre à jour la géométrie dans la topologie
                        pTopologyGraph.SetEdgeGeometry(pTopoEdge, pPath)

                        'Extraire la première composante
                        pTopoEdge = pEnumTopoEdge.Next
                    Loop

                    'Définir la nouvelle géométrie de l'élément
                    pGeometry = pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID)

                    'Interface pour corriger les éléments sans affecter la topologie
                    pWrite = CType(pFeature.Class, IFeatureClassWrite)

                    'Vérifier si la géométrie est invalide
                    If pGeometry Is Nothing Then
                        'Détruire l'élément
                        pWrite.RemoveFeature(pFeature)

                        'Si la géométrie est valide
                    Else
                        'Traiter le Z et le M
                        Call TraiterZ(pGeometry)
                        Call TraiterM(pGeometry)

                        'Corriger la géométrie de l'élément
                        pFeature.Shape = pGeometry

                        'Sauver la correction
                        pWrite.WriteFeature(pFeature)
                    End If
                End If
            End If

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pGeometry = Nothing
            pWrite = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Function qui permet de créer une nouvelle ligne à partir des centres des droites de la ligne d'entrée.
    ''' </summary>
    ''' 
    ''' <param name="pPath">Interface contenant une ligne.</param>
    ''' <param name="dDistMin">Contient la distance minimum entre 2 sommets d'une ligne.</param>
    ''' 
    ''' <returns>IPath contenant la nouvelle ligne traitée.</returns>
    ''' 
    Private Function TraiterCentreDroites(ByVal pPath As IPath, ByVal dDistMin As Double) As IPath
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing            'Interface pour densifier la ligne.
        Dim pPointColl As IPointCollection = Nothing    'Interface pour extraire les sommets de la ligne.
        Dim pPointCollAdd As IPointCollection = Nothing 'Interface pour ajouter les sommets de la nouvelle ligne.
        Dim pPoint As IPoint = Nothing                  'Interface contenant le nouveau point.
        Dim pPointA As IPoint = Nothing                 'Interface contenant le point précédent.
        Dim pPointB As IPoint = Nothing                 'Interface contenant le point suivant.

        'Définir la valeur par défaut
        TraiterCentreDroites = New ESRI.ArcGIS.Geometry.Path
        TraiterCentreDroites.SpatialReference = pPath.SpatialReference

        Try
            'Créer la ligne vide
            pPolyline = New Polyline
            pPolyline.SpatialReference = pPath.SpatialReference

            'Interface pour ajouter les sommets
            pPointCollAdd = CType(pPolyline, IPointCollection)
            'Ajouter les sommets à la polyligne
            pPointCollAdd.AddPointCollection(CType(pPath, IPointCollection))

            'Densifier la ligne
            pPolyline.Densify(dDistMin, 0)

            'Interface pour extraire les sommets de la ligne
            pPointColl = CType(pPolyline, IPointCollection)

            'Interface pour ajouter les sommets
            pPointCollAdd = CType(TraiterCentreDroites, IPointCollection)

            'Vérifier si seulement 2 sommets
            If pPointColl.PointCount = 2 Then
                'Ajouter la ligne originale
                pPointCollAdd.AddPointCollection(CType(pPath, IPointCollection))

                'Si plus de deux sommets
            Else
                'Définir le premier sommet
                pPointA = pPointColl.Point(0)
                'Ajouter le premier sommet
                pPointCollAdd.AddPoint(pPointA)

                'Traiter tous les sommets
                For j = 1 To pPointColl.PointCount - 1
                    'Définir le prochain sommet
                    pPointA = pPointColl.Point(j - 1)
                    pPointB = pPointColl.Point(j)

                    'Définir le centre de la droite
                    pPoint = New Point
                    pPoint.SpatialReference = pPath.SpatialReference
                    pPoint.X = (pPointA.X + pPointB.X) / 2
                    pPoint.Y = (pPointA.Y + pPointB.Y) / 2

                    'Ajouter le premier sommet
                    pPointCollAdd.AddPoint(pPoint)
                Next

                'Ajouter le dernier sommet
                pPointCollAdd.AddPoint(pPointB)
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pPoint = Nothing
            pPointA = Nothing
            pPointB = Nothing
        End Try
    End Function
#End Region

#Region "Routines et fonctions pour corriger la généralisation gauche/droite des éléments de type ligne"
    '''<summary>
    ''' Routine qui permet de corriger la généralisation des éléments de type ligne sélectionnés fractionnées. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargGenMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongGenMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur minimale des lignes.</param>
    '''<param name="pFeatureLayerGeneraliser"> Contient le FeatureLayer dans lequel les éléments généralisés seront créés.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerGeneralisationLignesFractionnees(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargGenMin As Double, ByVal dLongGenMin As Double,
                                                       ByVal dLongMin As Double, ByVal pFeatureLayerGeneraliser As IFeatureLayer, ByVal bCorriger As Boolean,
                                                       ByVal bCreerFichierErreurs As Boolean, ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la généralisation des lignes fractionnées (LargGenMin=" & dLargGenMin.ToString & ", LongGenMin=" & dLongGenMin.ToString & ", LongMin=" & dLongMin.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Corriger la généralisation de l'élément de type ligne fragmentée
                    Call GeneraliserLigneFractionneeElement(dPrecision, dDistLat, dLargGenMin, dLongGenMin, dLongMin, pFeatureLayerGeneraliser, bCorriger,
                                                            pTopologyGraph, pFeature, pBagErreurs)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la généralisation des lignes fractionnées")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()
                        'Sélectionner les éléments en erreur
                        'm_MxDocument.FocusMap.SelectByShape(pBagErreurs, Nothing, False)

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de généralisation des lignes fractionnées (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerGeneralisationLignes_1", "Erreur de généralisation de ligne fragmentées : Larg=" _
                                                                 & dLongMin.ToString & ", LongMin=" & dLongMin.ToString,
                                                                 m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de corriger la généralisation gauche/droite des éléments de type ligne sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargGenMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongGenMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur minimale des lignes.</param>
    '''<param name="pFeatureLayerGeneraliser"> Contient le FeatureLayer dans lequel les éléments généralisés seront créés.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerGeneralisationLignes(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargGenMin As Double, ByVal dLongGenMin As Double, ByVal dLongMin As Double,
                                            ByVal pFeatureLayerGeneraliser As IFeatureLayer, ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean,
                                            ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la généralisation des lignes (LargGenMin=" & dLargGenMin.ToString & ", LongGenMin=" & dLongGenMin.ToString & ", LongMin=" & dLongMin.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Corriger la généralisation de l'élément de type ligne
                    Call GeneraliserLigneElement(dPrecision, dDistLat, dLargGenMin, dLongGenMin, dLongMin, pFeatureLayerGeneraliser, bCorriger,
                                                 pTopologyGraph, pFeature, pBagErreurs)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la généralisation des lignes")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()
                        'Sélectionner les éléments en erreur
                        'm_MxDocument.FocusMap.SelectByShape(pBagErreurs, Nothing, False)

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de généralisation des lignes (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerGeneralisationLignes_1", "Erreur de généralisation de ligne : Larg=" _
                                                                 & dLongMin.ToString & ", LongMin=" & dLongMin.ToString,
                                                                 m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de généraliser un élément de type ligne fractionnée et selon une largeur et une longueur minimum de généralisation et une longueur minimale des lignes. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargGenMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongGenMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur minimale des lignes.</param>
    '''<param name="pFeatureLayerGeneraliser"> Contient le FeatureLayer dans lequel les éléments généralisés seront créés.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie des éléments à généraliser.</param>
    '''<param name="pFeature"> Interface contenant l'élément à généraliser.</param>
    '''<param name="pBagErreurs"> Contient les géométries d'erreurs.</param>
    ''' 
    Public Sub GeneraliserLigneFractionneeElement(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargGenMin As Double, ByVal dLongGenMin As Double, ByVal dLongMin As Double,
                                                  ByVal pFeatureLayerGeneraliser As IFeatureLayer, ByVal bCorriger As Boolean,
                                                  ByVal pTopologyGraph As ITopologyGraph, ByRef pFeature As IFeature, ByRef pBagErreurs As IGeometryBag)
        'Déclarer les variables de travail
        Dim pNewFeature As IFeature = Nothing               'Interface ESRI contenant un élément en sélection.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la polyligne à traiter.
        Dim pLigneFractionnee As IPolyline = Nothing        'Interface contenant la polyligne pfractionnée à traiter.
        Dim pPolylineGen As IPolyline = Nothing             'Interface contenant la polyligne de généralisation.
        Dim pPolylineErrTmp As IPolyline = Nothing          'Interface contenant la polyligne d'erreurs de généralisation temporaire.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant la polyligne d'erreurs de généralisation.
        Dim pSquelette As IPolyline = Nothing               'Interface contenant le squelette.
        Dim pSqueletteEnv As IPolyline = Nothing            'Interface contenant le squelette avec son enveloppe.
        Dim pSqueletteTmp As IPolyline = Nothing            'Interface contenant le squelette temporaire.
        Dim pBagDroitesTmp As IGeometryBag = Nothing        'Interface contenant les droites de Delaunay temporaire.
        Dim pBagDroites As IGeometryBag = Nothing           'Interface contenant les droites de Delaunay.
        Dim pBagDroitesEnv As IGeometryBag = Nothing        'Interface contenant les droites de Delaunay avec son enveloppe.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pPointsConnexion As IMultipoint = Nothing       'Interface contenant les points de connexion.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier une géométrie.

        Try
            'Traite tous les éléments sélectionnés
            If pFeature IsNot Nothing Then
                'vérifier si la géométrie de l'élément est de type ligne
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Définir le polygone à généraliser
                    'pPolyline = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID), IPolyline)
                    pPolyline = CType(pFeature.ShapeCopy, IPolyline)
                    'Projeter la géométrie
                    pPolyline.Project(pBagErreurs.SpatialReference)

                    'Extraire les points de connexion
                    pTopoOp = CType(pPolyline, ITopologicalOperator2)
                    pPointsConnexion = CType(pTopoOp.Boundary, IMultipoint)

                    'Fractionner une polyligne
                    pLigneFractionnee = FractionnerPolyligne(pPolyline)

                    'Généraliser la polyligne à droite
                    Call clsGeneraliserGeometrie.GeneraliserLigne(pLigneFractionnee, pPointsConnexion, dDistLat, dLargGenMin, dLongGenMin, dLongMin,
                                                                  pPolylineGen, pPolylineErr, pSquelette, pSqueletteEnv, pBagDroites, pBagDroitesEnv)

                    'Vérifier si une erreur de généralisation est présente
                    If bCorriger And Not pPolylineErr.IsEmpty Then
                        'Interface pour extraire le nombre d'erreurs
                        pGeomColl = CType(pBagErreurs, IGeometryCollection)
                        'Ajouter l'erreur dans le Bag
                        pGeomColl.AddGeometry(pPolylineErr)

                        'Vérifier si l'élément doit être détruit
                        If pPolylineGen.IsEmpty Then
                            'Détruire l'élément
                            pFeature.Delete()

                            'Si la géométrie doit être modifiée
                        Else
                            'Interface pour simplifier une géométrie.
                            pTopoOp = CType(pPolylineGen, ITopologicalOperator2)
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()
                            'Modifié la géométrie
                            pFeature.Shape = pPolylineGen
                            'Sauver la modification
                            pFeature.Store()
                        End If

                        ''Vérfier si la classe de géneralisation est spécifié
                        'If pFeatureLayerGeneraliser IsNot Nothing Then
                        '    'Créer l'élément généralisé
                        '    pNewFeature = pFeatureLayerGeneraliser.FeatureClass.CreateFeature
                        '    'Définir la géométrie de l'élément généralisé
                        '    pNewFeature.Shape = pPolylineErr
                        '    'Sauver l'élément
                        '    pNewFeature.Store()
                        'End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pNewFeature = Nothing
            pPolyline = Nothing
            pLigneFractionnee = Nothing
            pPolylineGen = Nothing
            pPolylineErr = Nothing
            pSquelette = Nothing
            pBagDroites = Nothing
            pPolylineErrTmp = Nothing
            pSqueletteTmp = Nothing
            pBagDroitesTmp = Nothing
            pGeomColl = Nothing
            pPointsConnexion = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de fractionner une polyligne selon les angles aigus et obtus.
    ''' La polyligne résultante contiendra des lignes (Path) de sens inversé entre les changement d'angle aigu et obtu des droites consécutives.
    '''</summary>
    '''
    '''<param name="pPolyline"> Interface contenant la polyligne à fractionner.</param>
    ''' 
    ''' <returns>IPolyline contenant la polyligne fractionnée.</returns>
    ''' 
    Public Function FractionnerPolyligne(ByVal pPolyline As IPolyline) As IPolyline
        'Déclarer les variables de travail
        Dim pLigneFractionnee As IPolyline = Nothing        'Interface contenant la ligne fractionnée.
        Dim pPath As IPath = Nothing                        'Interface contenant une partie de la ligne fractionnée.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour ajouter les lignes.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments.
        Dim pSegCollAdd As ISegmentCollection = Nothing     'Interface pour ajouter les segments.
        Dim pSegment As ISegment = Nothing                  'Interface contenant un segment.
        Dim pDemiSegment As ISegment = Nothing              'Interface contenant un demi segment.
        Dim pPoint As Point = Nothing                       'Interface contenant un point.
        Dim dAngleDepart As Double = -1                     'Contient l'angle de départ.
        Dim dAngle As Double = -1                           'Contient l'angle de la droite traitée.
        Dim dDiffDepart As Double = -1                      'Contient la différence d'angle entre deux droites consécutives de départ.
        Dim dDiff As Double = -1                            'Contient la différence d'angle entre deux droites consécutives.

        'Par defaut, la ligne fractionnée est la même que la ligne à traiter
        FractionnerPolyligne = pPolyline

        Try
            'Sortir si la polyligne est vide
            If pPolyline.IsEmpty Then Exit Function

            'Créer une ligne fractionnée vide
            pLigneFractionnee = New Polyline
            pLigneFractionnee.SpatialReference = pPolyline.SpatialReference
            'Interface pour ajouter une ligne
            pGeomCollAdd = CType(pLigneFractionnee, IGeometryCollection)

            'Interface pour extraire les composantes
            pGeomColl = CType(pPolyline, IGeometryCollection)
            'Traiter toutes les composantes
            For i = 0 To pGeomColl.GeometryCount - 1
                'Initialiser les angles
                dAngleDepart = -1
                dAngle = -1

                'Interface pour extraire les segments
                pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)
                'Traiter tous les segments de la composante
                For j = 1 To pSegColl.SegmentCount - 1
                    'Définir le segment à traiter
                    pSegment = pSegColl.Segment(j)

                    'Vérifier si l'angle de départ n'est pas initialisée
                    If dAngleDepart = -1 Then
                        'Initialiser l'angle de départ
                        dAngleDepart = clsGeneraliserGeometrie.Angle(pSegColl.Segment(j - 1).FromPoint, pSegColl.Segment(j - 1).ToPoint)
                        'Initialiser l'angle selon l'angle de départ
                        dAngle = clsGeneraliserGeometrie.Angle(pSegColl.Segment(j).FromPoint, pSegColl.Segment(j).ToPoint)

                        'Définir la différence d'angle
                        dDiff = dAngleDepart - dAngle
                        'Définir l'angle entre les deux droites consécutives
                        If dDiff > -180 And dDiff < 180 Then
                            dDiff = 180 - dDiff
                        ElseIf dDiff < -180 Then
                            dDiff = Math.Abs(180 + dDiff)
                        ElseIf dDiff > 180 Then
                            dDiff = 540 - dDiff
                        End If
                        'Debug.Print(dAngleDepart.ToString + "-" + dAngle.ToString + "=" + dDiff.ToString + ">" + dDiff.ToString)

                        'Créer une partie de ligne fractionnée vide
                        pPath = New Path
                        pPath.SpatialReference = pPolyline.SpatialReference

                        'Interface pour ajouter un segment
                        pSegCollAdd = CType(pPath, ISegmentCollection)
                        'Ajouter un segment
                        pSegCollAdd.AddSegment(pSegColl.Segment(j - 1))

                        'Si l'angle de départ est initialisée
                    Else
                        'Initialiser l'angle selon l'angle de départ
                        dAngle = clsGeneraliserGeometrie.Angle(pSegColl.Segment(j).FromPoint, pSegColl.Segment(j).ToPoint)

                        'Définir la différence d'angle
                        dDiff = dAngleDepart - dAngle
                        'Définir l'angle entre les deux droites consécutives
                        If dDiff > -180 And dDiff < 180 Then
                            dDiff = 180 - dDiff
                        ElseIf dDiff < -180 Then
                            dDiff = Math.Abs(180 + dDiff)
                        ElseIf dDiff > 180 Then
                            dDiff = 540 - dDiff
                        End If
                        'Debug.Print(dAngleDepart.ToString + "-" + dAngle.ToString + "=" + dDiff.ToString + ">" + dDiff.ToString)

                        'Vérifier si l'état (aigu ou obtu) de la différence d'angle est différent
                        If (dDiff < 180 And dDiffDepart > 180) Or (dDiff > 180 And dDiffDepart < 180) Then
                            'Définir le demi segment vide
                            pDemiSegment = New Line
                            pDemiSegment.SpatialReference = pPolyline.SpatialReference
                            'Définir le point vide
                            pPoint = New Point
                            pPoint.SpatialReference = pPolyline.SpatialReference
                            'Calculer la position du centre du segment
                            pPoint.X = (pSegColl.Segment(j - 1).FromPoint.X + pSegColl.Segment(j - 1).ToPoint.X) / 2
                            pPoint.Y = (pSegColl.Segment(j - 1).FromPoint.Y + pSegColl.Segment(j - 1).ToPoint.Y) / 2
                            'Définir le premier point
                            pDemiSegment.FromPoint = pSegColl.Segment(j - 1).FromPoint
                            'Définir le deuxième point
                            pDemiSegment.ToPoint = pPoint
                            'Ajouter un demi segment
                            pSegCollAdd.AddSegment(pDemiSegment)
                            'Changer l'orientation si l'angle est obtu
                            If dDiffDepart > 180 Then pPath.ReverseOrientation()
                            'Ajouter une ligne fractionnée
                            pGeomCollAdd.AddGeometry(pPath)

                            'Créer une partie de ligne fractionnée vide
                            pPath = New Path
                            pPath.SpatialReference = pPolyline.SpatialReference
                            'Définir le demi segment vide
                            pDemiSegment = New Line
                            pDemiSegment.SpatialReference = pPolyline.SpatialReference
                            'Définir le point vide
                            pPoint = New Point
                            pPoint.SpatialReference = pPolyline.SpatialReference
                            'Calculer la position du centre du segment
                            pPoint.X = (pSegColl.Segment(j - 1).FromPoint.X + pSegColl.Segment(j - 1).ToPoint.X) / 2
                            pPoint.Y = (pSegColl.Segment(j - 1).FromPoint.Y + pSegColl.Segment(j - 1).ToPoint.Y) / 2
                            'Interface pour ajouter un segment
                            pSegCollAdd = CType(pPath, ISegmentCollection)
                            'Définir le premier point
                            pDemiSegment.FromPoint = pPoint
                            'Définir le premier point
                            pDemiSegment.ToPoint = pSegColl.Segment(j - 1).ToPoint
                            'Ajouter un demi segment
                            pSegCollAdd.AddSegment(pDemiSegment)

                            'Si l'état (aigu ou obtu) de la différence d'angle est le même
                        Else
                            'Ajouter un segment
                            pSegCollAdd.AddSegment(pSegColl.Segment(j - 1))
                        End If
                    End If

                    'Initialiser l'angle de départ
                    dAngleDepart = dAngle
                    'Initialiser l'angle de différence de départ
                    dDiffDepart = dDiff
                Next
            Next

            'Ajouter un segment
            pSegCollAdd.AddSegment(pSegment)
            'Changer l'orientation si l'angle est obtu
            If dDiffDepart > 180 Then pPath.ReverseOrientation()
            'Ajouter une ligne fractionnée
            pGeomCollAdd.AddGeometry(pPath)

            'Retourner la polyligne fractionnée
            FractionnerPolyligne = pLigneFractionnee

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLigneFractionnee = Nothing
            pPath = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pSegColl = Nothing
            pSegCollAdd = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de généraliser un élément de type ligne selon une largeur et une longueur minimum de généralisation et une longueur minimale des lignes. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargGenMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongGenMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur minimale des lignes.</param>
    '''<param name="pFeatureLayerGeneraliser"> Contient le FeatureLayer dans lequel les éléments généralisés seront créés.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="pTopologyGraph"> Interface contenant la topologie des éléments à généraliser.</param>
    '''<param name="pFeature"> Interface contenant l'élément à généraliser.</param>
    '''<param name="pBagErreurs"> Contient les géométries d'erreurs.</param>
    ''' 
    Public Sub GeneraliserLigneElement(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargGenMin As Double, ByVal dLongGenMin As Double, ByVal dLongMin As Double,
                                       ByVal pFeatureLayerGeneraliser As IFeatureLayer, ByVal bCorriger As Boolean,
                                       ByVal pTopologyGraph As ITopologyGraph, ByRef pFeature As IFeature, ByRef pBagErreurs As IGeometryBag)
        'Déclarer les variables de travail
        Dim pNewFeature As IFeature = Nothing               'Interface ESRI contenant un élément en sélection.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant le super polygone à traiter.
        Dim pPolylineGen As IPolyline = Nothing             'Interface contenant le polygone de généralisation.
        Dim pPolylineErrTmp As IPolyline = Nothing          'Interface contenant la polyligne d'erreur de généralisation temporaire.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant la polyligne d'erreur de généralisation.
        Dim pSquelette As IPolyline = Nothing               'Interface contenant le squelette.
        Dim pSqueletteEnv As IPolyline = Nothing            'Interface contenant le squelette avec son enveloppe.
        Dim pSqueletteTmp As IPolyline = Nothing            'Interface contenant le squelette temporaire.
        Dim pBagDroitesTmp As IGeometryBag = Nothing        'Interface contenant les droites de Delaunay temporaire.
        Dim pBagDroites As IGeometryBag = Nothing           'Interface contenant les droites de Delaunay.
        Dim pBagDroitesEnv As IGeometryBag = Nothing        'Interface contenant les droites de Delaunay avec son enveloppe.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pPointsConnexion As IMultipoint = Nothing       'Interface contenant les points de connexion.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier une géométrie.

        Try
            'Traite tous les éléments sélectionnés
            If pFeature IsNot Nothing Then
                'vérifier si la géométrie de l'élément est de type ligne
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                    'Définir le polygone à généraliser
                    'pPolyline = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID), IPolyline)
                    pPolyline = CType(pFeature.ShapeCopy, IPolyline)
                    'Projeter la géométrie
                    pPolyline.Project(pBagErreurs.SpatialReference)

                    'Extraire les points de connexion
                    pTopoOp = CType(pPolyline, ITopologicalOperator2)
                    pPointsConnexion = CType(pTopoOp.Boundary, IMultipoint)

                    'Généraliser la polyligne à droite
                    Call clsGeneraliserGeometrie.GeneraliserPolyligne(pPolyline, pPointsConnexion, dDistLat, dLargGenMin, dLongGenMin, dLongMin,
                                                                      pPolylineGen, pPolylineErr, pSquelette, pSqueletteEnv, pBagDroites, pBagDroitesEnv)

                    'Si le résultat de la ligne généralisée est vide
                    If pPolylineGen.IsEmpty Then
                        'Inverser le sens de numérisation
                        pPolyline.ReverseOrientation()
                        'Généraliser la polyligne à droite
                        Call clsGeneraliserGeometrie.GeneraliserPolyligne(pPolyline, pPointsConnexion, dDistLat, dLargGenMin, dLongGenMin, dLongMin,
                                                                          pPolylineGen, pPolylineErr, pSquelette, pSqueletteEnv, pBagDroites, pBagDroitesEnv)
                    End If

                    'Définir la ligne pour généraliser dans l'autre sens
                    pPolylineGen.ReverseOrientation()
                    pPolyline = pPolylineGen

                    'Généraliser la polyligne à gauche
                    Call clsGeneraliserGeometrie.GeneraliserPolyligne(pPolyline, pPointsConnexion, dDistLat, dLargGenMin, dLongGenMin, dLongMin,
                                                                      pPolylineGen, pPolylineErrTmp, pSqueletteTmp, pSqueletteEnv, pBagDroitesTmp, pBagDroitesEnv)

                    'Interface pour ajouter les erreurs de généralisation
                    pGeomColl = CType(pPolylineErr, IGeometryCollection)
                    'Ajouter les erreurs de généralisation
                    pGeomColl.AddGeometryCollection(CType(pPolylineErrTmp, IGeometryCollection))

                    'Vérifier si une erreur de généralisation est présente
                    If bCorriger And Not pPolylineErr.IsEmpty Then
                        'Interface pour extraire le nombre d'erreurs
                        pGeomColl = CType(pBagErreurs, IGeometryCollection)
                        'Ajouter l'erreur dans le Bag
                        pGeomColl.AddGeometry(pPolylineErr)

                        'Vérifier si l'élément doit être détruit
                        If pPolylineGen.IsEmpty Then
                            'Détruire l'élément
                            pFeature.Delete()

                            'Si la géométrie doit être modifiée
                        Else
                            'Interface pour simplifier une géométrie.
                            pTopoOp = CType(pPolylineGen, ITopologicalOperator2)
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()
                            'Modifié la géométrie
                            pFeature.Shape = pPolylineGen
                            'Sauver la modification
                            pFeature.Store()
                        End If

                        ''Vérfier si la classe de géneralisation est spécifié
                        'If pFeatureLayerGeneraliser IsNot Nothing Then
                        '    'Créer l'élément généralisé
                        '    pNewFeature = pFeatureLayerGeneraliser.FeatureClass.CreateFeature
                        '    'Définir la géométrie de l'élément généralisé
                        '    pNewFeature.Shape = pPolylineErr
                        '    'Sauver l'élément
                        '    pNewFeature.Store()
                        'End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pNewFeature = Nothing
            pPolyline = Nothing
            pPolylineGen = Nothing
            pPolylineErr = Nothing
            pSquelette = Nothing
            pBagDroites = Nothing
            pPolylineErrTmp = Nothing
            pSqueletteTmp = Nothing
            pBagDroitesTmp = Nothing
            pGeomColl = Nothing
            pPointsConnexion = Nothing
            pTopoOp = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions pour corriger la généralisation extérieure des éléments de type surface"
    '''<summary>
    ''' Routine qui permet de corriger la généralisation extérieure des éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dSupMin"> Contient la superficie de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerGeneralisationExterieure(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                                ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la généralisation extérieure (LargMin=" & dLargMin.ToString & ", LongMin=" & dLongMin.ToString & ", SupMin=" & dSupMin.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Définir le polygone à généraliser
                    Call GeneraliserExterieureElement(dPrecision, dDistLat, dLargMin, dLongMin, dSupMin, bCorriger,
                                                      pTopologyGraph, pFeature, pBagErreurs)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la généralisation extérieure")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()
                        'Sélectionner les éléments en erreur
                        'm_MxDocument.FocusMap.SelectByShape(pBagErreurs, Nothing, False)

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de généralisation extérieure (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerGeneralisationExterieure_1", "Erreur de généralisation extérieure : Larg=" _
                                                                 & dLongMin.ToString & ", LongMin=" & dLongMin.ToString & ", SupMin=" & dSupMin.ToString, _
                                                                 m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de généraliser l'extérieure d'un élément de type surface selon une largeur et une longueur minimum de généralisation et une superficie minimale des surfaces. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dSupMin"> Contient la superficie de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="pFeature"> Interface contenant l'élément à généraliser.</param>
    '''<param name="pBagErreurs"> Contient les géométries d'erreurs.</param>
    ''' 
    Public Sub GeneraliserExterieureElement(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                            ByVal bCorriger As Boolean, ByVal pTopologyGraph As ITopologyGraph, ByRef pFeature As IFeature, ByRef pBagErreurs As IGeometryBag)
        'Déclarer les variables de travail
        Dim pNewFeature As IFeature = Nothing               'Interface ESRI contenant un élément en sélection.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface contenant le super polygone à traiter.
        Dim pPolygonGen As IPolygon = Nothing               'Interface contenant le polygone de généralisation.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant la polyligne d'erreur de généralisation.
        Dim pSquelette As IPolyline = Nothing               'Interface contenant le squelette du polygone.
        Dim pBagDroites As IGeometryBag = Nothing           'Interface contenant les droites de Delaunay.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les points de connexion.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing    'Interface pour extraire les points de connexion.
        Dim pTopoNode As ITopologyNode = Nothing            'Interface contenant un point de connexion.
        Dim pPointsConnexion As IMultipoint = Nothing       'Interface contenant les points de connexion.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier une géométrie.

        Try
            'Traite tous les éléments sélectionnés
            If pFeature IsNot Nothing Then
                'vérifier si la géométrie de l'élément est de type surface
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Définir le polygone à généraliser
                    'pPolygon = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID), IPolygon4)
                    pPolygon = CType(pFeature.ShapeCopy, IPolygon4)
                    'Projeter la géométrie
                    pPolygon.Project(pBagErreurs.SpatialReference)

                    ''Interface pour simplifier une géométrie.
                    'pTopoOp = CType(pPolygon, ITopologicalOperator2)
                    'pTopoOp.IsKnownSimple_2 = False
                    'pTopoOp.Simplify()

                    'Extraire les points de connexion
                    pPointsConnexion = clsGeneraliserGeometrie.ExtrairePointsIntersection(pFeature, pTopologyGraph)

                    'Généraliser l'extérieur du polygone
                    Call clsGeneraliserGeometrie.GeneraliserExterieurPolygone(pPolygon, pPointsConnexion, dDistLat, dLargMin, dLongMin, dSupMin,
                                                                              pPolygonGen, pPolylineErr, pSquelette, pBagDroites)

                    'Vérifier si une erreur de généralisation est présente
                    If bCorriger And Not pPolylineErr.IsEmpty Then
                        'Interface pour extraire le nombre d'erreurs
                        pGeomColl = CType(pBagErreurs, IGeometryCollection)
                        'Ajouter l'erreur dans le Bag
                        pGeomColl.AddGeometry(pPolylineErr)

                        'Vérifier si l'élément doit être détruit
                        If pPolygonGen.IsEmpty Then
                            'Détruire l'élément
                            pFeature.Delete()

                            'Si la géométrie doit être modifiée
                        Else
                            'Modifié la géométrie
                            pFeature.Shape = pPolygonGen
                            'Sauver la modification
                            pFeature.Store()
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pNewFeature = Nothing
            pPolygon = Nothing
            pPolygonGen = Nothing
            pPolylineErr = Nothing
            pSquelette = Nothing
            pBagDroites = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pPointsConnexion = Nothing
            pTopoOp = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions pour corriger la généralisation intérieure des éléments de type surface"
    '''<summary>
    ''' Routine qui permet de corriger la généralisation intérieure des éléments sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dSupMin"> Contient la superficie de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="pFeatureLayerGeneraliser"> Contient le FeatureLayer dans lequel les éléments généralisés seront créés.</param>
    '''<param name="pFeatureLayerSquelette"> Contient le FeatureLayer dans lequel les éléments du squelette seront créés.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="bCreerFichierErreurs"> Permet d'indiquer si on doit créer le fichier d'erreurs.</param>
    '''<param name="iNbErreurs"> Contient le nombre d'erreurs.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CorrigerGeneralisationInterieure(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                                ByVal pFeatureLayerGeneraliser As IFeatureLayer, ByVal pFeatureLayerSquelette As IFeatureLayer,
                                                ByVal bCorriger As Boolean, ByVal bCreerFichierErreurs As Boolean, ByRef iNbErreurs As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la Featureclass d'erreur.
        Dim pBagErreurs As IGeometryBag = Nothing           'Interface contenant les lignes et les droites en trop dans la géométrie d'un élément.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Créer le Bag vide des lignes en erreur
                pBagErreurs = New GeometryBag
                pBagErreurs.SpatialReference = m_MxDocument.FocusMap.SpatialReference
                'Interface pour extraire le nombre d'erreurs
                pGeomColl = CType(pBagErreurs, IGeometryCollection)
                iNbErreurs = pGeomColl.GeometryCount

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Correction de la généralisation intérieure (LargMin=" & dLargMin.ToString & ", LongMin=" & dLongMin.ToString & ", SupMin=" & dSupMin.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Définir le polygone à généraliser
                    Call GeneraliserInterieureElement(dPrecision, dDistLat, dLargMin, dLongMin, dSupMin, pFeatureLayerGeneraliser, pFeatureLayerSquelette, bCorriger,
                                                      pTopologyGraph, pFeature, pBagErreurs)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Retourner le nombre d'erreurs
                iNbErreurs = pGeomColl.GeometryCount

                'Vérifier la présecense d'une modification
                If pGeomColl.GeometryCount > 0 Or pFeatureLayerSquelette IsNot Nothing Then
                    'Vérifier si on doit corriger
                    If bCorriger Then
                        'Terminer l'opération UnDo
                        pEditor.StopOperation("Corriger la généralisation intérieure")

                        'Enlever la sélection
                        m_MxDocument.FocusMap.ClearSelection()
                        'Sélectionner les éléments en erreur
                        'm_MxDocument.FocusMap.SelectByShape(pBagErreurs, Nothing, False)

                        'Sinon
                    Else
                        'Annuler l'opération UnDo
                        pEditor.AbortOperation()
                    End If

                    'Vérifier si on doit créer le fichier d'erreurs
                    If bCreerFichierErreurs Then
                        'Initialiser le message d'exécution
                        pTrackCancel.Progressor.Message = "Création du FeatureLayer d'erreurs de généralisation intérieure (NbErr=" & iNbErreurs.ToString & ") ..."
                        'Créer le FeatureLayer des erreurs
                        pFeatureLayer = CreerFeatureLayerErreurs("CorrigerGeneralisationInterieure_1", "Erreur de généralisation intérieure : Larg=" _
                                                                 & dLongMin.ToString & ", LongMin=" & dLongMin.ToString & ", SupMin=" & dSupMin.ToString, _
                                                                 m_MxDocument.FocusMap.SpatialReference, esriGeometryType.esriGeometryPolyline, pBagErreurs)

                        'Ajouter le FeatureLayer d'erreurs dans la map active
                        m_MxDocument.FocusMap.AddLayer(pFeatureLayer)
                    End If

                    'Rafraîchir l'affichage
                    m_MxDocument.ActiveView.Refresh()

                    'Si aucune erreur
                Else
                    'Annuler l'opération UnDo
                    pEditor.AbortOperation()
                End If
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pFeatureLayer = Nothing
            pBagErreurs = Nothing
            pGeomColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de généraliser l'intérieure d'un élément de type surface selon une largeur et une longueur minimum de généralisation et une superficie minimale des surfaces.  
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLargMin"> Contient la largeur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="dSupMin"> Contient la superficie de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="pFeatureLayerGeneraliser"> Contient le FeatureLayer dans lequel les éléments généralisés seront créés.</param>
    '''<param name="pFeatureLayerSquelette"> Contient le FeatureLayer dans lequel les éléments du squelette seront créés.</param>
    '''<param name="bCorriger"> Permet d'indiquer si on doit corriger les lignes en trop des géométries des éléments.</param>
    '''<param name="pFeature"> Interface contenant l'élément à généraliser.</param>
    '''<param name="pBagErreurs"> Contient les géométries d'erreurs.</param>
    ''' 
    Public Sub GeneraliserInterieureElement(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                            ByVal pFeatureLayerGeneraliser As IFeatureLayer, ByVal pFeatureLayerSquelette As IFeatureLayer, ByVal bCorriger As Boolean,
                                            ByVal pTopologyGraph As ITopologyGraph, ByRef pFeature As IFeature, ByRef pBagErreurs As IGeometryBag)
        'Déclarer les variables de travail
        Dim pNewFeature As IFeature = Nothing               'Interface ESRI contenant un élément en sélection.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface contenant le super polygone à traiter.
        Dim pPolygonGen As IPolygon = Nothing               'Interface contenant le polygone de généralisation.
        Dim pPolylineErr As IPolyline = Nothing             'Interface contenant la polyligne d'erreur de généralisation.
        Dim pSquelette As IPolyline = Nothing               'Interface contenant le squelette du polygone.
        Dim pBagDroites As IGeometryBag = Nothing           'Interface contenant les droites de Delaunay.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les lignes en erreurs.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les points de connexion.
        Dim pPointsConnexion As IMultipoint = Nothing       'Interface contenant les points de connexion.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier une géométrie.

        Try
            'Traite tous les éléments sélectionnés
            If pFeature IsNot Nothing Then
                'vérifier si la géométrie de l'élément est de type surface
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Définir le polygone à généraliser
                    pPolygon = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID), IPolygon4)
                    'pPolygon = CType(pFeature.ShapeCopy, IPolygon4)
                    ''Projeter le polygone
                    'pPolygon.Project(pBagErreurs.SpatialReference)

                    'Extraire les points de connexion
                    pPointsConnexion = clsGeneraliserGeometrie.ExtrairePointsIntersection(pFeature, pTopologyGraph)

                    'Généraliser l'intérieur du polygone
                    Call clsGeneraliserGeometrie.GeneraliserInterieurPolygone(pPolygon, pPointsConnexion, dDistLat, dLargMin, dLongMin, dSupMin,
                                                                              pPolygonGen, pPolylineErr, pSquelette, pBagDroites)

                    'Vérifier si une erreur de généralisation est présente
                    If bCorriger And Not pPolylineErr.IsEmpty Then
                        'Interface pour extraire le nombre d'erreurs
                        pGeomColl = CType(pBagErreurs, IGeometryCollection)
                        'Ajouter l'erreur dans le Bag
                        pGeomColl.AddGeometry(pPolylineErr)

                        'Vérifier si l'élément doit être détruit
                        If pPolygonGen.IsEmpty Then
                            'Détruire l'élément
                            pFeature.Delete()

                            'Si la géométrie doit être modifiée
                        Else
                            'Interface pour simplifier une géométrie.
                            pTopoOp = CType(pPolygonGen, ITopologicalOperator2)
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()
                            'Modifié la géométrie
                            pFeature.Shape = pPolygonGen
                            'Sauver la modification
                            pFeature.Store()
                        End If

                        'Vérfier si la classe de gneralisation est spécifié
                        If pFeatureLayerGeneraliser IsNot Nothing Then
                            'Créer l'élément généralisé
                            pNewFeature = pFeatureLayerGeneraliser.FeatureClass.CreateFeature
                            'Définir la géométrie de l'élément généralisé
                            pNewFeature.Shape = pPolylineErr
                            'Sauver l'élément
                            pNewFeature.Store()
                        End If
                    End If

                    'Vérifier si on doit créer le squelette
                    If pFeatureLayerSquelette IsNot Nothing Then
                        'Créer l'élément du squelette
                        pNewFeature = pFeatureLayerSquelette.FeatureClass.CreateFeature
                        'Définir la géométrie de l'élément du squelette
                        pNewFeature.Shape = pSquelette
                        'Sauver l'élément
                        pNewFeature.Store()
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pNewFeature = Nothing
            pPolygon = Nothing
            pPolygonGen = Nothing
            pPolylineErr = Nothing
            pSquelette = Nothing
            pBagDroites = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pPointsConnexion = Nothing
            pTopoOp = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions pour créer les squelettes des éléments de type surface"
    '''<summary>
    ''' Routine qui permet de créer les squelettes des éléments de type surface sélectionnés. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale entre les sommets des géométries des éléments.</param>
    '''<param name="dDistMin"> Contient la distance minimale entre 2 sommets des géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="pFeatureLayerSquelette"> Contient le FeatureLayer dans lequel les éléments du squelette seront créés.</param>
    '''<param name="pTrackCancel"> Contient la barre de progression.</param>
    ''' 
    Public Sub CreerSquelettes(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dDistMin As Double, ByVal dLongMin As Double,
                               ByVal pFeatureLayerSquelette As IFeatureLayer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pEditor As IEditor = Nothing                    'Interface ESRI pour effectuer l'édition des éléments.
        Dim pGererMapLayer As clsGererMapLayer = Nothing    'Objet utiliser pour extraire la collection des FeatureLayers visibles.
        Dim pEnumFeature As IEnumFeature = Nothing          'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing                  'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing                'Interface contenant l'enveloppe de l'élément traité.
        Dim pTopologyGraph As ITopologyGraph = Nothing      'Interface contenant la topologie.
        Dim pFeatureLayersColl As Collection = Nothing      'Objet contenant la collection des FeatureLayers utilisés dans la Topologie.
        Dim pMap As IMap = Nothing                          'Interface pour sélectionner des éléments.

        Try
            'Interface pour vérifer si on est en mode édition
            pEditor = CType(m_Application.FindExtensionByName("ESRI Object Editor"), IEditor)

            'Vérifier si on est en mode édition
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'Débuter l'opération UnDo
                pEditor.StartOperation()

                'Objet utiliser pour extraire la collection des FeatureLayers
                pGererMapLayer = New clsGererMapLayer(m_MxDocument.FocusMap)
                'Définir la collection des FeatureLayers utilisés dans la topologie
                pFeatureLayersColl = pGererMapLayer.DefinirCollectionFeatureLayer(False)

                'Définir l'enveloppe des éléments sélectionnés
                'pEnvelope = EnveloppeElementsSelectionner(pEditor)
                pEnvelope = m_MxDocument.ActiveView.Extent

                'Création de la topologie
                pTrackCancel.Progressor.Message = "Création de la topologie : Précision=" & dPrecision.ToString & "..."
                pTopologyGraph = CreerTopologyGraph(pEnvelope, pFeatureLayersColl, dPrecision)

                'Initialiser le message d'exécution
                pTrackCancel.Progressor.Message = "Création des squelettes (DistMin=" & dDistMin.ToString & ", LongMin=" & dLongMin.ToString & ") ..."

                'Interface pour extraire le premier élément de la sélection
                pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)
                'Réinitialiser la recherche des éléments
                pEnumFeature.Reset()
                'Extraire le premier élément de la sélection
                pFeature = pEnumFeature.Next

                'Traite tous les éléments sélectionnés
                Do Until pFeature Is Nothing
                    'Créer le squelette de l'élément
                    Call CreerSqueletteElement(dPrecision, dDistLat, dDistMin, dLongMin, pFeatureLayerSquelette, pTopologyGraph, pFeature)

                    'Vérifier si un Cancel a été effectué
                    If pTrackCancel.Continue = False Then Exit Do

                    'Extraire le prochain élément de la sélection
                    pFeature = pEnumFeature.Next
                Loop

                'Terminer l'opération UnDo
                pEditor.StopOperation("Créer les squelettes")

                'Enlever la sélection
                m_MxDocument.FocusMap.ClearSelection()

                'Sélectionner les éléments en erreur
                'm_MxDocument.FocusMap.SelectByShape(pBagErreurs, Nothing, False)

                'Rafraîchir l'affichage
                m_MxDocument.ActiveView.Refresh()

                'Sinon
            Else
                'Annuler l'opération UnDo
                pEditor.AbortOperation()
            End If

            'Désactiver l'interface d'édition
            pEditor = Nothing

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Annuler l'opération UnDo
            If Not pEditor Is Nothing Then pEditor.AbortOperation()
            'Vider la mémoire
            pEditor = Nothing
            pGererMapLayer = Nothing
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
            pTopologyGraph = Nothing
            pFeatureLayersColl = Nothing
            pMap = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de créer le squelette d'un élément de type surface. 
    '''</summary>
    ''' 
    '''<param name="dPrecision"> Contient la précision des données utilisée pour la topologie des éléments.</param>
    '''<param name="dDistLat"> Contient la distance latérale minimale entre les sommets des géométries des éléments.</param>
    '''<param name="dDistMin"> Contient la distance minimale entre 2 sommets des géométries des éléments.</param>
    '''<param name="dLongMin"> Contient la longueur de généralisation minimale utilisée pour généraliser les géométries des éléments.</param>
    '''<param name="pFeatureLayerSquelette"> Contient le FeatureLayer dans lequel les éléments du squelette seront créés.</param>
    '''<param name="pFeature"> Interface contenant l'élément à généraliser.</param>
    ''' 
    Public Sub CreerSqueletteElement(ByVal dPrecision As Double, ByVal dDistLat As Double, ByVal dDistMin As Double, ByVal dLongMin As Double,
                                     ByVal pFeatureLayerSquelette As IFeatureLayer, ByVal pTopologyGraph As ITopologyGraph, ByRef pFeature As IFeature)
        'Déclarer les variables de travail
        Dim pNewFeature As IFeature = Nothing               'Interface ESRI contenant un élément en sélection.
        Dim pPolygon As IPolygon4 = Nothing                 'Interface contenant le super polygone à traiter.
        Dim pSquelette As IPolyline = Nothing               'Interface contenant le squelette du polygone.
        Dim pBagDroites As IGeometryBag = Nothing           'Interface contenant les droites de Delaunay.
        Dim pPointsConnexion As IMultipoint = Nothing       'Interface contenant les points de connexion.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour simplifier une géométrie.

        Try
            'Traite tous les éléments sélectionnés
            If pFeature IsNot Nothing Then
                'Vérifier si la géométrie de l'élément est de type surface
                If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'Définir le polygone à généraliser
                    pPolygon = CType(pTopologyGraph.GetParentGeometry(CType(pFeature.Class, IFeatureClass), pFeature.OID), IPolygon4)

                    'Extraire les points de connexion
                    pPointsConnexion = clsGeneraliserGeometrie.ExtrairePointsIntersection(pFeature, pTopologyGraph)

                    'Créer le squelette de l'élément
                    Call clsGeneraliserGeometrie.CreerSquelettePolygoneDelaunay(pPolygon, pPointsConnexion, dDistLat, dDistMin, dLongMin, pSquelette, pBagDroites)

                    'Vérifier si on doit créer le squelette
                    If pFeatureLayerSquelette IsNot Nothing Then
                        'Créer l'élément du squelette
                        pNewFeature = pFeatureLayerSquelette.FeatureClass.CreateFeature
                        'Définir la géométrie de l'élément du squelette
                        pNewFeature.Shape = pSquelette
                        'Sauver l'élément
                        pNewFeature.Store()
                    End If
                End If
            End If

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pNewFeature = Nothing
            pPolygon = Nothing
            pSquelette = Nothing
            pBagDroites = Nothing
            pPointsConnexion = Nothing
            pTopoOp = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions utilitaires"
    '''<summary>
    ''' Fonction qui permet d'extraire l'enveloppe des éléments sélectionnés et qui peuvent être modifiés. 
    '''</summary>
    ''' 
    '''<param name="pEditor"> Interface contenant les éléments sélectionnés qui peuvent être modifiés.</param>
    ''' 
    '''<returns>IEnvelope contenant l'étendue des éléments sélectionnées.</returns>
    ''' 
    Public Function EnveloppeElementsSelectionner(ByVal pEditor As IEditor) As IEnvelope
        'Déclarer les variables de travail
        Dim pEnumFeature As IEnumFeature = Nothing      'Interface ESRI utilisé pour extraire les éléments de la sélection.
        Dim pFeature As IFeature = Nothing              'Interface ESRI contenant un élément en sélection.
        Dim pEnvelope As IEnvelope = Nothing            'Interface contenant l'enveloppe de l'élément traité.

        'Par défaut l'enveloppe est invalide
        EnveloppeElementsSelectionner = Nothing

        Try
            'Interface pour extraire le premier élément de la sélection
            pEnumFeature = CType(pEditor.EditSelection, IEnumFeature)

            'Extraire le premier élément de la sélection
            pFeature = pEnumFeature.Next

            'Traite tous les éléments sélectionnés
            Do Until pFeature Is Nothing
                'Vérifier si l'envelope n'est pas initialisé
                If pEnvelope Is Nothing Then
                    'Initialiser le premier enveloppe
                    pEnvelope = pFeature.Shape.Envelope
                Else
                    'Union des enveloppes des autres éléments
                    pEnvelope.Union(pFeature.Shape.Envelope)
                End If

                'Extraire le prochain élément de la sélection
                pFeature = pEnumFeature.Next
            Loop

            'Agrandir l'enveloppe de tous les éléments sélectionnés
            pEnvelope.Expand(1.5, 1.5, True)

            'Retourner l'enveloppe
            EnveloppeElementsSelectionner = pEnvelope

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEnumFeature = Nothing
            pFeature = Nothing
            pEnvelope = Nothing
        End Try
    End Function

    '''<summary>
    '''Fonction qui permet de créer et retourner la topologie en mémoire des éléments entre une collection de FeatureLayer.
    '''</summary>
    '''
    '''<param name="pEnvelope">Interface ESRI contenant l'enveloppe de création de la topologie traitée.</param>
    '''<param name="qFeatureLayerColl">Interface ESRI contenant les FeatureLayers pour traiter la topologie.</param> 
    '''<param name="dTolerance">Tolerance de proximité.</param> 
    '''
    '''<returns>"ITopologyGraph" contenant la topologie des classes de données, "Nothing" sinon.</returns>
    '''
    Public Function CreerTopologyGraph(ByVal pEnvelope As IEnvelope, ByVal qFeatureLayerColl As Collection, ByVal dTolerance As Double) As ITopologyGraph4
        'Déclarer les variables de travail
        Dim qType As Type = Nothing                         'Interface contenant le type d'objet à créer.
        Dim oObjet As System.Object = Nothing               'Interface contenant l'objet correspondant à l'application.
        Dim pTopologyExt As ITopologyExtension = Nothing    'Interface contenant l'extension de la topologie.
        Dim pMapTopology As IMapTopology2 = Nothing         'Interface utilisé pour créer la topologie.
        Dim pTopologyGraph As ITopologyGraph4 = Nothing     'Interface contenant la topologie.
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant la classe de données.

        'Définir la valeur de retour par défaut
        CreerTopologyGraph = Nothing

        Try
            'Définir l'extension de topologie
            qType = Type.GetTypeFromProgID("esriEditorExt.TopologyExtension")
            oObjet = Activator.CreateInstance(qType)
            pTopologyExt = CType(oObjet, ITopologyExtension)

            'Définir l'interface pour créer la topologie
            pMapTopology = CType(pTopologyExt.MapTopology, IMapTopology2)

            'S'assurer que laliste des Layers est vide
            pMapTopology.ClearLayers()

            'Traiter tous les FeatureLayer présents
            For Each pFeatureLayer In qFeatureLayerColl
                'Ajouter le FeatureLayer à la topologie
                pMapTopology.AddLayer(pFeatureLayer)
            Next

            'Changer la référence spatiale selon l'enveloppe
            pMapTopology.SpatialReference = pEnvelope.SpatialReference

            'Définir la tolérance de connexion et de partage
            pMapTopology.ClusterTolerance = dTolerance

            'Interface pour construre la topologie
            pTopologyGraph = CType(pMapTopology.Cache, ITopologyGraph4)
            pTopologyGraph.SetEmpty()

            Try
                'Construire la topologie
                pTopologyGraph.Build(pEnvelope, False)
            Catch ex As OutOfMemoryException
                'Retourner une erreur de création de la topologie
                Throw New Exception("Incapable de créer la topologie : OutOfMemoryException")
            Catch ex As Exception
                'Retourner une erreur de création de la topologie
                Throw New Exception("Incapable de créer la topologie : " & ex.Message)
            End Try

            'Retourner la topologie
            CreerTopologyGraph = pTopologyGraph

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            qType = Nothing
            oObjet = Nothing
            pTopologyExt = Nothing
            pTopologyGraph = Nothing
            pMapTopology = Nothing
            pFeatureLayer = Nothing
            'Récupération de la mémoire disponible
            GC.Collect()
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de créer un FeatureLayer contenant une nouovelle classe d'erreurs en mémoire. 
    '''</summary>
    ''' 
    '''<param name="sNomLayer"> Contient le nom du FeatureLayer d'erreurs à créer.</param>
    '''<param name="sDescription"> Contient la description des erreurs du FeatureLayer d'erreurs à créer.</param>
    '''<param name="pSpatialReference"> Interface contenant la référence spatiale du FeatureLayer d'erreurs à créer.</param>
    '''<param name="pEsriGeometryType"> Interface contenant le type de géométrie du FeatureLayer d'erreurs à créer.</param>
    '''<param name="pGeometryBag"> Interface contenant les géométries du FeatureLayer d'erreurs à créer.</param>
    ''' 
    Public Function CreerFeatureLayerErreurs(ByVal sNomLayer As String, ByVal sDescription As String, ByVal pSpatialReference As ISpatialReference, _
                                             ByVal pEsriGeometryType As esriGeometryType, ByVal pGeometryBag As IGeometryBag) As IFeatureLayer
        'Déclarer les variables de travail
        Dim pFeatureClassErreur As IFeatureClass = Nothing      'Interfaqce contenant les éléments en erreur.
        Dim pFeatureCursorErreur As IFeatureCursor = Nothing    'Interface contenant les erreurs à écrire.
        Dim pFeatureBuffer As IFeatureBuffer = Nothing          'Interface contenant l'élément en erreur.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les lignes en erreurs.
        Dim pArea As IArea = Nothing                            'Interface utilisé pour extraire la superficie d'une surface.
        Dim pPolyline As IPolyline = Nothing                    'Interface utilisé pour extraire la longueur d'une ligne.
        Dim pPointColl As IPointCollection = Nothing            'Interface utilisé pour extraire le nombre de sommets.
        Dim pGeometry As IGeometry = Nothing                    'Interface contenant la géométrie en erreur.
        Dim pMultipoint As IMultipoint = Nothing                'Interface contenant les points de la géométrie en erreur.
        Dim pZAware As IZAware = Nothing                        'Interface pour désactiver le Z.
        Dim pMAware As IMAware = Nothing                        'Interface pour désactiver le M.

        'Créer un nouveau FeatureLayer
        CreerFeatureLayerErreurs = New FeatureLayer

        Try
            'Créer une classe d'erreurs en mémoire
            pFeatureClassErreur = CreateInMemoryFeatureClass(sNomLayer, pSpatialReference, pEsriGeometryType)

            'Interface pour créer les erreurs
            pFeatureCursorErreur = pFeatureClassErreur.Insert(True)

            'Interface pour extraire les géométries en erreurs
            pGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter toutes les erreurs
            For i = 0 To pGeomColl.GeometryCount - 1
                'Créer un FeatureBuffer Point
                pFeatureBuffer = pFeatureClassErreur.CreateFeatureBuffer

                'Définir la géométrie
                pGeometry = pGeomColl.Geometry(i)
                'Traiter le Z
                pZAware = CType(pGeometry, IZAware)
                pZAware.ZAware = False
                'Traiter le M
                pMAware = CType(pGeometry, IMAware)
                pMAware.MAware = False

                'Vérifier si le type de géométrie est valide 
                If pGeometry.GeometryType = pEsriGeometryType Then
                    'Définir la géométrie
                    pFeatureBuffer.Shape = pGeometry

                    'Si le type est un multipoint
                ElseIf pEsriGeometryType = esriGeometryType.esriGeometryMultipoint Then
                    'Définir un multipoint vide
                    pMultipoint = New Multipoint
                    pMultipoint.SpatialReference = pGeometry.SpatialReference

                    'Interface pout ajouter les points
                    pPointColl = CType(pMultipoint, IPointCollection)
                    'Ajouter les points
                    pPointColl.AddPointCollection(CType(pGeometry, IPointCollection))

                    'Définir la géométrie
                    pGeometry = pMultipoint
                    pFeatureBuffer.Shape = pGeometry
                End If

                'Définir la description
                pFeatureBuffer.Value(1) = sDescription

                'Vérifier si le type de geométrie est un multipoint
                If pGeomColl.Geometry(i).GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                    'Définir la valeur obtenue
                    pFeatureBuffer.Value(2) = i

                    'Si le type de geométrie est une multipoint
                ElseIf pGeomColl.Geometry(i).GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryMultipoint Then
                    'Interface utilisé pour extraire le nombre de sommets.
                    pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)
                    'Définir la valeur obtenue
                    pFeatureBuffer.Value(2) = pPointColl.PointCount

                    'Si le type de geométrie est une polyline
                ElseIf pGeomColl.Geometry(i).GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
                    'Interface utilisé pour extraire le nombre de sommets.
                    pPolyline = CType(pGeomColl.Geometry(i), IPolyline)
                    'Définir la valeur obtenue
                    pFeatureBuffer.Value(2) = pPolyline.Length

                    'Si le type de geométrie est un polygon
                ElseIf pGeomColl.Geometry(i).GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then
                    'Interface utilisé pour extraire le nombre de sommets.
                    pArea = CType(pGeomColl.Geometry(i), IArea)
                    'Définir la valeur obtenue
                    pFeatureBuffer.Value(2) = pArea.Area
                End If

                'Insérer un nouvel élément dans la FeatureClass d'erreur
                pFeatureCursorErreur.InsertFeature(pFeatureBuffer)
            Next

            'Définir la FeatureClass du FeatureLayer
            CreerFeatureLayerErreurs.FeatureClass = pFeatureClassErreur
            'Définir le nom du FeatureLayer selon le nom de la FeatureClass
            CreerFeatureLayerErreurs.Name = pFeatureClassErreur.AliasName
            'Rendre visible le FeatureLayer
            CreerFeatureLayerErreurs.Visible = True

        Catch ex As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureClassErreur = Nothing
            pFeatureCursorErreur = Nothing
            pFeatureBuffer = Nothing
            pGeomColl = Nothing
            pArea = Nothing
            pPolyline = Nothing
            pPointColl = Nothing
            pGeometry = Nothing
            pMultipoint = Nothing
            pZAware = Nothing
            pMAware = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de créer une nouvelle FeatureClass en mémoire.
    '''</summary>
    '''
    '''<param name="sNom">Nom de la classe à créer.</param>
    '''<param name="pSpatialReference">Interface contenant la référence spatiale de la FeatureClass des erreurs.</param>
    '''<param name="pEsriGeometryType">Indique le type de géométrie de la FeatureClass.</param>
    ''' 
    '''<returns>"IFeatureClass" contenant la description et la géométrie trouvées.</returns>
    '''
    Public Function CreateInMemoryFeatureClass(ByVal sNom As String, ByVal pSpatialReference As ISpatialReference, ByVal pEsriGeometryType As esriGeometryType) As IFeatureClass
        'Déclarer les variables de travail
        Dim pWorkspaceFactory As IWorkspaceFactory = Nothing    'Interface pour créer un Workspace en mémoire.
        Dim pName As IName = Nothing                            'Interface pour ouvrir un workspace.
        Dim pFeatureWorkspace As IFeatureWorkspace = Nothing    'Interface contenant un FeatureWorkspace.
        Dim pFields As IFields = Nothing                        'Interface pour contenant les attributs de la Featureclass.
        Dim pFieldsEdit As IFieldsEdit = Nothing                'Interface pour créer les attributs.
        Dim pFieldEdit As IFieldEdit = Nothing                  'Interface pour créer un attribut.
        Dim pClone As IClone = Nothing                          'Interface utilisé pour cloner un attribut.
        Dim pGeometryDef As IGeometryDefEdit = Nothing          'Interface ESRI utilisé pour créer la structure d'un géométrie.
        Dim qFactoryType As Type = Nothing                      'Interface contenant le type d'objet à créer.
        Dim pUID As New UID                                     'Interface pour générer un UID.

        'Définir la valeur par défaut
        CreateInMemoryFeatureClass = Nothing

        Try
            'Définir le type de Factory
            qFactoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.InMemoryWorkspaceFactory")
            'Générer l'interface pour créer un Workspace en mémoire
            pWorkspaceFactory = CType(Activator.CreateInstance(qFactoryType), IWorkspaceFactory)

            'Créer un nouveau UID
            pUID.Generate()

            'Creer un workspace en mémoire
            pName = CType(pWorkspaceFactory.Create("", pUID.Value.ToString, Nothing, 0), IName)

            'Définir le FeatureWorkspace pour créer une Featureclass
            pFeatureWorkspace = CType(pName.Open, IFeatureWorkspace)

            'Définir le type d'élément
            pUID.Value = "esriGeodatabase.Feature"

            'Interface pour créer des attributs
            pFieldsEdit = New Fields

            'Définir le nombre d'attributs
            pFieldsEdit.FieldCount_2 = 4

            'Créer l'attribut du OBJECTID
            pFieldEdit = New Field
            With pFieldEdit
                .Name_2 = "OBJECTID"
                .AliasName_2 = "OBJECTID"
                .Type_2 = esriFieldType.esriFieldTypeOID
            End With
            'Ajouter l'attribut
            pFieldsEdit.Field_2(0) = pFieldEdit

            'Créer l'attribut de description
            pFieldEdit = New Field
            With pFieldEdit
                .Name_2 = "DESCRIPTION"
                .AliasName_2 = "DESCRIPTION"
                .Type_2 = esriFieldType.esriFieldTypeString
                .Length_2 = 256
                .IsNullable_2 = True
            End With
            'Ajouter l'attribut
            pFieldsEdit.Field_2(1) = pFieldEdit

            'Créer l'attribut de la valeur obtenue
            pFieldEdit = New Field
            With pFieldEdit
                .Name_2 = "VALEUR"
                .AliasName_2 = "VALEUR"
                .Type_2 = esriFieldType.esriFieldTypeSingle
                .IsNullable_2 = True
            End With
            'Ajouter l'attribut
            pFieldsEdit.Field_2(2) = pFieldEdit

            'Cloner l'attribut
            pFieldEdit = New Field
            'Vérifier si le Type de Géométrie n'est pas spécifié
            If pEsriGeometryType <> esriGeometryType.esriGeometryNull Then
                'Interface pour définir le type de géométrie
                pGeometryDef = New GeometryDefClass()

                'Définir la référence spatiale
                pGeometryDef.SpatialReference_2 = pSpatialReference
                'Définir le type de géométrie
                pGeometryDef.GeometryType_2 = pEsriGeometryType
                'Enlever le Z
                pGeometryDef.HasZ_2 = False
                'Enlever le M
                pGeometryDef.HasM_2 = False
                ' Set the grid count to 1 and the grid size to 0 to allow ArcGIS to
                ' determine a valid grid size.
                pGeometryDef.GridCount_2 = 1
                pGeometryDef.GridSize_2(0) = 0

                'définir le nom de l'attribut
                pFieldEdit.Name_2 = "SHAPE"
                'Définir le type d'attribut
                pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry
                'Ajouter la définition de la géométrie dans l'attribut
                pFieldEdit.GeometryDef_2 = pGeometryDef
            End If
            'Ajouter l'attribut
            pFieldsEdit.Field_2(3) = pFieldEdit

            'Créer la Featureclass
            CreateInMemoryFeatureClass = pFeatureWorkspace.CreateFeatureClass(sNom, pFieldsEdit, pUID, Nothing, esriFeatureType.esriFTSimple, "SHAPE", "")

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pWorkspaceFactory = Nothing
            pName = Nothing
            pFeatureWorkspace = Nothing
            pFields = Nothing
            pFieldsEdit = Nothing
            pFieldEdit = Nothing
            pClone = Nothing
            pGeometryDef = Nothing
            qFactoryType = Nothing
            pUID = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet de traiter le Z d'une Géométrie.
    '''</summary>
    '''
    '''<param name=" pGeometry "> Interface contenant la géométrie à traiter.</param>
    '''<param name=" dZ "> Contient la valeur du Z.</param>
    '''
    Public Sub TraiterZ(ByRef pGeometry As IGeometry, Optional ByVal dZ As Double = 0)
        'Déclarer les variables de travail
        Dim pZAware As IZAware = Nothing                'Interface ESRI utilisée pour traiter le Z.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface qui permet d'accéder aux géométries
        Dim pPointColl As IPointCollection = Nothing    'Interface qui permet d'accéder aux sommets de la géométrie
        Dim pPoint As IPoint = Nothing                  'Interface qui permet de modifier le Z
        Dim pZ As IZ = Nothing                          'Interface utilisé pour calculer le Z invalide
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Écrire une trace de début
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Debut")

            'Interface pour traiter le Z
            pZAware = CType(pGeometry, IZAware)

            'Vérifier la présence du Z
            If pZAware.ZAware Then
                'Vérifier si on doit corriger le Z
                If pZAware.ZSimple = False Then
                    'Vérifier si la géométrie est un Point
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Définir le point
                        pPoint = CType(pGeometry, IPoint)
                        'Interface pour traiter le Z
                        pZAware = CType(pPoint, IZAware)
                        'Vérifier si on doit corriger le Z
                        If pZAware.ZSimple = False Then
                            'Définir le Z du point
                            pPoint.Z = dZ
                        End If

                        'Vérifier si la géométrie est un Bag
                    ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                        'Interface utilisé pour accéder aux géométries
                        pGeomColl = CType(pGeometry, IGeometryCollection)
                        'Traiter toutes les géométries
                        For i = 0 To pGeomColl.GeometryCount - 1
                            'Traiter le Z
                            Call TraiterZ(pGeomColl.Geometry(i), dZ)
                        Next

                        'Vérifier si la géométrie est un Multipoint,une Polyline ou un Polygon
                    Else
                        Try
                            'Interface pour corriger le Z par interpollation
                            pZ = CType(pGeometry, IZ)
                            'Corriger le Z invalide par interpollation
                            pZ.CalculateNonSimpleZs()
                        Catch ex As Exception
                            'On ne fait rien
                        End Try

                        'Vérifier si on doit corriger le Z
                        If pZAware.ZSimple = False Then
                            'Interface utilisé pour accéder aux sommets de la géométrie
                            pPointColl = CType(pGeometry, IPointCollection)
                            'Traiter tous les sommets de la géométrie
                            For i = 0 To pPointColl.PointCount - 1
                                'Interface pour définir le Z
                                pPoint = pPointColl.Point(i)
                                'Interface pour traiter le Z
                                pZAware = CType(pPoint, IZAware)
                                'Vérifier si on doit corriger le Z
                                If pZAware.ZSimple = False Then
                                    'Définir le Z du point
                                    pPoint.Z = dZ
                                End If
                                'Conserver les modifications
                                pPointColl.UpdatePoint(i, pPoint)
                            Next
                        End If
                    End If
                End If

                'Si aucun Z
            Else
                'Enlever le 3D
                pZAware.ZAware = True
                pZAware.DropZs()
                pZAware.ZAware = False
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pZAware = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pPoint = Nothing
            pZ = Nothing
            'Écrire une trace de fin
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Fin")
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de traiter le M d'une Géométrie.
    '''</summary>
    '''
    '''<param name=" pGeometry "> Interface contenant la géométrie à traiter.</param>
    '''<param name=" dM "> Contient la valeur du M.</param>
    '''
    Public Sub TraiterM(ByRef pGeometry As IGeometry, Optional ByVal dM As Double = 0)
        'Déclarer les variables de travail
        Dim pMAware As IMAware = Nothing                'Interface ESRI utilisée pour traiter le M.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface qui permet d'accéder aux géométries
        Dim pPointColl As IPointCollection = Nothing    'Interface qui permet d'accéder aux sommets de la géométrie
        Dim pPoint As IPoint = Nothing                  'Interface qui permet de modifier le M
        Dim i As Integer = Nothing                      'Compteur

        Try
            'Écrire une trace de début
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Debut")

            'Interface pour traiter le M
            pMAware = CType(pGeometry, IMAware)

            'Vérifier la présence du M
            If pMAware.MAware Then
                'Corriger le M au besoin
                If pMAware.MSimple = False Then
                    'Vérifier si la géométrie est un Point
                    If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                        'Définir le point
                        pPoint = CType(pGeometry, IPoint)
                        'Interface pour traiter le M
                        pMAware = CType(pPoint, IMAware)
                        'Vérifier si on doit corriger le M
                        If pMAware.MSimple = False Then
                            'Définir le Z du point
                            pPoint.M = dM
                        End If

                        'Vérifier si la géométrie est un Bag
                    ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                        'Interface utilisé pour accéder aux géométries
                        pGeomColl = CType(pGeometry, IGeometryCollection)
                        'Traiter toutes les géométries
                        For i = 0 To pGeomColl.GeometryCount - 1
                            'Traiter le M
                            Call TraiterM(pGeomColl.Geometry(i), dM)
                        Next

                        'Vérifier si la géométrie est un Multipoint,une Polyline ou un Polygon
                    Else
                        'Interface utilisé pour accéder aux sommets de la géométrie
                        pPointColl = CType(pGeometry, IPointCollection)
                        'Traiter tous les sommets de la géométrie
                        For i = 0 To pPointColl.PointCount - 1
                            'Interface pour définir le M
                            pPoint = pPointColl.Point(i)
                            'Interface pour traiter le M
                            pMAware = CType(pPoint, IMAware)
                            'Vérifier si on doit corriger le M
                            If pMAware.MSimple = False Then
                                'Définir le Z du point
                                pPoint.M = dM
                            End If
                            'Conserver les modifications
                            pPointColl.UpdatePoint(i, pPoint)
                        Next
                    End If
                End If

                'Si aucun M
            Else
                'Enlever le M
                pMAware.MAware = True
                pMAware.DropMs()
                pMAware.MAware = False
            End If

        Catch erreur As Exception
            MessageBox.Show(erreur.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Finally
            'Vider la mémoire
            pMAware = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pPoint = Nothing
            'Écrire une trace de fin
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Fin")
        End Try
    End Sub

    '''<summary>
    ''' Fonction qui permet de retourner le buffer de la géométrie spécifiée en paramètre selon une largeur.
    ''' 
    ''' Si la géométrie est vide, le buffer sera vide.
    ''' Si la largeur est nulle, la largeur changée pour la valeur 0.001.
    '''</summary>
    '''
    '''<param name="pGeometry">Interface ESRI de la Géométrie à traiter.</param>
    '''<param name="dLargeur">Largeur utilisée pour créer le buffer de la géométrie.</param>
    ''' 
    '''<returns>La fonction va retourner un "IPolygon". Sinon "Nothing".</returns>
    '''
    Public Function fpBufferGeometrie(ByVal pGeometry As IGeometry, ByVal dLargeur As Double) As IPolygon
        'Déclarer les variables de travail
        Dim pPolygon As IPolygon = Nothing              'Interface contenant le buffer de la géométrie
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface utilisé pour créer le buffer
        Dim pMultipoint As IMultipoint = Nothing        'Interface contenant le point
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface utilisé pour ajouter le point dans le Multipoint
        Dim i As Integer = Nothing                      'Compteur

        'Définir la valeur par défaut
        fpBufferGeometrie = Nothing

        Try
            'Écrire une trace de début
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Debut")

            'Vérifier si la largeur est nulle
            If dLargeur <= 0 Then dLargeur = 0.001

            'Vérifier si la géométrie est vide ou qu'il n'y a pas de largeur
            If pGeometry.IsEmpty Then
                'Créer un polygon vide
                pPolygon = CType(New Polygon, IPolygon)
                pPolygon.SpatialReference = pGeometry.SpatialReference

                'Vérifier si la géométrie est un Bag
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryBag Then
                'Créer un polygon vide
                pPolygon = CType(New Polygon, IPolygon)
                pPolygon.SpatialReference = pGeometry.SpatialReference
                'Traiter toutes les composantes de géométries
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Créer le buffer
                    pGeomColl = CType(fpBufferGeometrie(pGeomColl.Geometry(i), dLargeur), IGeometryCollection)
                    'Ajouter le polygon précédent
                    pGeomColl.AddGeometry(pPolygon)
                Next
                'Définir le buffer
                pPolygon = CType(pGeomColl, IPolygon)

                'Vérifier si la géométrie est un point
            ElseIf pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Créer un multipoint vide
                pMultipoint = CType(New Multipoint, IMultipoint)
                pMultipoint.SpatialReference = pGeometry.SpatialReference
                'Ajouter le point dans le multipoint
                pGeomColl = CType(pMultipoint, IGeometryCollection)
                pGeomColl.AddGeometry(pGeometry)
                'Interface pour créer le buffer
                pTopoOp = CType(pMultipoint, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
                'Créer le buffer
                pPolygon = CType(pTopoOp.Buffer(dLargeur), IPolygon)

                'Pour les autres types
            Else
                'Simplifier la géométrie
                pTopoOp = CType(pGeometry, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()
                'Créer le buffer
                pPolygon = CType(pTopoOp.Buffer(dLargeur), IPolygon)
            End If

            'Retourner la même géométrie
            fpBufferGeometrie = pPolygon

        Catch e As Exception
            'Message d'erreur
            Err.Raise(vbObjectError + 1, "", e.ToString)
        Finally
            'Écrire une trace de fin
            'Call cits.CartoNord.UtilitairesNet.EcrireMessageTrace("Fin")
            'Vider la mémoire
            pPolygon = Nothing
            pTopoOp = Nothing
            pMultipoint = Nothing
            pGeomColl = Nothing
            i = Nothing
        End Try
    End Function

    '''<summary>
    ''' Routine qui permet d'initialiser la barre de progression.
    ''' 
    '''<param name="iMin">Valeur minimum.</param>
    '''<param name="iMax">Valeur maximum.</param>
    '''<param name="pTrackCancel">Interface contenant la barre de progression.</param>
    ''' 
    '''</summary>
    '''
    Private Sub InitBarreProgression(ByVal iMin As Integer, ByVal iMax As Integer, ByRef pTrackCancel As ITrackCancel)
        'Déclarer les variables de travail
        Dim pStepPro As IStepProgressor = Nothing   'Interface qui permet de modifier les paramètres de la barre de progression.

        Try
            'sortir si le progressor est absent
            If pTrackCancel.Progressor Is Nothing Then Exit Sub

            'Interface pour modifier les paramètres de la barre de progression.
            pTrackCancel.Progressor = m_Application.StatusBar.ProgressBar
            pStepPro = CType(pTrackCancel.Progressor, IStepProgressor)

            'Changer les paramètres
            pStepPro.MinRange = iMin
            pStepPro.MaxRange = iMax
            pStepPro.Position = 0
            pStepPro.Show()

        Catch ex As Exception
            'Retourner l'erreur
            Throw ex
        Finally
            'Vider la mémoire
            pStepPro = Nothing
        End Try
    End Sub
#End Region
End Module
