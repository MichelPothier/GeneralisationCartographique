Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase

'''<summary>
''' Classe qui permet de manipuler les différents types de Layer contenu dans une Map.
'''</summary>
'''
'''<remarks>
''' Auteur : Michel Pothier
''' Date : 6 Mai 2011
'''</remarks>
''' 
Public Class clsGererMapLayer
    'Déclarer les variables globales
    '''<summary>Interface contenant la Map à Gérer.</summary>
    Protected gpMap As IMap = Nothing
    '''<summary>Collection des FeatureLayers de la Map.</summary>
    Protected gqFeatureLayerCollection As Collection = Nothing
    '''<summary>Collection des RasterLayers de la Map.</summary>
    Protected gqRasterLayerCollection As Collection = Nothing
    '''<summary>Collection des FeatureLayers de la Map.</summary>
    Protected gpEnvelope As IEnvelope = Nothing

#Region "Constructeur"
    '''<summary>
    ''' Routine qui permet d'initialiser la classe.
    ''' 
    '''<param name="pMap"> Interface ESRI utilisée pour extraire les FeatureLayers visible ou non.</param>
    '''<param name="bNonVisible"> Indique si on doit aussi extraire les FeatureLayers non visibles.</param>
    ''' 
    '''</summary>
    '''
    Public Sub New(ByVal pMap As IMap, Optional ByVal bNonVisible As Boolean = False)
        'Définir les valeur par défaut
        gpMap = pMap

        'Définir la collection des FeatureLayer visibles par défaut
        gqFeatureLayerCollection = DefinirCollectionFeatureLayer(bNonVisible)

        'Définir la collection des RasterLayer visibles par défaut
        gqRasterLayerCollection = DefinirCollectionRasterLayer(bNonVisible)
    End Sub

    '''<summary>
    '''Routine qui permet de vider la mémoire des objets de la classe.
    '''</summary>
    '''
    Protected Overrides Sub finalize()
        'Vider la mémoire
        gpMap = Nothing
        gqFeatureLayerCollection = Nothing
        gqRasterLayerCollection = Nothing
        gpEnvelope = Nothing
        'Récupération de la mémoire disponible
        GC.Collect()
    End Sub
#End Region

#Region "Propriétés"
    '''<summary>
    ''' Propriété qui permet de définir et retourner la Map traitée.
    '''</summary>
    ''' 
    Public Property Map() As IMap
        Get
            Map = gpMap
        End Get
        Set(ByVal value As IMap)
            gpMap = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la collection de FeatureLayer traitée.
    '''</summary>
    ''' 
    Public Property FeatureLayerCollection() As Collection
        Get
            FeatureLayerCollection = gqFeatureLayerCollection
        End Get
        Set(ByVal value As Collection)
            gqFeatureLayerCollection = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner la collection de RasterLayer traitée.
    '''</summary>
    ''' 
    Public Property RasterLayerCollection() As Collection
        Get
            RasterLayerCollection = gqRasterLayerCollection
        End Get
        Set(ByVal value As Collection)
            gqRasterLayerCollection = value
        End Set
    End Property

    '''<summary>
    ''' Propriété qui permet de définir et retourner l'enveloppe de la collection de FeatureLayer traitée.
    '''</summary>
    ''' 
    Public Property EnvelopeFeatureLayerCollection() As IEnvelope
        Get
            EnvelopeFeatureLayerCollection = gpEnvelope
        End Get
        Set(ByVal value As IEnvelope)
            gpEnvelope = value
        End Set
    End Property
#End Region

#Region "Routine et fonction publiques"

    '''<summary>
    ''' Fonction qui permet d'extraire le FeatureLayer correspondant au nom spécifié.
    '''</summary>
    '''
    '''<param name="sTexte">Texte utilisé pour trouver le FeatureLayer par défaut.</param>
    '''<param name="bNonVisible"> Indique si on doit aussi extraire les FeatureLayers non visibles.</param>
    ''' 
    '''<returns>"IFeatureLayer" correspondant au texte recherché, sinon "Nothing".</returns>
    ''' 
    Public Function ExtraireFeatureLayerByName(ByVal sTexte As String, Optional ByVal bNonVisible As Boolean = False) As IFeatureLayer
        'Déclarer la variables de travail
        Dim qFeatureLayerColl As Collection = Nothing   'Contient la liste des FeatureLayer visibles
        Dim pFeatureLayer As IFeatureLayer = Nothing    'Interface contenant une classe de données

        'Définir la valeur de retour par défaut
        ExtraireFeatureLayerByName = Nothing

        Try
            'Définir la liste des FeatureLayer
            qFeatureLayerColl = DefinirCollectionFeatureLayer(bNonVisible)

            'Traiter tous les FeatureLayer
            For i = 1 To qFeatureLayerColl.Count
                'Définir le FeatureLayer
                pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                'Vérifier la présence du texte recherché pour la valeur par défaut
                If sTexte.ToUpper = pFeatureLayer.Name.ToUpper Then
                    'Définir le featurelayer
                    ExtraireFeatureLayerByName = pFeatureLayer
                    'Sortir
                    Exit For
                End If
            Next

            'Vérifier si aucun trouvé
            If ExtraireFeatureLayerByName Is Nothing Then
                'Traiter tous les FeatureLayer
                For i = 1 To qFeatureLayerColl.Count
                    'Définir le FeatureLayer
                    pFeatureLayer = CType(qFeatureLayerColl.Item(i), IFeatureLayer)
                    'Vérifier la présence du texte recherché pour la valeur par défaut
                    If InStr(pFeatureLayer.Name.ToUpper, sTexte.ToUpper) > 0 Then
                        'Définir le featurelayer
                        ExtraireFeatureLayerByName = pFeatureLayer
                        'Sortir
                        Exit For
                    End If
                Next
            End If

        Catch erreur As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pFeatureLayer = Nothing
            qFeatureLayerColl = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de définir la collection des RasterLayers présents dans la Map.
    ''' On peut indiquer si on veut aussi extraire les RasterLayers non visibles.
    '''</summary>
    ''' 
    '''<param name="bNonVisible"> Indique si on doit aussi extraire les RasterLayer non visibles.</param>
    ''' 
    '''<return>"Collection" contenant les "IRasterLayer" visible ou non selon ce qui est demandé.</return>
    ''' 
    Public Function DefinirCollectionRasterLayer(ByVal bNonVisible As Boolean) As Collection
        'Déclarer les variables de travail
        Dim pActiveView As IActiveView = Nothing                'Interface contenant la fenêtre active.
        Dim pLayer As ILayer = Nothing                          'Interface contenant un Layer.
        Dim pGroupLayer As IGroupLayer = Nothing                'Interface contenant un Groupe de Layers.
        Dim pRasterLayer As IRasterLayer = Nothing              'Interface contenant un RasterLayer.
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le Layer d'un RasterCatalog.
        Dim pFeatureCursor As IFeatureCursor = Nothing          'Interface pour extraire les items d'un RasterCatalog.
        Dim pSpatialFilter As ISpatialFilter = Nothing          'Interface utilisé pour sélectionner les Rasters du Catalogue dans la vue active.
        Dim pRasterCatalogItem As IRasterCatalogItem = Nothing  'Interface contenant un Item du RasterCatalog.

        'Retourner le résultat par défaut
        DefinirCollectionRasterLayer = gqRasterLayerCollection

        Try
            'Définir la collection des RasterLayer vide
            gqRasterLayerCollection = New Collection
            gpEnvelope = Nothing

            'Définir la vue active
            pActiveView = CType(gpMap, IActiveView)

            'Traiter tous les Layers
            For i = 0 To gpMap.LayerCount - 1
                'Définir le Layer à traiter
                pLayer = gpMap.Layer(i)

                'Vérifier si on tient on doit extraire le Layer même s'il n'est pas visible
                If pLayer.Visible = True Or bNonVisible = True Then
                    'Vérifier le Layer est un RasterLayer
                    If TypeOf pLayer Is IRasterLayer Then
                        'Définir le RasterLayer
                        pRasterLayer = CType(pLayer, IRasterLayer)

                        'Vérifier la présence du Raster
                        If Not pRasterLayer.Raster Is Nothing Then
                            'Vérifier si le RasterLayer est absent de la collection
                            If Not gqRasterLayerCollection.Contains(pRasterLayer.Name) Then
                                'Ajouter un nouveau RasterLayer dans la collection
                                gqRasterLayerCollection.Add(pRasterLayer, pRasterLayer.Name)
                                'Ajuster l'envelope selon la collection des RasterLayer
                                If gpEnvelope Is Nothing Then
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope = pRasterLayer.AreaOfInterest
                                Else
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope.Union(pRasterLayer.AreaOfInterest)
                                End If
                            End If
                        End If

                        'Si l'object est un GdbRasterCatalogLayer (Catalogue d'images)
                    ElseIf TypeOf (pLayer) Is IGdbRasterCatalogLayer Then
                        'Interface pour extraire le FeatureCatalog
                        pFeatureLayer = CType(pLayer, IFeatureLayer)

                        'Créer une nouvelle requête spatiale
                        pSpatialFilter = New SpatialFilter

                        'Définir la requête spatiale
                        pSpatialFilter.Geometry = pActiveView.Extent
                        pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                        pSpatialFilter.OutputSpatialReference(pFeatureLayer.FeatureClass.ShapeFieldName) = gpMap.SpatialReference
                        pSpatialFilter.GeometryField = pFeatureLayer.FeatureClass.ShapeFieldName

                        'Interface pour extraire les Raster
                        pFeatureCursor = pFeatureLayer.Search(pSpatialFilter, True)

                        'Extraire le premier item du RasterCatalog
                        pRasterCatalogItem = CType(pFeatureCursor.NextFeature, IRasterCatalogItem)

                        'Traiter tous les Raster
                        Do Until pRasterCatalogItem Is Nothing
                            'Créer un nouveau RasterLayer vide
                            pRasterLayer = New RasterLayer

                            'Définir le RasterLayer à partir du Catalogue
                            pRasterLayer.CreateFromDataset(CType(pRasterCatalogItem.RasterDataset, IRasterDataset))

                            'Vérifier si le RasterLayer est absent de la collection
                            If Not gqRasterLayerCollection.Contains(pRasterLayer.Name) Then
                                'Ajouter un nouveau RasterLayer dans la collection
                                gqRasterLayerCollection.Add(pRasterLayer, pRasterLayer.Name)
                                'Ajuster l'envelope selon la collection des RasterLayer
                                If gpEnvelope Is Nothing Then
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope = pRasterLayer.AreaOfInterest
                                Else
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope.Union(pRasterLayer.AreaOfInterest)
                                End If
                            End If

                            'Extraire le prochain item du RasterCatalog
                            pRasterCatalogItem = CType(pFeatureCursor.NextFeature, IRasterCatalogItem)
                        Loop

                        'Vérifier les autres Layer dans un GroupLayer
                    ElseIf TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer = CType(pLayer, IGroupLayer)

                        'Trouver les autres RasterLayer dans un GroupLayer
                        Call DefinirCollectionRasterLayerGroup(gqRasterLayerCollection, pGroupLayer, bNonVisible)
                    End If
                End If
            Next

            'Retourner le résultat par défaut
            DefinirCollectionRasterLayer = gqRasterLayerCollection

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pActiveView = Nothing
            pLayer = Nothing
            pGroupLayer = Nothing
            pRasterLayer = Nothing
            pFeatureLayer = Nothing
            pFeatureCursor = Nothing
            pSpatialFilter = Nothing
            pRasterCatalogItem = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet de définir la collection des FeatureLayers présents dans la Map.
    ''' On peut indiquer si on veut aussi extraire les FeatureLayers non visibles.
    '''</summary>
    ''' 
    '''<param name="bNonVisible"> Indique si on doit aussi extraire les FeatureLayers non visibles.</param>
    '''<param name="pEsriGeometryType"> Contient le type de géométrie des FeatureLayers recherchés.</param>
    ''' 
    '''<return>"Collection" contenant les "IFeatureLayer" visible ou non selon ce qui est demandé.</return>
    ''' 
    Public Function DefinirCollectionFeatureLayer(ByVal bNonVisible As Boolean, _
    Optional ByVal pEsriGeometryType As esriGeometryType = esriGeometryType.esriGeometryAny) As Collection
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer As IGroupLayer = Nothing            'Interface contenant un Groupe de Layers
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant un FeatureLayer

        'Retourner le résultat par défaut
        DefinirCollectionFeatureLayer = gqFeatureLayerCollection

        Try
            'Définir la collection de FeatureLayer vide
            gqFeatureLayerCollection = New Collection
            gpEnvelope = Nothing

            'Traiter tous les Layers
            For i = 0 To gpMap.LayerCount - 1
                'Définir le Layer à traiter
                pLayer = gpMap.Layer(i)

                'Vérifier si on tient on doit extraire le Layer même s'il n'est pas visible
                If pLayer.Visible = True Or bNonVisible = True Then
                    'Vérifier le Layer est un FeatureLayer
                    If TypeOf pLayer Is IFeatureLayer Then
                        'Définir le FeatureLayer
                        pFeatureLayer = CType(pLayer, IFeatureLayer)

                        'Vérifier la présence de la FeatureClass
                        If Not pFeatureLayer.FeatureClass Is Nothing Then
                            'Vérifier le type de géométrie correspond à ce qui est recherché
                            If pEsriGeometryType = esriGeometryType.esriGeometryAny _
                            Or pFeatureLayer.FeatureClass.ShapeType = pEsriGeometryType Then
                                'Vérifier si le FeatureLayer est absent de la collection
                                If Not gqFeatureLayerCollection.Contains(pFeatureLayer.Name) Then
                                    'Ajouter un nouveau FeatureLayer dans la collection
                                    gqFeatureLayerCollection.Add(pFeatureLayer, pFeatureLayer.Name)
                                    'Ajuster l'envelope selon la collection des FeatureLayers
                                    If gpEnvelope Is Nothing Then
                                        'Conserver l'envelope du FeatureLayer
                                        gpEnvelope = pFeatureLayer.AreaOfInterest
                                    Else
                                        'Conserver l'envelope du FeatureLayer
                                        gpEnvelope.Union(pFeatureLayer.AreaOfInterest)
                                    End If
                                End If
                            End If
                        End If

                        'Vérifier les autres Layer dans un GroupLayer
                    ElseIf TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer = CType(pLayer, IGroupLayer)

                        'Trouver les autres FeatureLayer dans un GroupLayer
                        Call DefinirCollectionFeatureLayerGroup(gqFeatureLayerCollection, pGroupLayer, bNonVisible)
                    End If
                End If
            Next

            'Retourner le résultat par défaut
            DefinirCollectionFeatureLayer = gqFeatureLayerCollection

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
            pFeatureLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'extraire le GroupLayer dans lequel le Layer recherché est présent.
    '''</summary>
    '''
    '''<param name="pLayerRechercher">Interface contenant le Layer à rechercher dans la Map active.</param>
    '''<param name="nPosition">Position su Layer dans le GroupLayer.</param>
    ''' 
    '''<returns>"Collection" contenant les "IGroupLayer" recherchés.</returns>
    ''' 
    Public Function DefinirGroupLayer(ByVal pLayerRechercher As ILayer, ByRef nPosition As Integer) As IGroupLayer
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing              'Interface contenant un Layer
        Dim pGroupLayer As IGroupLayer = Nothing    'Interface contenant un GroupLayer

        'Définir la valeur par défaut
        DefinirGroupLayer = Nothing

        Try
            'Vérifier si le Layer est valide
            If pLayerRechercher Is Nothing Then Return Nothing

            'Traiter tous les Layers
            For i = 0 To gpMap.LayerCount - 1
                'Définir le Layer à traiter
                pLayer = gpMap.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayer IsNot pLayerRechercher Then
                    'Vérifier les autres Layer dans un GroupLayer
                    If TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer = CType(pLayer, IGroupLayer)

                        'Trouver les autres GroupLayer dans un GroupLayer
                        DefinirGroupLayer = ExtraireGroupLayerGroup(pGroupLayer, pLayerRechercher, nPosition)
                    End If

                    'Sortir de la boucle si le GroupLayer a été trouvé
                    If Not DefinirGroupLayer Is Nothing Then Exit For
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'indiquer si le FeatureLayer est visible ou non dans la IMap.
    '''</summary>
    ''' 
    '''<param name="pLayerRechercher"> Interface ESRI contenant le Layer à rechercher.</param>
    '''<param name="bPresent"> Contient l'indication si le Layer à rechercher est présent dans la Map.</param>
    ''' 
    '''<return>"Collection" contenant les "IFeatureLayer" visible ou non.</return>
    ''' 
    Public Function EstVisible(ByVal pLayerRechercher As ILayer, ByVal bPresent As Boolean) As Boolean
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing              'Interface contenant un Layer
        Dim pGroupLayer As IGroupLayer = Nothing    'Interface contenant un Groupe de Layers

        Try
            'Traiter tous les Layers
            For i = 0 To gpMap.LayerCount - 1
                'Définir le Layer à traiter
                pLayer = gpMap.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayer Is pLayerRechercher Then
                    'Retourner l'indication s'il est visible ou non
                    EstVisible = pLayer.Visible

                    'Sortir de la recherche
                    Exit For

                    'Si ce n'est pas le Layer recherché et que c'est un GroupLayer
                ElseIf TypeOf pLayer Is IGroupLayer Then
                    'Définir le GroupLayer
                    pGroupLayer = CType(pLayer, IGroupLayer)

                    'Retourner l'indication s'il est visible ou non
                    EstVisible = EstVisibleGroup(pGroupLayer, pLayer, bPresent)

                    'Sortir si le Layer est présent dans le GroupLayer
                    If bPresent Then Exit For
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
        End Try
    End Function
#End Region

#Region "Routine et fonction privées"
    '''<summary>
    ''' Routine qui permet de définir la collection des RasterLayers contenus dans un GroupLayer.
    ''' On peut indiquer si on veut aussi extraire les RasterLayers non visibles.
    '''</summary>
    ''' 
    '''<param name="qRasterLayerColl">Collection des RasterLayer.</param>
    '''<param name="pGroupLayer">Interface ESRI contenant un group de Layers.</param>
    '''<param name="bNonVisible">Indique si on doit aussi extraire les RasterLayers non visibles.</param>
    ''' 
    Private Sub DefinirCollectionRasterLayerGroup(ByRef qRasterLayerColl As Collection, ByVal pGroupLayer As IGroupLayer, ByVal bNonVisible As Boolean)
        'Déclarer les variables de travail
        Dim pActiveView As IActiveView = Nothing                'Interface contenant la fenêtre active.
        Dim pLayer As ILayer = Nothing                          'Interface contenant un Layer
        Dim pGroupLayer2 As IGroupLayer = Nothing               'Interface contenant un GroupLayer
        Dim pRasterLayer As IRasterLayer = Nothing              'Interface contenant un RasterLayer
        Dim pCompositeLayer As ICompositeLayer = Nothing        'Interface utiliser pour extraire un Layer dans un GroupLayer
        Dim pFeatureLayer As IFeatureLayer = Nothing            'Interface contenant le Layer d'un RasterCatalog.
        Dim pFeatureCursor As IFeatureCursor = Nothing          'Interface pour extraire les items d'un RasterCatalog.
        Dim pSpatialFilter As ISpatialFilter = Nothing          'Interface utilisé pour sélectionner les Rasters du Catalogue dans la vue active.
        Dim pRasterCatalogItem As IRasterCatalogItem = Nothing  'Interface contenant un Item du RasterCatalog.

        Try
            'Interface pour accéder aux Layers du GroupLayer
            pCompositeLayer = CType(pGroupLayer, ICompositeLayer)

            'Définir la vue active
            pActiveView = CType(gpMap, IActiveView)

            'Trouver le Groupe de Layer
            For i = 0 To pCompositeLayer.Count - 1
                'Interface pour comparer le nom du Layer
                pLayer = pCompositeLayer.Layer(i)

                'Vérifier si on tient compte du selectable
                If pLayer.Visible = True Or bNonVisible = True Then
                    'Vérifier le Layer est un RasterLayer
                    If TypeOf pLayer Is IRasterLayer Then
                        'Définir le RasterLayer
                        pRasterLayer = CType(pLayer, IRasterLayer)

                        'Vérifier la présence du Raster
                        If Not pRasterLayer.Raster Is Nothing Then
                            'Vérifier si le RasterLayer est absent de la collection
                            If Not gqRasterLayerCollection.Contains(pRasterLayer.Name) Then
                                'Ajouter un nouveau RasterLayer dans la collection
                                gqRasterLayerCollection.Add(pRasterLayer, pRasterLayer.Name)
                                'Ajuster l'envelope selon la collection des RasterLayers
                                If gpEnvelope Is Nothing Then
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope = pRasterLayer.AreaOfInterest
                                Else
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope.Union(pRasterLayer.AreaOfInterest)
                                End If
                            End If
                        End If

                        'Si l'object est un GdbRasterCatalogLayer (Catalogue d'images)
                    ElseIf TypeOf (pLayer) Is IGdbRasterCatalogLayer Then
                        'Interface pour extraire le FeatureCatalog
                        pFeatureLayer = CType(pLayer, IFeatureLayer)

                        'Créer une nouvelle requête spatiale
                        pSpatialFilter = New SpatialFilter

                        'Définir la requête spatiale
                        pSpatialFilter.Geometry = pActiveView.Extent
                        pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                        pSpatialFilter.OutputSpatialReference(pFeatureLayer.FeatureClass.ShapeFieldName) = gpMap.SpatialReference
                        pSpatialFilter.GeometryField = pFeatureLayer.FeatureClass.ShapeFieldName

                        'Interface pour extraire les Raster
                        pFeatureCursor = pFeatureLayer.Search(pSpatialFilter, True)

                        'Extraire le premier item du RasterCatalog
                        pRasterCatalogItem = CType(pFeatureCursor.NextFeature, IRasterCatalogItem)

                        'Traiter tous les Raster
                        Do Until pRasterCatalogItem Is Nothing
                            'Créer un nouveau RasterLayer vide
                            pRasterLayer = New RasterLayer

                            'Définir le RasterLayer à partir du Catalogue
                            pRasterLayer.CreateFromDataset(CType(pRasterCatalogItem.RasterDataset, IRasterDataset))

                            'Vérifier si le RasterLayer est absent de la collection
                            If Not gqRasterLayerCollection.Contains(pRasterLayer.Name) Then
                                'Ajouter un nouveau RasterLayer dans la collection
                                gqRasterLayerCollection.Add(pRasterLayer, pRasterLayer.Name)
                                'Ajuster l'envelope selon la collection des RasterLayer
                                If gpEnvelope Is Nothing Then
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope = pRasterLayer.AreaOfInterest
                                Else
                                    'Conserver l'envelope du RasterLayer
                                    gpEnvelope.Union(pRasterLayer.AreaOfInterest)
                                End If
                            End If

                            'Extraire le prochain item du RasterCatalog
                            pRasterCatalogItem = CType(pFeatureCursor.NextFeature, IRasterCatalogItem)
                        Loop

                        'Vérifier les autres noms dans un GroupLayer
                    ElseIf TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer2 = CType(pLayer, IGroupLayer)

                        'Trouver les autres RasterLayer dans un GroupLayer
                        Call DefinirCollectionFeatureLayerGroup(qRasterLayerColl, pGroupLayer2, bNonVisible)
                    End If
                End If
            Next i

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pActiveView = Nothing
            pLayer = Nothing
            pGroupLayer2 = Nothing
            pRasterLayer = Nothing
            pCompositeLayer = Nothing
            pFeatureLayer = Nothing
            pFeatureCursor = Nothing
            pSpatialFilter = Nothing
            pRasterCatalogItem = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet de définir la collection des FeatureLayers contenus dans un GroupLayer.
    ''' On peut indiquer si on veut aussi extraire les FeatureLayers non visibles.
    '''</summary>
    ''' 
    '''<param name="qFeatureLayerColl">Collection des FeatureLayer.</param>
    '''<param name="pGroupLayer">Interface ESRI contenant un group de Layers.</param>
    '''<param name="bNonVisible">Indique si on doit aussi extraire les FeatureLayers non visibles.</param>
    '''<param name="pEsriGeometryType"> Contient le type de géométrie des FeatureLayers recherchés.</param>
    ''' 
    Private Sub DefinirCollectionFeatureLayerGroup(ByRef qFeatureLayerColl As Collection, ByVal pGroupLayer As IGroupLayer, _
    ByVal bNonVisible As Boolean, Optional ByVal pEsriGeometryType As esriGeometryType = esriGeometryType.esriGeometryAny)
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer2 As IGroupLayer = Nothing           'Interface contenant un GroupLayer
        Dim pFeatureLayer As IFeatureLayer = Nothing        'Interface contenant un FeatureLayer
        Dim pCompositeLayer As ICompositeLayer = Nothing    'Interface utiliser pour extraire un Layer dans un GroupLayer

        Try
            'Interface pour accéder aux Layers du GroupLayer
            pCompositeLayer = CType(pGroupLayer, ICompositeLayer)

            'Trouver le Groupe de Layer
            For i = 0 To pCompositeLayer.Count - 1
                'Interface pour comparer le nom du Layer
                pLayer = pCompositeLayer.Layer(i)

                'Vérifier si on tient compte du selectable
                If pLayer.Visible = True Or bNonVisible = True Then
                    'Vérifier le Layer est un FeatureLayer
                    If TypeOf pLayer Is IFeatureLayer Then
                        'Définir le FeatureLayer
                        pFeatureLayer = CType(pLayer, IFeatureLayer)

                        'Vérifier la présence de la FeatureClass
                        If Not pFeatureLayer.FeatureClass Is Nothing Then
                            'Vérifier le type de géométrie correspond à ce qui est recherché
                            If pEsriGeometryType = esriGeometryType.esriGeometryAny _
                            Or pFeatureLayer.FeatureClass.ShapeType = pEsriGeometryType Then
                                'Vérifier si le FeatureLayer est absent de la collection
                                If Not gqFeatureLayerCollection.Contains(pFeatureLayer.Name) Then
                                    'Ajouter un nouveau FeatureLayer dans la collection
                                    gqFeatureLayerCollection.Add(pFeatureLayer, pFeatureLayer.Name)
                                    'Ajuster l'envelope selon la collection des FeatureLayers
                                    If gpEnvelope Is Nothing Then
                                        'Conserver l'envelope du FeatureLayer
                                        gpEnvelope = pFeatureLayer.AreaOfInterest
                                    Else
                                        'Conserver l'envelope du FeatureLayer
                                        gpEnvelope.Union(pFeatureLayer.AreaOfInterest)
                                    End If
                                End If
                            End If
                        End If

                        'Vérifier les autres noms dans un GroupLayer
                    ElseIf TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer2 = CType(pLayer, IGroupLayer)

                        'Trouver les autres FeatureLayer dans un GroupLayer
                        Call DefinirCollectionFeatureLayerGroup(qFeatureLayerColl, pGroupLayer2, bNonVisible)
                    End If
                End If
            Next i

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer2 = Nothing
            pFeatureLayer = Nothing
            pCompositeLayer = Nothing
        End Try
    End Sub

    '''<summary>
    ''' Routine qui permet d'extraire le GroupLayer contenant un FeatureLayer.
    ''' Le GroupLayer recherché est contenu dans un GroupLayer.
    '''</summary>
    ''' 
    '''<param name="pGroupLayer">Interface ESRI contenant un groupe de Layers.</param>
    '''<param name="pLayerRechercher">Interface ESRI contenant le Layer à rechercher.</param>
    '''<param name="nPosition">Position su Layer dans le GroupLayer.</param>
    ''' 
    Private Function ExtraireGroupLayerGroup(ByVal pGroupLayer As IGroupLayer, ByVal pLayerRechercher As ILayer, ByRef nPosition As Integer) As IGroupLayer
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer2 As IGroupLayer = Nothing           'Interface contenant un GroupLayer
        Dim pCompositeLayer As ICompositeLayer = Nothing    'Interface utiliser pour extraire un Layer dans un GroupLayer

        'Initialiser les variables de travail
        ExtraireGroupLayerGroup = Nothing

        Try
            'Interface pour accéder aux Layers du GroupLayer
            pCompositeLayer = CType(pGroupLayer, ICompositeLayer)

            'Trouver le Groupe de Layer
            For i = 0 To pCompositeLayer.Count - 1
                'Interface pour comparer le nom du Layer
                pLayer = pCompositeLayer.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayerRechercher Is pLayer Then
                    'Retourner le Groupe du Layer recherché
                    ExtraireGroupLayerGroup = pGroupLayer
                    'Définir la position du Layer dans le GroupLayer
                    nPosition = i
                    'Sortir
                    Exit For
                Else
                    'Vérifier les autres noms dans un GroupLayer
                    If TypeOf pLayer Is IGroupLayer Then
                        'Définir le GroupLayer
                        pGroupLayer2 = CType(pLayer, IGroupLayer)

                        'Trouver les autres GroupLayer dans un GroupLayer
                        ExtraireGroupLayerGroup = ExtraireGroupLayerGroup(pGroupLayer2, pLayerRechercher, nPosition)
                    End If

                    'Sortir
                    If Not ExtraireGroupLayerGroup Is Nothing Then Exit For
                End If
            Next i

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer2 = Nothing
            pCompositeLayer = Nothing
        End Try
    End Function

    '''<summary>
    ''' Fonction qui permet d'indiquer si le FeatureLayer est visible ou non dans la IMap.
    '''</summary>
    ''' 
    '''<param name="pGroupLayer">Interface ESRI contenant un group de Layers.</param>
    '''<param name="pLayerRechercher"> Interface ESRI contenant le Layer à rechercher.</param>
    '''<param name="bPresent"> Contient l'indication si le Layer à rechercher est présent dans la Map.</param>
    ''' 
    '''<return>"Collection" contenant les "IFeatureLayer" visible ou non.</return>
    ''' 
    Private Function EstVisibleGroup(ByVal pGroupLayer As IGroupLayer, ByVal pLayerRechercher As ILayer, ByVal bPresent As Boolean) As Boolean
        'Déclarer les variables de travail
        Dim pLayer As ILayer = Nothing                      'Interface contenant un Layer
        Dim pGroupLayer2 As IGroupLayer = Nothing           'Interface contenant un Groupe de Layers
        Dim pCompositeLayer As ICompositeLayer = Nothing    'Interface utiliser pour extraire un Layer dans un GroupLayer

        Try
            'Interface pour accéder aux Layers du GroupLayer
            pCompositeLayer = CType(pGroupLayer, ICompositeLayer)

            'Trouver le Groupe de Layer
            For i = 0 To pCompositeLayer.Count - 1
                'Interface pour comparer le nom du Layer
                pLayer = pCompositeLayer.Layer(i)

                'Vérifier si le Layer trouvé est le même que celui recherché
                If pLayer Is pLayerRechercher Then
                    'Retourner l'indication s'il est visible ou non
                    EstVisibleGroup = pLayer.Visible

                    'Sortir de la recherche
                    Exit For

                    'Si ce n'est pas le Layer recherché et que c'est un GroupLayer
                ElseIf TypeOf pLayer Is IGroupLayer Then
                    'Définir le GroupLayer
                    pGroupLayer2 = CType(pLayer, IGroupLayer)

                    'Retourner l'indication s'il est visible ou non
                    EstVisibleGroup = EstVisibleGroup(pGroupLayer2, pLayer, bPresent)

                    'Sortir si le Layer est présent dans le GroupLayer
                    If bPresent Then Exit For
                End If
            Next

        Catch e As Exception
            'Message d'erreur
            Throw
        Finally
            'Vider la mémoire
            pLayer = Nothing
            pGroupLayer = Nothing
            pCompositeLayer = Nothing
        End Try
    End Function
#End Region
End Class
