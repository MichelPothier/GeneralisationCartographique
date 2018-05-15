Imports System
Imports System.Collections.Generic
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase

''' <summary>
''' Classe qui permet d'identifier et corriger la généralisation d'une géométrie.
''' </summary>
Public Class clsGeneraliserGeometrie
    Inherits TriangulationDelaunay

#Region "Routines et fonctions publiques"
    ''' <summary>
    ''' Routine qui permet de retourner le Polygone de généralisation intérieure et la Polyline d'erreur de généralisation d'un polygone.
    ''' </summary>
    ''' 
    '''<param name="pPolygon">Polygone utilisée pour effectuer la généralisation.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre le polygon et les éléments en relation.</param>
    '''<param name="dDistLat"> Distance latérale utilisée pour éliminer des sommets en trop.</param>
    '''<param name="dLargMin"> Largeur minimum utilisée pour généraliser.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour généraliser.</param>
    '''<param name="dSupMin">Contient la superficie de généralisation minimum.</param>
    '''<param name="pPolygonGen"> Interface contenant le polygone de généralisation.</param>
    '''<param name="pPolylineErr"> Interface contenant la polyligne d'erreur de généralisation.</param>
    '''<param name="pSquelette"> Interface contenant le squelette du polygon.</param>
    '''<param name="pBagDroites"> Interface contenant les droites de Delaunay.</param>
    ''' 
    ''' <remarks>La généralisation est effectuée à partir des lignes des triangles de Delaunay.</remarks>
    ''' 
    Public Shared Sub GeneraliserInterieurPolygone(ByVal pPolygon As IPolygon4, ByVal pPointsConnexion As IMultipoint,
                                                   ByVal dDistLat As Double, ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                                   ByRef pPolygonGen As IPolygon, ByRef pPolylineErr As IPolyline,
                                                   ByRef pSquelette As IPolyline, ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les anneaux extérieurs.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface pour ajouter les lignes du squelette.
        Dim pGeomCollAddD As IGeometryCollection = Nothing      'Interface pour ajouter les droites de Delaunay.
        Dim pGeomCollAddL As IGeometryCollection = Nothing      'Interface pour ajouter les lignes de généralisation intérieures.
        Dim pGeomCollAddS As IGeometryCollection = Nothing      'Interface pour ajouter les polygones de généralisation intérieures.
        Dim pRingExt As IRing = Nothing                         'Interface contenant l'anneau extérieur.
        Dim pPolygonTmp As IPolygon = Nothing                   'Interface contenant le polygone temporaire de traitement.
        Dim pPolygonGenTmp As IPolygon = Nothing                'Interface contenant le polygone de généralisation temporaire.
        Dim pPolylineErrTmp As IPolyline = Nothing              'Interface contenant la polyligne d'erreur de généralisation temporaire.
        Dim pSqueletteTmp As IPolyline = Nothing                'Interface contenant la polyligne du squelette temporaire.
        Dim pBagDroitesTmp As IGeometryBag = Nothing            'Interface contenant les droites de Delaunay temporaires.
        Dim pGeomCollTmp As IGeometryCollection = Nothing       'Interface pour ajouter les anneaux extérieurs dans le polygone temporaire.
        Dim pPointsConnexionTmp As IMultipoint = Nothing        'Interface contenant les points de connexion temporaire.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface utilisé pour extraire les points de connexion temporaire.

        Try
            'Enlever les sommets en trop
            'pPolygon.Generalize(dDistLat)

            'Densifer les sommets du polygone
            pPolygon.Densify(dLargMin / 2, 0)

            'Créer la polyligne du squelette vide
            pSquelette = New Polyline
            pSquelette.SpatialReference = pPolygon.SpatialReference

            'Créer la polyligne d'erreur de généralisation vide
            pPolylineErr = New Polyline
            pPolylineErr.SpatialReference = pPolygon.SpatialReference

            'Créer le polygone de généralisation vide
            pPolygonGen = New Polygon
            pPolygonGen.SpatialReference = pPolygon.SpatialReference

            'Créer le Bag des droites de Delaunay vide
            pBagDroites = New GeometryBag
            pBagDroites.SpatialReference = pPolygon.SpatialReference

            'Vérifier si le polygone n'est pas vide
            If Not pPolygon.IsEmpty Then
                'Interface pour ajouter les lignes du squelette
                pGeomCollAdd = CType(pSquelette, IGeometryCollection)
                'Interface pour ajouter les lignes d'erreur de généralisation
                pGeomCollAddL = CType(pPolylineErr, IGeometryCollection)
                'Interface pour ajouter les anneaux de généralisation
                pGeomCollAddS = CType(pPolygonGen, IGeometryCollection)
                'Interface pour ajouter les droites des lignes de Delaunay
                pGeomCollAddD = CType(pBagDroites, IGeometryCollection)
                'Interface pour extraire les anneaux extérieurs
                pGeomColl = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

                'Traiter toutes les composantes
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Définir la composante
                    pRingExt = CType(pGeomColl.Geometry(i), IRing)

                    'Vérifier si l'anneau n'est pas vide
                    If Not pRingExt.IsEmpty Then
                        'Créer un nouveau polygone vide
                        pPolygonTmp = New Polygon
                        pPolygonTmp.SpatialReference = pPolygon.SpatialReference
                        'Ajouter l'anneau extérieur
                        pGeomCollTmp = CType(pPolygonTmp, IGeometryCollection)
                        pGeomCollTmp.AddGeometry(pRingExt)
                        'Ajouter les anneaux intérieures
                        pGeomCollTmp.AddGeometryCollection(CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection))

                        'Interface pour extraire les points d'intersection spécifique au polygone temporaire
                        pTopoOp = CType(pPointsConnexion, ITopologicalOperator2)
                        'Extraire les points d'intersection spécifique au polygone temporaire
                        pPointsConnexionTmp = CType(pTopoOp.Intersect(pPolygonTmp, esriGeometryDimension.esriGeometry0Dimension), IMultipoint)

                        'Généralisation intérieure des polygones temporaires
                        Call GeneraliserAnneauInterieur(pPolygonTmp, pPointsConnexionTmp, dLargMin, dLongMin, dSupMin, pPolygonGenTmp, pPolylineErrTmp, pSqueletteTmp, pBagDroitesTmp)

                        'Ajouter les lignes du squelette
                        pGeomCollAdd.AddGeometryCollection(CType(pSqueletteTmp, IGeometryCollection))

                        'Ajouter les lignes de généralisation
                        pGeomCollAddL.AddGeometryCollection(CType(pPolylineErrTmp, IGeometryCollection))

                        'Ajouter les polygones de généralisation
                        pGeomCollAddS.AddGeometryCollection(CType(pPolygonGenTmp, IGeometryCollection))

                        'Ajouter les droites de Delaunay
                        pGeomCollAddD.AddGeometryCollection(CType(pBagDroitesTmp, IGeometryCollection))
                    End If
                Next
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pGeomCollAddD = Nothing
            pGeomCollAddL = Nothing
            pGeomCollAddS = Nothing
            pRingExt = Nothing
            pPointsConnexionTmp = Nothing
            pPolygonTmp = Nothing
            pPolygonGenTmp = Nothing
            pPolylineErrTmp = Nothing
            pSqueletteTmp = Nothing
            pBagDroitesTmp = Nothing
            pGeomCollTmp = Nothing
            pPointsConnexionTmp = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de retourner le Polygone de généralisation extérieure et la Polyline d'erreur de généralisation d'un polygon.
    ''' </summary>
    ''' 
    '''<param name="pPolygon">Polygone utilisée pour effectuer la généralisation.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre le polygon et les éléments en relation.</param>
    '''<param name="dDistLat"> Distance latérale utilisée pour éliminer des sommets en trop.</param>
    '''<param name="dLargMin"> Largeur minimum utilisée pour généraliser.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour généraliser.</param>
    '''<param name="dSupMin"> Contient la superficie de généralisation minimum.</param>
    '''<param name="pPolygonGen"> Interface contenant le polygone de généralisation.</param>
    '''<param name="pPolylineErr"> Interface contenant la polyligne d'erreur de généralisation.</param>
    '''<param name="pSquelette"> Interface contenant le squelette du polygon.</param>
    '''<param name="pBagDroites"> Interface contenant les droites de Delaunay.</param>
    ''' 
    ''' <remarks>La généralisation est effectuée à partir des lignes des triangles de Delaunay.</remarks>
    ''' 
    Public Shared Sub GeneraliserExterieurPolygone(ByVal pPolygon As IPolygon, ByVal pPointsConnexion As IMultipoint,
                                                   ByVal dDistLat As Double, ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                                   ByRef pPolygonGen As IPolygon, ByRef pPolylineErr As IPolyline,
                                                   ByRef pSquelette As IPolyline, ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pPolygonExt As IPolygon4 = Nothing                  'Interface contenant le polygone extérieure générer à partir du polygone.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour effectuer des opérations spatiales.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les anneaux extérieurs.
        Dim pGeomCollAddD As IGeometryCollection = Nothing      'Interface pour ajouter les lignes des droites de Delaunay.
        Dim pGeomCollAddQ As IGeometryCollection = Nothing      'Interface pour ajouter les lignes du squelette.
        Dim pGeomCollAddL As IGeometryCollection = Nothing      'Interface pour ajouter les lignes de généralisation extérieures.
        Dim pGeomCollAddS As IGeometryCollection = Nothing      'Interface pour ajouter les polygones de généralisation extérieures.
        Dim pRingExt As IRing = Nothing                         'Interface contenant l'anneau extérieur.
        Dim pPolygonTmp As IPolygon = Nothing                   'Interface contenant le polygone temporaire de traitement.
        Dim pPolygonGenTmp As IPolygon = Nothing                'Interface contenant le polygone de généralisation temporaire.
        Dim pPolylineErrTmp As IPolyline = Nothing              'Interface contenant la polyligne d'erreur de généralisation temporaire.
        Dim pSqueletteTmp As IPolyline = Nothing                'Interface contenant le squelette temporaire.
        Dim pBagDroitesTmp As IGeometryBag = Nothing            'Interface contenant les droites de Delaunay temporaires.
        Dim pGeomCollTmp As IGeometryCollection = Nothing       'Interface pour ajouter les anneaux extérieurs dans le polygone temporaire.
        Dim pPointsConnexionTmp As IMultipoint = Nothing        'Interface contenant les points de connexion temporaire.
        Dim pEnvelope As IPolygon = Nothing                     'Interface contenant le premier anneau du polygone extérieur.

        Try
            'Enlever les sommets en trop
            'pPolygon.Generalize(dDistLat)

            'Densifer les sommets du polygone
            pPolygon.Densify(dLargMin / 2, 0)

            'Créer la polyligne de généralisation extérieure vide
            pBagDroites = New GeometryBag
            pBagDroites.SpatialReference = pPolygon.SpatialReference

            'Créer la polyligne du squelette extérieure vide
            pSquelette = New Polyline
            pSquelette.SpatialReference = pPolygon.SpatialReference

            'Créer la polyligne de généralisation extérieure vide
            pPolylineErr = New Polyline
            pPolylineErr.SpatialReference = pPolygon.SpatialReference

            'Créer le polygone de généralisation extérieure vide
            pPolygonGen = New Polygon
            pPolygonGen.SpatialReference = pPolygon.SpatialReference

            'Vérifier si le polygone n'est pas vide
            If Not pPolygon.IsEmpty Then
                'Définir le polygone extérieur
                pPolygonExt = CreerPolygoneExterieur(pPolygon, dLargMin)

                'Interface pour ajouter les droites de Delaunay
                pGeomCollAddD = CType(pBagDroites, IGeometryCollection)
                'Interface pour ajouter les lignes du squelette
                pGeomCollAddQ = CType(pSquelette, IGeometryCollection)
                'Interface pour ajouter les lignes d'erreur de généralisation
                pGeomCollAddL = CType(pPolylineErr, IGeometryCollection)
                'Interface pour ajouter les anneaux de généralisation
                pGeomCollAddS = CType(pPolygonGen, IGeometryCollection)
                'Interface pour extraire les anneaux extérieurs
                pGeomColl = CType(pPolygonExt.ExteriorRingBag, IGeometryCollection)

                'Traiter toutes les composantes
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Définir la composante
                    pRingExt = CType(pGeomColl.Geometry(i), IRing)

                    'Vérifier si l'anneau n'est pas vide
                    If Not pRingExt.IsEmpty Then
                        'Créer un nouveau polygone vide
                        pPolygonTmp = New Polygon
                        pPolygonTmp.SpatialReference = pPolygon.SpatialReference
                        'Ajouter l'anneau extérieur
                        pGeomCollTmp = CType(pPolygonTmp, IGeometryCollection)
                        pGeomCollTmp.AddGeometry(pRingExt)
                        'Ajouter les anneaux intérieures
                        pGeomCollTmp.AddGeometryCollection(CType(pPolygonExt.InteriorRingBag(pRingExt), IGeometryCollection))

                        'Interface pour extraire les points d'intersection spécifique au polygone temporaire
                        pTopoOp = CType(pPointsConnexion, ITopologicalOperator2)
                        'Extraire les points d'intersection spécifique au polygone temporaire
                        pPointsConnexionTmp = CType(pTopoOp.Intersect(pPolygonTmp, esriGeometryDimension.esriGeometry0Dimension), IMultipoint)

                        'Généralisation extérieure des polygones temporaires
                        Call GeneraliserAnneauExterieur(pPolygonTmp, pPointsConnexion, dLargMin, dLongMin, dSupMin, pPolygonGenTmp, pPolylineErrTmp, pSqueletteTmp, pBagDroitesTmp)

                        'Ajouter les lignes de généralisation
                        pGeomCollAddD.AddGeometryCollection(CType(pBagDroitesTmp, IGeometryCollection))

                        'Ajouter les lignes du squelette
                        pGeomCollAddQ.AddGeometryCollection(CType(pSqueletteTmp, IGeometryCollection))

                        'Ajouter les lignes de généralisation
                        pGeomCollAddL.AddGeometryCollection(CType(pPolylineErrTmp, IGeometryCollection))

                        'Ajouter les polygones de généralisation
                        pGeomCollAddS.AddGeometryCollection(CType(pPolygonGenTmp, IGeometryCollection))
                    End If
                Next

                'Vérifier s'il n'y a pas de généralisation à effectuer
                If pPolylineErr.IsEmpty Then
                    'Définir le polygone initiale
                    pPolygonGen = pPolygon

                    'S'il y a de la généralisation à effectuer
                Else
                    'Interface pour extraire le premier anneau du polygone extérieur
                    pGeomColl = CType(pPolygonExt, IGeometryCollection)

                    'Créer un nouveau polygone vide
                    pEnvelope = New Polygon
                    pEnvelope.SpatialReference = pPolygon.SpatialReference

                    'Interface pour ajouter le premier anneau du polygone extérieur
                    pGeomCollAddS = CType(pEnvelope, IGeometryCollection)
                    'Ajouter le premier anneau du polygone extérieur
                    pGeomCollAddS.AddGeometry(pGeomColl.Geometry(0))

                    'Interface pour créer le polygone suite à la généralisation extérieure
                    pTopoOp = CType(pEnvelope, ITopologicalOperator2)
                    'Créer le polygone suite à la généralisation du polygone extérieur
                    pPolygonGen = CType(pTopoOp.Difference(pPolygonGen), IPolygon)
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPolygonExt = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pGeomCollAddD = Nothing
            pGeomCollAddQ = Nothing
            pGeomCollAddL = Nothing
            pGeomCollAddS = Nothing
            pRingExt = Nothing
            pPolygonTmp = Nothing
            pPolygonGenTmp = Nothing
            pPolylineErrTmp = Nothing
            pSqueletteTmp = Nothing
            pBagDroitesTmp = Nothing
            pGeomCollTmp = Nothing
            pPointsConnexionTmp = Nothing
            pEnvelope = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de retourner la Polyligne de généralisation et la Polyline d'erreurs du résultat de la généralisation effectuée pour une polyligne.
    ''' La généralisation est effectuée à partir des lignes des triangles de Delaunay.
    ''' </summary>
    ''' 
    '''<param name="pPolyline"> Polyligne utilisée pour effectuer la généralisation.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre la polyligne et les éléments en relation.</param>
    '''<param name="dDistLat"> Distance latérale utilisée pour éliminer des sommets en trop.</param>
    '''<param name="dLargGenMin"> Largeur minimum utilisée pour généraliser.</param>
    '''<param name="dLongGenMin"> Longueur minimale utilisée pour généraliser.</param>
    '''<param name="dLongMin"> Longueur minimale d'une ligne.</param>
    '''<param name="pPolylineGen"> Interface contenant les lignes de généralisation.</param>
    '''<param name="pPolylineErr"> Interface contenant les lignes de généralisation en erreur.</param>
    '''<param name="pSquelette"> Interface contenant le squelette de la polyligne.</param>
    '''<param name="pSqueletteEnv"> Interface contenant le squelette de la polyligne avec son enveloppe.</param>
    '''<param name="pBagDroites"> Interface contenant les droites des triangles de Delaunay.</param>
    '''<param name="pBagDroitesEnv"> Interface contenant les droites des triangles de Delaunay avec son enveloppe.</param>
    ''' 
    Public Shared Sub GeneraliserPolyligne(ByVal pPolyline As IPolyline, ByVal pPointsConnexion As IMultipoint, ByVal dDistLat As Double,
                                           ByVal dLargGenMin As Double, ByVal dLongGenMin As Double, ByVal dLongMin As Double, _
                                           ByRef pPolylineGen As IPolyline, ByRef pPolylineErr As IPolyline, ByRef pSquelette As IPolyline,
                                           ByRef pSqueletteEnv As IPolyline, ByRef pBagDroites As IGeometryBag, ByRef pBagDroitesEnv As IGeometryBag)
        'Déclarer les variables de travail
        Dim pPath As IPath = Nothing                            'Interface contenant une ligne à généraliser.
        Dim pPolylineTmp As IPolyline = Nothing                 'Interface contenant une ligne temporaire.
        Dim pPolylineErrTmp As IPolyline = Nothing              'Interface contenant une ligne d'erreur de généralisation temporaire.
        Dim pPolylineGenTmp As IPolyline = Nothing              'Interface contenant une ligne généralisée temporaire.
        Dim pDictLiens As New Dictionary(Of Integer, Noeud)     'Dictionnaire contenant l'information des sommets de la polyligne à traiter.
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionnaire contenant les numéros des droites à utiliser pour effectuer la généralisation.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les composantes d'une polyligne.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface pour ajouter les composantes d'une polyligne.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour soustraire une géométrie.

        Try
            'Enlever les sommets en trop
            'pPolyline.Generalize(dDistLat)

            'Densifer les sommets de la polyligne
            pPolyline.Densify(dLargGenMin / 2, 0)

            'Créer la ligne vide
            pPolylineErr = New Polyline
            pPolylineErr.SpatialReference = pPolyline.SpatialReference
            'Créer la ligne vide
            pPolylineGen = New Polyline
            pPolylineGen.SpatialReference = pPolyline.SpatialReference

            'Interface pour extraire les composantes d'une polyligne.
            pGeomColl = CType(pPolyline, IGeometryCollection)

            'Traiter toutes les composantes de la ligne
            For i = 0 To pGeomColl.GeometryCount - 1
                'Définir la ligne à généraliser
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Vérifier si la ligne est plus grande que la longueur minimum de généralisation
                If pPath.Length > dLongGenMin Then
                    'Créer la ligne vide
                    pPolylineTmp = New Polyline
                    pPolylineTmp.SpatialReference = pPolyline.SpatialReference
                    'Interface pour ajouter les composantes d'une polyligne
                    pGeomCollAdd = CType(pPolylineTmp, IGeometryCollection)
                    'Ajouter la ligne à la polyligne temporaire
                    pGeomCollAdd.AddGeometry(pPath)

                    'Créer le squelette de la ligne selon la triangulation de Delaunay
                    Call CreerSqueletteLigneDelaunay(pPolylineTmp, dLargGenMin, dLongGenMin, pSquelette, pDictLiens,
                                                     pSqueletteEnv, pBagDroites, pBagDroitesEnv)
                    'Call CreerSqueletteLigneDelaunay(pPath, dLargGenMin, dLongGenMin, pSquelette, pDictLiens,
                    '                                 pSqueletteEnv, pBagDroites, pBagDroitesEnv)

                    'Créer la polyligne d'erreur de généralisation
                    Call CreerPolyligneErreurGeneralisation(pPolylineTmp, pSquelette, pBagDroites, dLargGenMin, dLongGenMin, pPolylineErrTmp)

                    'Créer la polyligne de généralisation
                    Call CreerPolyligneGeneralisation(pPolylineTmp, pDictLiens, pBagDroitesEnv, pPolylineErrTmp, pPolylineGenTmp)

                    'Interface pour ajouter les composantes d'une polyligne
                    pGeomCollAdd = CType(pPolylineErr, IGeometryCollection)
                    'Ajouter la ligne à la polyligne d'erreur de généralisation
                    pGeomCollAdd.AddGeometryCollection(CType(pPolylineErrTmp, IGeometryCollection))

                    'Interface pour ajouter les composantes d'une polyligne
                    pGeomCollAdd = CType(pPolylineGen, IGeometryCollection)
                    'Ajouter la ligne à la polyligne de généralisation
                    pGeomCollAdd.AddGeometryCollection(CType(pPolylineGenTmp, IGeometryCollection))

                    'Si la longueur est plus petite
                Else
                    'Interface pour ajouter la ligne à la polyligne de généralisation
                    pGeomCollAdd = CType(pPolylineGen, IGeometryCollection)
                    'Ajouter la ligne à la polyligne de généralisation
                    pGeomCollAdd.AddGeometry(pPath)
                End If
            Next

            'Détruire les lignes non-connectées qui sont inférieures à la longueur minimale d'une ligne
            Call TraiterLongueurLigneMinimale(pPolylineGen, pPointsConnexion, dLongMin)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPath = Nothing
            pPolylineTmp = Nothing
            pPolylineErrTmp = Nothing
            pPolylineGenTmp = Nothing
            pDictLiens = Nothing
            pDictDroites = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de retourner la Polyligne de généralisation et la Polyline d'erreurs du résultat de la généralisation effectuée pour une polyligne.
    ''' La généralisation est effectuée à partir des lignes des triangles de Delaunay.
    ''' </summary>
    ''' 
    '''<param name="pPolyline"> Polyligne utilisée pour effectuer la généralisation.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre la polyligne et les éléments en relation.</param>
    '''<param name="dDistLat"> Distance latérale utilisée pour éliminer des sommets en trop.</param>
    '''<param name="dLargGenMin"> Largeur minimum utilisée pour généraliser.</param>
    '''<param name="dLongGenMin"> Longueur minimale utilisée pour généraliser.</param>
    '''<param name="dLongMin"> Longueur minimale d'une ligne.</param>
    '''<param name="pPolylineGen"> Interface contenant les lignes de généralisation.</param>
    '''<param name="pPolylineErr"> Interface contenant les lignes de généralisation en erreur.</param>
    '''<param name="pSquelette"> Interface contenant le squelette de la polyligne.</param>
    '''<param name="pSqueletteEnv"> Interface contenant le squelette de la polyligne avec son enveloppe.</param>
    '''<param name="pBagDroites"> Interface contenant les droites des triangles de Delaunay.</param>
    '''<param name="pBagDroitesEnv"> Interface contenant les droites des triangles de Delaunay avec son enveloppe.</param>
    ''' 
    Public Shared Sub GeneraliserLigne(ByVal pPolyline As IPolyline, ByVal pPointsConnexion As IMultipoint, ByVal dDistLat As Double,
                                       ByVal dLargGenMin As Double, ByVal dLongGenMin As Double, ByVal dLongMin As Double, _
                                       ByRef pPolylineGen As IPolyline, ByRef pPolylineErr As IPolyline, ByRef pSquelette As IPolyline,
                                       ByRef pSqueletteEnv As IPolyline, ByRef pBagDroites As IGeometryBag, ByRef pBagDroitesEnv As IGeometryBag)
        'Déclarer les variables de travail
        Dim pPath As IPath = Nothing                            'Interface contenant une ligne à généraliser.
        Dim pPolylineTmp As IPolyline = Nothing                 'Interface contenant une ligne temporaire.
        Dim pPolylineErrTmp As IPolyline = Nothing              'Interface contenant une ligne d'erreur de généralisation temporaire.
        Dim pPolylineGenTmp As IPolyline = Nothing              'Interface contenant une ligne généralisée temporaire.
        Dim pBagDroitesTmp As IGeometryBag = Nothing            'Interface contenant les droites des triangles de Delaunay temporaire.
        Dim pBagDroitesEnvTmp As IGeometryBag = Nothing         'Interface contenant les droites des triangles de Delaunay avec son enveloppe temporaire.
        Dim pDictLiens As New Dictionary(Of Integer, Noeud)     'Dictionnaire contenant l'information des sommets de la polyligne à traiter.
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionnaire contenant les numéros des droites à utiliser pour effectuer la généralisation.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les composantes d'une polyligne.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface pour ajouter les composantes d'une polyligne.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour soustraire une géométrie.

        Try
            'Enlever les sommets en trop
            'pPolyline.Generalize(dDistLat)

            'Densifer les sommets de la polyligne
            pPolyline.Densify(dLargGenMin / 2, 0)

            'Créer la ligne vide
            pPolylineErr = New Polyline
            pPolylineErr.SpatialReference = pPolyline.SpatialReference
            'Créer la ligne vide
            pPolylineGen = New Polyline
            pPolylineGen.SpatialReference = pPolyline.SpatialReference
            'Créer le Bag vide
            pBagDroites = New GeometryBag
            pBagDroites.SpatialReference = pPolyline.SpatialReference
            'Créer le Bag vide
            pBagDroitesEnv = New GeometryBag
            pBagDroitesEnv.SpatialReference = pPolyline.SpatialReference

            'Interface pour extraire les composantes d'une polyligne.
            pGeomColl = CType(pPolyline, IGeometryCollection)

            'Traiter toutes les composantes de la ligne
            For i = 0 To pGeomColl.GeometryCount - 1
                'Définir la ligne à généraliser
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Vérifier si la ligne est plus grande que la longueur minimum de généralisation
                If pPath.Length > dLongGenMin Then
                    'Créer la ligne vide
                    pPolylineTmp = New Polyline
                    pPolylineTmp.SpatialReference = pPolyline.SpatialReference
                    'Interface pour ajouter les composantes d'une polyligne
                    pGeomCollAdd = CType(pPolylineTmp, IGeometryCollection)
                    'Ajouter la ligne à la polyligne temporaire
                    pGeomCollAdd.AddGeometry(pPath)

                    'Créer le squelette de la ligne selon la triangulation de Delaunay
                    Call CreerSqueletteLigneDelaunay(pPath, dLargGenMin, dLongGenMin, pSquelette, pDictLiens,
                                                     pSqueletteEnv, pBagDroitesTmp, pBagDroitesEnvTmp)

                    'Créer la polyligne d'erreur de généralisation
                    Call CreerPolyligneErreurGeneralisation(pPolylineTmp, pSquelette, pBagDroitesTmp, dLargGenMin, dLongGenMin, pPolylineErrTmp)

                    'Créer la polyligne de généralisation
                    Call CreerPolyligneGeneralisation(pPolylineTmp, pDictLiens, pBagDroitesEnvTmp, pPolylineErrTmp, pPolylineGenTmp)

                    'Interface pour ajouter les composantes d'une polyligne
                    pGeomCollAdd = CType(pPolylineErr, IGeometryCollection)
                    'Ajouter la ligne à la polyligne d'erreur de généralisation
                    pGeomCollAdd.AddGeometryCollection(CType(pPolylineErrTmp, IGeometryCollection))

                    'Interface pour ajouter les composantes d'une polyligne
                    pGeomCollAdd = CType(pPolylineGen, IGeometryCollection)
                    'Ajouter la ligne à la polyligne de généralisation
                    pGeomCollAdd.AddGeometryCollection(CType(pPolylineGenTmp, IGeometryCollection))

                    'Interface pour ajouter les composantes des droites
                    pGeomCollAdd = CType(pBagDroites, IGeometryCollection)
                    'Ajouter les droites
                    pGeomCollAdd.AddGeometryCollection(CType(pBagDroitesTmp, IGeometryCollection))

                    'Interface pour ajouter les composantes des droites
                    pGeomCollAdd = CType(pBagDroitesEnv, IGeometryCollection)
                    'Ajouter les droites
                    pGeomCollAdd.AddGeometryCollection(CType(pBagDroitesEnvTmp, IGeometryCollection))

                    'Si la longueur est plus petite
                Else
                    'Interface pour ajouter la ligne à la polyligne de généralisation
                    pGeomCollAdd = CType(pPolylineGen, IGeometryCollection)
                    'Ajouter la ligne à la polyligne de généralisation
                    pGeomCollAdd.AddGeometry(pPath)
                End If
            Next

            'Détruire les lignes non-connectées qui sont inférieures à la longueur minimale d'une ligne
            Call TraiterLongueurLigneMinimale(pPolylineGen, pPointsConnexion, dLongMin)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPath = Nothing
            pPolylineTmp = Nothing
            pPolylineErrTmp = Nothing
            pPolylineGenTmp = Nothing
            pBagDroitesTmp = Nothing
            pBagDroitesEnvTmp = Nothing
            pDictLiens = Nothing
            pDictDroites = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions privées"
    ''' <summary>
    ''' Routine qui permet de détruire les lignes non-connectées qui sont inférieures à la longueur minimale d'une ligne
    ''' </summary>
    ''' 
    '''<param name="pPolyline">Interface contenant la polyligne à traiter.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion.</param>
    '''<param name="dLongMin">Contient la longueur minimale d'une ligne.</param>
    ''' 
    Private Shared Sub TraiterLongueurLigneMinimale(ByRef pPolyline As IPolyline, ByVal pPointsConnexion As IMultipoint, ByVal dLongMin As Double)
        'Déclarer les variables de travail
        Dim pExtremite As IMultipoint = Nothing             'Interface contenant les extrémités de la polyligne.
        Dim pLimite As IMultipoint = Nothing                'Interface contenant les limites de la polyligne.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter des points.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire des lignes.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier la connexion.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour corriger la topologie.
        Dim bTraiter As Boolean = True                      'Indiquer si on doit traiter.

        Try
            'Traiter tant qu'il y a une correction
            Do While bTraiter
                'Interface pour extraire les extrémités des lignes
                pTopoOp = CType(pPolyline, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()

                'Extraire les limites de la polyligne
                pLimite = IdentifierLimiteSquelette(pPolyline)

                'Extraire les extrémités des lignes
                pTopoOp = CType(pTopoOp.Boundary, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()

                'Enlever les limites des extrémités afin de conserver que les extrémités intérieures
                pLimite = CType(pTopoOp.Difference(pLimite), IMultipoint)
                'Interface pour ajouter des points
                pPointColl = CType(pLimite, IPointCollection)
                'Ajouter des points
                pPointColl.AddPointCollection(CType(pPointsConnexion, IPointCollection))
                'Interface pour simplifier les points de connexion
                pTopoOp = CType(pLimite, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()

                'Interface pour vérifier la connexion
                pRelOp = CType(pLimite, IRelationalOperator)

                'Interface pour extraire les lignes
                pGeomColl = CType(pPolyline, IGeometryCollection)

                'Indiquer qu'on ne doit plus traiter
                bTraiter = False

                'Traiter toutes les lignes
                For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                    'Interface contenant une ligne
                    pPath = CType(pGeomColl.Geometry(i), IPath)

                    'Vérifier si la longueur de la ligne est inférieure à la longueur minimale
                    If pPath.Length < dLongMin Then
                        'Vérifier si la ligne est fermée ou si le point de début ou de fin ne sont pas connecté
                        If pPath.IsClosed Or pRelOp.Disjoint(pPath.FromPoint) Or pRelOp.Disjoint(pPath.ToPoint) Then
                            'Détruire la ligne
                            pGeomColl.RemoveGeometries(i, 1)

                            'Indiquer qu'on doit traiter de nouveau
                            bTraiter = True
                        End If
                    End If
                Next
            Loop

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pExtremite = Nothing
            pLimite = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet d'ajouter les sommets de connexion des lignes des triangles de Delaunay au polygone.
    ''' </summary>
    ''' 
    '''<param name="pLignesInt">Interface contenant les lignes intérieures des triangles de Delaunay.</param>
    '''<param name="pPolygon">Polygone utilisée pour effectuer la généralisation d'un anneau intérieur.</param>
    ''' 
    Private Shared Sub ConnecterLignesPolygone(ByVal pLignesInt As IPolyline, ByRef pPolygon As IPolygon)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface contenant les lignes intérieures des triangles de Delaunay.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets de connexion au polygone.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne des triangles de Delaunay.
        Dim pRing As IRing = Nothing                        'Interface contenant un anneau du polygone.
        Dim pHitTest As IHitTest = Nothing                  'Interface pour tester la présence du sommet recherché
        Dim pRingColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes du polygone.
        Dim pNewPoint As IPoint = Nothing                   'Interface contenant le nouveau point trouvé (pas utilisé).
        Dim dDistance As Double = Nothing                   'Interface contenant la distance calculée entre le point de recherche et le sommet trouvé.
        Dim nNumeroPartie As Integer = Nothing              'Numéro de partie trouvée.
        Dim nNumeroSommet As Integer = Nothing              'Numéro de sommet de la partie trouvée.
        Dim bCoteDroit As Boolean = Nothing                 'Indiquer si le point trouvé est du côté droit de la géométrie.

        Try
            'Interface pour vérifier la connexion
            pHitTest = CType(pPolygon, IHitTest)

            'Interface pour extraire les anneaux du polygone
            pRingColl = CType(pPolygon, IGeometryCollection)

            'Interface pour extraire les lignes des triangles
            pGeomColl = CType(pLignesInt, IGeometryCollection)

            'Traiter toutes les lignes du triangle
            For i = 0 To pGeomColl.GeometryCount - 1
                'Interface contenant une ligne de triangle
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Vérifier si le premier point de la ligne n'intersecte pas avec un sommet du polygone
                If Not pHitTest.HitTest(pPath.FromPoint, 0.002, esriGeometryHitPartType.esriGeometryPartVertex, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Vérifier si le premier point de la ligne intersecte avec la limite du polygone
                    If pHitTest.HitTest(pPath.FromPoint, 0.002, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                        'Interface pour extraire le sommet de la composante du polygone
                        pPointColl = CType(pRingColl.Geometry(nNumeroPartie), IPointCollection)

                        'Insérer un nouveau sommet
                        pPointColl.InsertPoints(nNumeroSommet + 1, 1, pPath.FromPoint)
                    End If
                End If

                'Vérifier si le dernier point de la ligne n'intersecte pas avec un sommet du polygone
                If Not pHitTest.HitTest(pPath.ToPoint, 0.002, esriGeometryHitPartType.esriGeometryPartVertex, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Vérifier si le dernier point de la ligne intersecte avec la limite du polygone
                    If pHitTest.HitTest(pPath.ToPoint, 0.002, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                        'Interface pour extraire le sommet de la composante du polygone
                        pPointColl = CType(pRingColl.Geometry(nNumeroPartie), IPointCollection)

                        'Insérer un nouveau sommet
                        pPointColl.InsertPoints(nNumeroSommet + 1, 1, pPath.ToPoint)
                    End If
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPointColl = Nothing
            pPath = Nothing
            pRing = Nothing
            pHitTest = Nothing
            pRingColl = Nothing
            pNewPoint = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de retourner le Polygone de généralisation intérieure et la Polyline d'erreur généralisation d'un anneau.
    ''' </summary>
    ''' 
    '''<param name="pPolygon">Polygone utilisée pour effectuer la généralisation d'un anneau intérieur.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre le polygon et les éléments en relation.</param>
    '''<param name="dLargMin"> Largeur minimum utilisée pour généraliser.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour généraliser.</param>
    '''<param name="dSupMin"> Contient la superficie de généralisation minimum.</param>
    '''<param name="pPolygonGen"> Interface contenant le polygone de généralisation.</param>
    '''<param name="pPolylineErr"> Interface contenant la polyligne d'erreur de généralisation.</param>
    '''<param name="pSquelette"> Interface contenant le squelette de la polyligne.</param>
    '''<param name="pBagDroites"> Interface contenant les droites des triangles de Delaunay utilisées pour effectuer la généralisation.</param>
    ''' 
    Private Shared Sub GeneraliserAnneauInterieur(ByVal pPolygon As IPolygon, ByVal pPointsConnexion As IMultipoint,
                                                  ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                                  ByRef pPolygonGen As IPolygon, ByRef pPolylineErr As IPolyline,
                                                  ByRef pSquelette As IPolyline, ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pLignesInt As IPolyline = Nothing                   'Interface contenant les lignes intérieures des triangles de Delaunay.
        Dim pPointsNonConnexion As IMultipoint = Nothing        'Interface contenant les points du squelette qui doivent être connectés.
        Dim pBagSommetsPolygone As IGeometryBag = Nothing       'Interface contenant les sommets du polygone à traiter.
        Dim pBagLignesPrimaire As IGeometryBag = Nothing        'Interface contenant les lignes du squelette primaire.
        Dim pBagLignesBase As IGeometryBag = Nothing            'Interface contenant les lignes du squelette de base.
        Dim pBagLignesSimple As IGeometryBag = Nothing          'Interface contenant les lignes simples du squelette primaire.
        Dim pDictLiens As New Dictionary(Of Integer, Noeud)     'Dictionnaire contenant l'information des liens entre les sommets du polygone à traiter.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour effectuer des opérations spatiales.

        Try
            'Créer les lignes des triangles intérieures
            pLignesInt = CreerPolyligneTrianglesDelaunay(pPolygon)

            'Connecter les lignes des triangles intérieures au polygone
            Call ConnecterLignesPolygone(pLignesInt, pPolygon)

            'Créer le squelette de base du polygone selon les lignes des triangles de Delaunay
            Call CreerSqueletteBaseDelaunay(pPolygon, pLignesInt, pSquelette, pBagDroites)

            'Connecter le squelette aux points de connexion des éléments en relation
            'Call ConnecterSquelettePointsConnexion(pPointsConnexion, dLargMin, dLongMin, pSquelette)
            'Call ConnecterSquelettePointsConnexion(pPointsConnexion, pSquelette)
            Call ConnecterSquelettePointsConnexion(pPolygon, pPointsConnexion, pSquelette)

            'Enlever les extrémités de lignes superflux dans le squelette et obtenir les points du squelette non connectés
            pPointsNonConnexion = EnleverExtremiteLigne(pPointsConnexion, pSquelette, dLongMin)

            'Connecter les lignes non connectées du squelette avec la limite du polygone
            Call ConnecterSquelettePolygone(pPolygon, pPointsNonConnexion, pSquelette)

            'Créer la polyligne d'erreur de généralisation
            Call CreerPolyligneErreurGeneralisation(pSquelette, pBagDroites, dLargMin, dLongMin, pPolylineErr)

            'Créer le polyone de généralisation
            Call CreerPolygoneGeneralisation(pPolygon, pBagDroites, dLargMin, dLongMin, dSupMin, pPolylineErr, pPolygonGen)

            'Interface pour extraire la partie du squelette à l'extérieure du polygone de généralisation
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            'Extraire la partie du squelette à l'extérieure du polygone de généralisation
            pPolylineErr = CType(pTopoOp.Difference(pPolygonGen), IPolyline)

            'Créer le polyone de généralisation
            Call CreerPolygoneGeneralisation(pPolygon, pBagDroites, dLargMin, dLongMin, dSupMin, pPolylineErr, pPolygonGen)

            'Extraire la partie du squelette à l'intérieure du polygone de généralisation
            pSquelette = CType(pTopoOp.Intersect(pPolygonGen, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLignesInt = Nothing
            pPointsNonConnexion = Nothing
            pBagSommetsPolygone = Nothing
            pBagLignesPrimaire = Nothing
            pBagLignesBase = Nothing
            pBagLignesSimple = Nothing
            pDictLiens = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de retourner le Polygone de généralisation extérieure et la Polyline d'erreur de généralisation d'un anneau extérieur.
    ''' </summary>
    ''' 
    '''<param name="pPolygon">Polygone utilisée pour effectuer la généralisation d'un polygone extérieur.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre le polygon et les éléments en relation.</param>
    '''<param name="dLargMin"> Largeur minimum utilisée pour généraliser.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour généraliser.</param>
    '''<param name="dSupMin"> Contient la superficie de généralisation minimum.</param>
    '''<param name="pPolygonGen"> Interface contenant le polygone de généralisation extérieur.</param>
    '''<param name="pPolylineErr"> Interface contenant la polyligne d'erreur de généralisation.</param>
    '''<param name="pSquelette"> Interface contenant le squelette de la polyligne.</param>
    '''<param name="pBagDroites"> Interface contenant les droites des triangles de Delaunay utilisées pour effectuer la généralisation.</param>
    ''' 
    Private Shared Sub GeneraliserAnneauExterieur(ByVal pPolygon As IPolygon, ByVal pPointsConnexion As IMultipoint,
                                                  ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                                  ByRef pPolygonGen As IPolygon, ByRef pPolylineErr As IPolyline,
                                                  ByRef pSquelette As IPolyline, ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pLignesExt As IPolyline = Nothing                   'Interface contenant les lignes extérieures des triangles de Delaunay.
        Dim pPointsNonConnexion As IMultipoint = Nothing        'Interface contenant les points du squelette qui doivent être connectés.
        Dim pBagSommetsPolygone As IGeometryBag = Nothing       'Interface contenant les sommets du polygone à traiter.
        Dim pBagLignesPrimaire As IGeometryBag = Nothing        'Interface contenant les lignes du squelette primaire.
        Dim pBagLignesBase As IGeometryBag = Nothing            'Interface contenant les lignes du squelette de base.
        Dim pBagLignesSimple As IGeometryBag = Nothing          'Interface contenant les lignes simples du squelette primaire.
        Dim pDictLiens As New Dictionary(Of Integer, Noeud)     'Dictionnaire contenant l'information des liens entre les sommets du polygone à traiter.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour effectuer des opérations spatiales.

        Try
            'Créer les lignes des triangles extérieures
            pLignesExt = CreerPolyligneTrianglesDelaunay(pPolygon)

            'Connecter les lignes des triangles intérieures au polygone
            Call ConnecterLignesPolygone(pLignesExt, pPolygon)

            'Créer le squelette de base du polygone selon les lignes des triangles de Delaunay
            Call CreerSqueletteBaseDelaunay(pPolygon, pLignesExt, pSquelette, pBagDroites)

            'Aucun point de connexion des éléments en relations
            pPointsConnexion = New Multipoint
            pPointsConnexion.SpatialReference = pPolygon.SpatialReference

            'Connecter le squelette aux points de connexion des éléments en relation
            'Call ConnecterSquelettePointsConnexion(pPointsConnexion, dLargMin, dLongMin, pSquelette)
            'Call ConnecterSquelettePointsConnexion(pPointsConnexion, pSquelette)
            Call ConnecterSquelettePointsConnexion(pPolygon, pPointsConnexion, pSquelette)

            'Enlever les extrémités de lignes superflux dans le squelette et retourner les points du squelette non-connectés
            pPointsNonConnexion = EnleverExtremiteLigne(pPointsConnexion, pSquelette, dLongMin)

            'Connecter les lignes non connectées du squelette avec la limite du polygone
            Call ConnecterSquelettePolygone(pPolygon, pPointsNonConnexion, pSquelette)

            'Créer la polyligne d'erreur de généralisation
            Call CreerPolyligneErreurGeneralisation(pSquelette, pBagDroites, dLargMin, dLongMin, pPolylineErr)

            'Créer le polyone de généralisation
            Call CreerPolygoneGeneralisation(pPolygon, pBagDroites, dLargMin, dLongMin, dSupMin, pPolylineErr, pPolygonGen)

            'Interface pour extraire la partie du squelette à l'extérieure du polygone de généralisation
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            'Extraire la partie du squelette à l'extérieure du polygone de généralisation
            pPolylineErr = CType(pTopoOp.Difference(pPolygonGen), IPolyline)

            'Créer le polyone de généralisation
            Call CreerPolygoneGeneralisation(pPolygon, pBagDroites, dLargMin, dLongMin, dSupMin, pPolylineErr, pPolygonGen)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLignesExt = Nothing
            pPointsNonConnexion = Nothing
            pBagSommetsPolygone = Nothing
            pBagLignesPrimaire = Nothing
            pBagLignesBase = Nothing
            pBagLignesSimple = Nothing
            pDictLiens = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Function qui permet de créer la polyligne d'erreur de généralisation d'une géométrie.
    ''' </summary>
    ''' 
    ''' <param name="pSquelette">Interface contenant le squelette de la géométrie.</param>
    ''' <param name="pBagLignesTriangles">Interface contenant les lignes des triangles.</param>
    ''' <param name="dLargMin">Contient la largeur de généralisation minimum.</param>
    ''' <param name="dLongMin">Contient la longueur de généralisation minimum.</param>
    ''' <param name="pPolylineErr">Interface contenant la polyligne d'erreur de généralisation.</param>
    ''' 
    Private Shared Sub CreerPolyligneErreurGeneralisation(ByVal pSquelette As IPolyline, ByVal pBagLignesTriangles As IGeometryBag, _
                                                          ByVal dLargMin As Double, ByVal dLongMin As Double, ByRef pPolylineErr As IPolyline)
        'Déclarer les variables de travail
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionnaire utilisé pour indiquer les droites à traiter.
        Dim pMultiPoint As IMultipoint = Nothing            'Interface contenant les points des lignes de généralisation.
        Dim pBagPoints As IGeometryBag = Nothing            'Interface contenant les points des lignes de généralisation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour simplifier.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter une relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement d'une relation spatiale.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface utilisé pour extraire les composantes d'une géométrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface utilisé pour ajouter les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les sommets d'une composante.
        Dim pPointCollAdd As IPointCollection = Nothing     'Interface utilisé pour ajouter les sommets d'une composante.
        Dim pDroiteColl As IGeometryCollection = Nothing    'Interface pour extraire les composantes des lignes des Triangles.
        Dim pDroite As IPolyline = Nothing                  'Interface contenant une droite des lignes des Triangles.
        Dim pPoint As IPoint = Nothing                      'Interface contenant un sommet d'une composante.
        Dim pPath As IPath = Nothing                        'Interface contenant une composante de type ligne.
        Dim dLargeur As Double = 0                          'Contient la largeur du polygon à un point du squelette.
        Dim iSel As Integer = -1                            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                            'Numéro de séquence de la géométrie en relation.
        Dim iNo As Integer = 0      'Contient le numéro de la droite.
        Dim k As Integer = 0        'Compteur de sommets.

        Try
            'Définir la valeur par défaut
            pPolylineErr = New Polyline
            pPolylineErr.SpatialReference = pSquelette.SpatialReference

            'Créer un nouveau Bag vide des points de généralisation
            pBagPoints = New GeometryBag
            pBagPoints.SpatialReference = pPolylineErr.SpatialReference
            'Interface pour ajouter les points dans le Bag
            pGeomColl = CType(pBagPoints, IGeometryCollection)
            'Interface pour extraire les sommets du squelette
            pPointColl = CType(pSquelette, IPointCollection)
            'Traiter tous les sommets du squelette
            For i = 0 To pPointColl.PointCount - 1
                'Ajouter le point dans le Bag
                pGeomColl.AddGeometry(pPointColl.Point(i))
            Next

            'Interface pour ajouter les lignes nulles dans le Bag des lignes des triangles
            pGeomColl = CType(pBagLignesTriangles, IGeometryCollection)

            'Interface pour traiter les relations spatiales entre les sommets du polygone et les lignes des triangles
            pRelOpNxM = CType(pBagPoints, IRelationalOperatorNxM)

            'Traiter la relation spatiale entre les sommets du polygone et les lignes des triangles
            pRelResult = pRelOpNxM.Within(CType(pBagLignesTriangles, IGeometryBag))

            '---------------------------------------------------
            'Créer le dictionnaire des droites
            '---------------------------------------------------
            'Trier le résultat selon les sommets du squelette
            pRelResult.SortLeft()

            'Traiter toutes les relations Lignes-Sommets
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si la ligne existe
                If Not pDictDroites.ContainsKey(iSel) Then
                    'Ajouter la ligne dans le dictionnaire
                    pDictDroites.Add(iSel, iRel)
                End If
            Next

            'Interface pour extraire les composantes des Lignes des Triangles
            pDroiteColl = CType(pBagLignesTriangles, IGeometryCollection)

            'Interface pour extraire les composantes du squelette
            pGeomColl = CType(pSquelette, IGeometryCollection)

            'Interface pour ajouter les composantes des lignes de généralisation
            pGeomCollAdd = CType(pPolylineErr, IGeometryCollection)

            'Traiter toutes les composantes du squelette
            For i = 0 To pGeomColl.GeometryCount - 1
                'Créer une nouvelle ligne de généralisation vide
                pPath = New Path
                pPath.SpatialReference = pPolylineErr.SpatialReference
                'Interface pour ajouter les sommets de la ligne de généralisation
                pPointCollAdd = CType(pPath, IPointCollection)

                'Interface pour extraire les sommets de la composante du squelette
                pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)

                'Traiter tous les sommets de la composante du squelette
                For j = 0 To pPointColl.PointCount - 1
                    'Interface contenant le point du squelette
                    pPoint = pPointColl.Point(j)

                    'Vérifier si la droite est présente dans le dictionnaire
                    If pDictDroites.ContainsKey(k) Then
                        'Définir le noméro de la droite
                        iNo = pDictDroites.Item(k)

                        'Interface contenant la droite
                        pDroite = CType(pDroiteColl.Geometry(iNo), IPolyline)

                        'Vérifier si la largeur est inférieure ou égale la la largeur minimum
                        If pDroite.Length <= dLargMin Then
                            'Ajouter le sommet du squelette à la ligne de généralisation
                            pPointCollAdd.AddPoint(pPoint)

                            'Si la largeur est supérieure la la largeur minimum
                        Else
                            'Vérifier si la longueur de la ligne de généralisation est supérieures ou égale à la longueur minimum
                            If pPath.Length >= dLongMin Then
                                'Ajouter la ligne de généralisation
                                pGeomCollAdd.AddGeometry(pPath)
                            End If

                            'Créer une nouvelle ligne de généralisation vide
                            pPath = New Path
                            pPath.SpatialReference = pPolylineErr.SpatialReference
                            'Interface pour ajouter les sommets de la ligne de généralisation
                            pPointCollAdd = CType(pPath, IPointCollection)
                        End If
                    Else
                        'Debug.Print("Erreur " & k.ToString)
                        'Ajouter le sommet du squelette à la ligne de généralisation
                        pPointCollAdd.AddPoint(pPoint)
                    End If

                    'Compter les sommets
                    k = k + 1

                    'Debug.Print((i + 1).ToString & "." & (j + 1).ToString & " = " & dLargeur.ToString)
                Next

                'Vérifier si la longueur de la ligne de généralisation est supérieure ou égale à la longueur minimum
                If pPath.Length >= dLongMin Then
                    'Ajouter la ligne de généralisation
                    pGeomCollAdd.AddGeometry(pPath)
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pDictDroites = Nothing
            pMultiPoint = Nothing
            pBagPoints = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pPoint = Nothing
            pPath = Nothing
            pDroiteColl = Nothing
            pDroite = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Function qui permet de créer la polyligne d'erreur de généralisation d'une géométrie.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Interface contenant la polyligne à généraliser.</param>
    ''' <param name="pSquelette">Interface contenant le squelette de la géométrie.</param>
    ''' <param name="pBagLignesTriangles">Interface contenant les lignes des triangles.</param>
    ''' <param name="dLargMin">Contient la largeur de généralisation minimum.</param>
    ''' <param name="dLongMin">Contient la longueur de généralisation minimum.</param>
    ''' <param name="pPolylineErr">Interface contenant la polyligne d'erreur de généralisation.</param>
    ''' 
    Private Shared Sub CreerPolyligneErreurGeneralisation(ByVal pPolyline As IPolyline, ByVal pSquelette As IPolyline, ByVal pBagLignesTriangles As IGeometryBag, _
                                                          ByVal dLargMin As Double, ByVal dLongMin As Double, ByRef pPolylineErr As IPolyline)
        'Déclarer les variables de travail
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionnaire utilisé pour indiquer les droites à traiter.
        Dim pSqueletteFil As IPolyline = Nothing           'Interface contenant le squelette filtré.
        Dim pMultiPoint As IMultipoint = Nothing            'Interface contenant les points des lignes de généralisation.
        Dim pBagPoints As IGeometryBag = Nothing            'Interface contenant les points des lignes de généralisation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour simplifier.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface qui permet de vérifier une relation spatiale.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter une relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement d'une relation spatiale.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface utilisé pour extraire les composantes d'une géométrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface utilisé pour ajouter les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les sommets d'une composante.
        Dim pPointCollAdd As IPointCollection = Nothing     'Interface utilisé pour ajouter les sommets d'une composante.
        Dim pDroiteColl As IGeometryCollection = Nothing    'Interface pour extraire les composantes des lignes des Triangles.
        Dim pDroite As IPolyline = Nothing                  'Interface contenant une droite des lignes des Triangles.
        Dim pPoint As IPoint = Nothing                      'Interface contenant un sommet d'une composante.
        Dim pPath As IPath = Nothing                        'Interface contenant une composante de type ligne.
        Dim pClone As IClone = Nothing                      'Interface utilisé pour cloner une géométrie.
        Dim dLargeur As Double = 0                          'Contient la largeur du polygon à un point du squelette.
        Dim iSel As Integer = -1                            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                            'Numéro de séquence de la géométrie en relation.
        Dim iNo As Integer = 0      'Contient le numéro de la droite.
        Dim k As Integer = 0        'Compteur de sommets.

        Try
            'Interface pour cloner une géométrie
            pClone = CType(pSquelette, IClone)
            'Définir le squelette filtré par défaut
            pSqueletteFil = CType(pClone.Clone, IPolyline)
            'Vérifier si la longueur minimale est supérieure à zéro
            If dLongMin > 0 Then
                'Définir le squelette filtré vide
                pSqueletteFil = New Polyline
                pSqueletteFil.SpatialReference = pSquelette.SpatialReference
                'Interface pour construire le squelette filtré des réseaux plus grands que la longueur minimale
                pTopoOp = CType(pSqueletteFil, ITopologicalOperator2)
                'Construire le squelette filtré à partir de toutes les parties des réseaux plus grands que la longueur minimale
                pTopoOp.ConstructUnion(CType(RegrouperReseauxPolyligne(pSquelette, dLongMin), IEnumGeometry))
            End If

            'Modifier la longueur minimale
            dLongMin = 0

            'Définir la valeur par défaut
            pPolylineErr = New Polyline
            pPolylineErr.SpatialReference = pSquelette.SpatialReference

            'Créer un nouveau Bag vide des points de généralisation
            pBagPoints = New GeometryBag
            pBagPoints.SpatialReference = pPolylineErr.SpatialReference
            'Interface pour ajouter les points dans le Bag
            pGeomColl = CType(pBagPoints, IGeometryCollection)
            'Interface pour extraire les sommets du squelette
            pPointColl = CType(pSqueletteFil, IPointCollection)
            'Traiter tous les sommets du squelette
            For i = 0 To pPointColl.PointCount - 1
                'Ajouter le point dans le Bag
                pGeomColl.AddGeometry(pPointColl.Point(i))
            Next

            'Interface pour ajouter les lignes nulles dans le Bag des lignes des triangles
            pGeomColl = CType(pBagLignesTriangles, IGeometryCollection)

            'Interface pour traiter les relations spatiales entre les sommets du polygone et les lignes des triangles
            pRelOpNxM = CType(pBagPoints, IRelationalOperatorNxM)

            'Traiter la relation spatiale entre les sommets du polygone et les lignes des triangles
            pRelResult = pRelOpNxM.Within(CType(pBagLignesTriangles, IGeometryBag))

            '---------------------------------------------------
            'Créer le dictionnaire des droites
            '---------------------------------------------------
            'Trier le résultat selon les sommets du squelette
            pRelResult.SortLeft()

            'Traiter toutes les relations Lignes-Sommets
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Vérifier si la ligne existe
                If Not pDictDroites.ContainsKey(iSel) Then
                    'Ajouter la ligne dans le dictionnaire
                    pDictDroites.Add(iSel, iRel)
                End If
            Next

            'Interface pour extraire les composantes des Lignes des Triangles
            pDroiteColl = CType(pBagLignesTriangles, IGeometryCollection)

            'Interface pour extraire les composantes du squelette
            pGeomColl = CType(pSqueletteFil, IGeometryCollection)

            'Interface pour ajouter les composantes des lignes de généralisation
            pGeomCollAdd = CType(pPolylineErr, IGeometryCollection)

            'Interface qui permet de vérifier une relation spatiale
            pRelOp = CType(pPolyline, IRelationalOperator)

            'Traiter toutes les composantes du squelette
            For i = 0 To pGeomColl.GeometryCount - 1
                'Créer une nouvelle ligne d'erreur de généralisation vide
                pPath = New Path
                pPath.SpatialReference = pPolylineErr.SpatialReference
                'Interface pour ajouter les sommets de la ligne de généralisation
                pPointCollAdd = CType(pPath, IPointCollection)

                'Interface pour extraire les sommets de la composante du squelette
                pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)

                'Traiter tous les sommets de la composante du squelette
                For j = 0 To pPointColl.PointCount - 1
                    'Interface contenant le point du squelette
                    pPoint = pPointColl.Point(j)

                    'Vérifier si la droite est présente dans le dictionnaire
                    If pDictDroites.ContainsKey(k) Then
                        'Définir le noméro de la droite
                        iNo = pDictDroites.Item(k)

                        'Interface contenant la droite
                        pDroite = CType(pDroiteColl.Geometry(iNo), IPolyline)

                        'Vérifier si la largeur est inférieure ou égale la la largeur minimum
                        If pDroite.Length <= dLargMin Then
                            'Ajouter le sommet du squelette à la ligne de généralisation
                            pPointCollAdd.AddPoint(pPoint)

                            'Si la largeur est supérieure la la largeur minimum
                        Else
                            'Vérifier que la ligne est supérieure à 0
                            If pPath.Length > 0 Then
                                'Vérifier si la longueur de la ligne de généralisation est supérieures ou égale à la longueur minimum
                                If pPath.Length >= dLongMin Or (pRelOp.Disjoint(pPath.FromPoint) And pRelOp.Disjoint(pPath.ToPoint)) Then
                                    'Ajouter la ligne d'erreur de généralisation
                                    pGeomCollAdd.AddGeometry(pPath)
                                End If
                            End If

                            'Créer une nouvelle ligne d'erreur de généralisation vide
                            pPath = New Path
                            pPath.SpatialReference = pPolylineErr.SpatialReference
                            'Interface pour ajouter les sommets de la ligne de généralisation
                            pPointCollAdd = CType(pPath, IPointCollection)
                        End If
                    Else
                        'Debug.Print("Erreur " & k.ToString)
                        'Ajouter le sommet du squelette à la ligne de généralisation
                        pPointCollAdd.AddPoint(pPoint)
                    End If

                    'Compter les sommets
                    k = k + 1

                    'Debug.Print((i + 1).ToString & "." & (j + 1).ToString & " = " & dLargeur.ToString)
                Next

                'Vérifier que la ligne est supérieure à 0
                If pPath.Length > 0 Then
                    'Vérifier si la longueur de la ligne de généralisation est supérieure ou égale à la longueur minimum
                    If pPath.Length >= dLongMin Or (pRelOp.Disjoint(pPath.FromPoint) And pRelOp.Disjoint(pPath.ToPoint)) Then
                        'Ajouter la ligne d'erreur de généralisation
                        pGeomCollAdd.AddGeometry(pPath)
                    End If
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pDictDroites = Nothing
            pSqueletteFil = Nothing
            pMultiPoint = Nothing
            pBagPoints = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pPoint = Nothing
            pPath = Nothing
            pDroiteColl = Nothing
            pDroite = Nothing
            pRelOp = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            pClone = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Function qui permet de regrouper dans un BAG les réseaux d'un ensemble de lignes qui sont plus grand qu'une longueur minimale.
    ''' </summary>
    ''' 
    ''' <remarks>Un réseau est un ensemble de lignes qui se touchent entre elles.</remarks>
    ''' 
    ''' <param name="pPolyline">Interface contenant un ou plusieurs réseaux linéaires.</param>
    ''' <param name="dLongMin">Contient la longueur de minimale d'un réseau.</param>
    ''' 
    ''' <returns>GeometryBag contenant tous les réseaux linéaires plus grand que la longueur minimale</returns>
    ''' 
    Private Shared Function RegrouperReseauxPolyligne(ByVal pPolyline As IPolyline, Optional ByVal dLongMin As Double = 0) As IGeometryBag
        'Déclarer les variables de travail
        Dim pBuffer As IPolygon4 = Nothing                  'Interface contenant le tampon de la polyligne.
        Dim pPolygon As IPolygon = Nothing                  'Interface contenant un polygone du tampon.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface qui permet d'ajouter des géométries.
        Dim pGeomCollRes As IGeometryCollection = Nothing   'Interface qui permet d'ajouter des réseaux.
        Dim pGeomCollExt As IGeometryCollection = Nothing   'Interface qui permet d'extraire les polygones extérieurs.
        Dim pReseau As IPolyline = Nothing                  'Interface contenant un réseau de lignes. 
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface qui permet d'effectuer des opérations spatiales.
        Dim pSpatialRefTol As ISpatialReferenceTolerance = Nothing  'Interface qui permet d'initialiser la tolérance XY.

        'Par défaut, le squellete résultant est un clone
        RegrouperReseauxPolyligne = New GeometryBag
        RegrouperReseauxPolyligne.SpatialReference = pPolyline.SpatialReference

        Try
            'Vérifier si la polyligne n'est pas vide
            If Not pPolyline.IsEmpty Then
                'Interface pour ajouter les réseaux
                pGeomCollRes = CType(RegrouperReseauxPolyligne, IGeometryCollection)

                'Interface pour extraire la tolérance
                pSpatialRefTol = CType(pPolyline.SpatialReference, ISpatialReferenceTolerance)

                'Interface pour créer le tampon de la polyligne
                pTopoOp = CType(pPolyline, ITopologicalOperator2)
                'Créer le Tampon de la polyligne
                pBuffer = CType(pTopoOp.Buffer(pSpatialRefTol.XYTolerance * 2), IPolygon4)

                'Vérifier si plusieurs surfaces sont présentes
                If pBuffer.ExteriorRingCount > 1 Then
                    'Définir le polygon du tampon
                    pGeomCollExt = CType(pBuffer.ExteriorRingBag(), IGeometryCollection)
                    'Traiter toutes les surfaces extérieurs
                    For i = 0 To pGeomCollExt.GeometryCount - 1
                        'Créer un polygone vide
                        pPolygon = New Polygon
                        pPolygon.SpatialReference = pPolyline.SpatialReference
                        'Définir le polygone extérieur
                        pRingExt = CType(pGeomCollExt.Geometry(i), IRing)
                        'Interface pour ajouter des géométries
                        pGeomColl = CType(pPolygon, IGeometryCollection)
                        'Ajouter les géométries du polygone
                        pGeomColl.AddGeometry(pRingExt)
                        pGeomColl.AddGeometryCollection(CType(pBuffer.InteriorRingBag(pRingExt), IGeometryCollection))
                        'Extraire les lignes qui se retrouve à l'intérieur du polygone
                        pReseau = CType(pTopoOp.Intersect(pPolygon, esriGeometryDimension.esriGeometry1Dimension), IPolyline)
                        'Vérifier si le réseau est plus grand que la longueur minimale
                        If pReseau.Length > dLongMin Then
                            'Ajouter le réseau dans le Bag des réseaux
                            pGeomCollRes.AddGeometry(pReseau)
                        End If
                    Next

                    'Si une seule surface, alors un seul réseau
                Else
                    'Vérifier si le réseau est plus grand que la longueur minimale
                    If pPolyline.Length > dLongMin Then
                        'Ajouter le réseau dans le Bag des réseaux
                        pGeomCollRes.AddGeometry(pPolyline)
                    End If
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pBuffer = Nothing
            pPolygon = Nothing
            pGeomCollRes = Nothing
            pGeomCollExt = Nothing
            pReseau = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet de créer et retourner la polyligne de généralisation.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Interface contenant la polyligne à généraliser.</param>
    ''' <param name="pDictLiens">Dictionnaire contenant l'information des sommets de la ligne à traiter.</param>
    ''' <param name="pBagDroites">Interface contenant les droites de Delaunay.</param>
    ''' <param name="pPolylineErr">Interface contenant la polyligne d'erreurs de généralisation.</param>
    ''' <param name="pPolylineGen">Interface contenant la polyligne de généralisation.</param>
    ''' 
    Private Shared Sub CreerPolyligneGeneralisation(ByVal pPolyline As IPolyline, ByVal pDictLiens As Dictionary(Of Integer, Noeud), ByVal pBagDroites As IGeometryBag, _
                                                    ByVal pPolylineErr As IPolyline, ByRef pPolylineGen As IPolyline)
        'Déclarer les variables de travail
        Dim pNoeudLien As Noeud = Nothing                       'Contient l'information d'un lien entre les sommets.
        Dim pPointColl As IPointCollection = Nothing            'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pPointCollAdd As IPointCollection = Nothing         'Interface utilisé pour ajouter les sommets d'une géométrie.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface qui permet d'extraire et ajouter des composantes de géométrie.
        Dim pPath As IPath = Nothing                            'Interface contenant une ligne. 
        Dim pRelOp As IRelationalOperator = Nothing             'Interface qui permet de vérifier une relation spatiale.
        Dim pTopoOP As ITopologicalOperator2 = Nothing          'Interface qui permet d'effectuer des opérations spatiales.
        Dim pPoint As IPoint = Nothing                          'Interface contenant le point du centre de la droite.
        Dim pPointA As IPoint = Nothing                         'Interface contenant le premier sommet d'une droite.
        Dim pPointB As IPoint = Nothing                         'Interface contenant le deuxième sommet d'une droite.
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionaire contenant les droites de Delaunay à utiliser.
        Dim pDroite As Droite = Nothing                         'Objet contenant la droite à traiter.

        Try
            'Vérifier s'il n'y a pas de la généralisation à effectuer
            If pPolylineErr.IsEmpty Then
                'Redéfinir la polyligne initiale
                pPolylineGen = pPolyline

                'S'il y a de la généralisation à effectuer
            Else
                'Créer le dictionnaire des droites de généralisation utilisé pour généraliser la polyline à traiter
                Call CreerDictDroiteGeneralisation(pPolylineErr, pBagDroites, pDictDroites)

                'Définir la valeur par défaut
                pPolylineGen = New Polyline
                pPolylineGen.SpatialReference = pPolyline.SpatialReference
                'Interface pour ajouter les lignes de généralisation
                pGeomColl = CType(pPolylineGen, IGeometryCollection)

                'Créer une nouvelle ligne vide
                pPath = New Path
                pPath.SpatialReference = pPolyline.SpatialReference
                'Interface pour ajouter les sommets de la ligne
                pPointCollAdd = CType(pPath, IPointCollection)

                'Interface pour extraire les sommets de la ligne
                pPointColl = CType(pPolyline, IPointCollection)

                'Traiter tous les sommets de la ligne
                For i = 0 To pPointColl.PointCount - 1
                    'Définir l'information du lien
                    pNoeudLien = pDictLiens.Item(i)

                    'Vérifier si le sommet possède des droites
                    If pNoeudLien.Droites.Count > 0 Then
                        'Traiter tous les sommets des liens
                        For j = 0 To pNoeudLien.Droites.Count - 1
                            'Définir le numéro de sommet en lien
                            pDroite = pNoeudLien.Droites.Item(j)

                            'Vérifier si la droite doit être traité
                            If pDictDroites.ContainsKey(pDroite.No) Then
                                'Vérifier si on traite le premier sommet et le premier lien
                                If i = 0 And j = 0 Then
                                    'Ajouter le sommet tel quel dans la ligne pour conserver la connexion
                                    pPointCollAdd.AddPoint(pPointColl.Point(i))
                                End If

                                'Définir le point de début de la droite
                                pPointA = pDroite.PointDeb

                                'Définir le point de fin de la droite
                                pPointB = pDroite.PointFin

                                'Créer le point vide
                                pPoint = New Point
                                pPoint.SpatialReference = pPolyline.SpatialReference
                                'Définir le centre de la droite
                                pPoint.X = (pPointA.X + pPointB.X) / 2
                                pPoint.Y = (pPointA.Y + pPointB.Y) / 2
                                'Ajouter le sommet tel quel dans la ligne
                                pPointCollAdd.AddPoint(pPoint)

                                'Vérifier si on traite le dernier sommet et le dernier lien
                                If i = pPointColl.PointCount - 1 And j = pNoeudLien.Droites.Count - 1 Then
                                    'Ajouter le sommet tel quel dans la ligne pour conserver la connexion
                                    pPointCollAdd.AddPoint(pPointColl.Point(i))
                                End If

                            Else
                                'Ajouter le sommet tel quel dans la ligne
                                pPointCollAdd.AddPoint(pPointColl.Point(i))
                            End If
                        Next

                        'Si le sommet ne possède pas de lien avec un autre sommet
                    Else
                        'Ajouter le sommet tel quel dans la ligne
                        pPointCollAdd.AddPoint(pPointColl.Point(i))
                    End If
                Next

                'Ajouter la lignes la polyligne de généralisation
                pGeomColl.AddGeometry(pPath)

                'Simplifier la polyligne de généralisation
                pTopoOP = CType(pPolylineGen, ITopologicalOperator2)
                pTopoOP.IsKnownSimple_2 = False
                pTopoOP.Simplify()

                'Éliminer les triangles de la polyligne de généralisation.
                'Call EliminerTrianglePolyligne(pPolyline, pPolylineErr, pPolylineGen)
                'Éliminer les triangles de la polyligne de généralisation.
                Call EliminerTrianglePolyligne(pPolyline, pPolylineGen)
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pNoeudLien = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pTopoOP = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pPoint = Nothing
            pPointA = Nothing
            pPointB = Nothing
            pDictDroites = Nothing
            pDroite = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet d'éliminer les triangles de la polyligne de généralisation.
    ''' La partie la plus longue de chaque triangle est éliminée.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Interface contenant la polyligne à généraliser.</param>
    ''' <param name="pPolylineErr">Interface contenant la polyligne d'erreurs de généralisation.</param>
    ''' <param name="pPolylineGen">Interface contenant la polyligne de généralisation.</param>
    ''' 
    Private Shared Sub EliminerTrianglePolyligne(ByVal pPolyline As IPolyline, ByVal pPolylineErr As IPolyline, ByRef pPolylineGen As IPolyline)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface qui permet d'extraire  des composantes de géométrie.
        Dim pPath As IPath = Nothing                            'Interface contenant une ligne. 
        Dim pRelOp As IRelationalOperator = Nothing             'Interface qui permet de vérifier une relation spatiale.
        Dim pTopoOP As ITopologicalOperator2 = Nothing          'Interface qui permet d'effectuer des opérations spatiales.

        Try
            'Interface pour effectuer des opérations spatiales.
            pTopoOP = CType(pPolylineGen, ITopologicalOperator2)
            'Enlever la partie des lignes en erreur
            pPolylineGen = CType(pTopoOP.Difference(pPolylineErr), IPolyline)

            'Interface pour vérifier les relations spatiales
            pRelOp = CType(pPolyline, IRelationalOperator)
            'Interface pour extraire et ajouter des lignes dans la ligne généralisée
            pGeomColl = CType(pPolylineGen, IGeometryCollection)
            'Traiter tous les composantes de la ligne généralisée
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Interface pour extraire la longueur de la ligne
                pPath = CType(pGeomColl.Geometry(i), IPath)
                'Vérifier si la ligne est disjoint de la ligne initiale - (Ligne à détruire)
                If pRelOp.Disjoint(pPath.FromPoint) Then
                    'Retirer la ligne à détruire de la ligne généralisée
                    pGeomColl.RemoveGeometries(i, 1)
                End If
            Next
            'Ajouter les lignes en erreur dans la ligne généralisée
            pGeomColl.AddGeometryCollection(CType(pPolylineErr, IGeometryCollection))
            'Simplifier la géométrie
            pTopoOP = CType(pPolylineGen, ITopologicalOperator2)
            pTopoOP.IsKnownSimple_2 = False
            pTopoOP.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pTopoOP = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de'éliminer les triangles de la polyligne de généralisation.
    ''' La partie la plus longue de chaque triangle est éliminée.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Interface contenant la polyligne à généraliser.</param>
    ''' <param name="pPolylineGen">Interface contenant la polyligne de généralisation.</param>
    ''' 
    Private Shared Sub EliminerTrianglePolyligne(ByVal pPolyline As IPolyline, ByRef pPolylineGen As IPolyline)
        'Déclarer les variables de travail
        Dim pTriangles As IPolyline = Nothing                   'Interface contenant les triangles.
        Dim pLignes As IPolyline = Nothing                      'Interface contenant les lignes à détruire.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface qui permet d'extraire  des composantes de géométrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface qui permet d'ajouter des composantes de géométrie.
        Dim pPointColl As IPointCollection = Nothing            'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pPointCollAdd As IPointCollection = Nothing         'Interface utilisé pour ajouter les sommets d'une géométrie.
        Dim pSegColl As ISegmentCollection = Nothing            'Interface pour extraire les segments.
        Dim pSegCollAdd As ISegmentCollection = Nothing         'Interface pour ajouter les segments.
        Dim pPath As IPath = Nothing                            'Interface contenant une ligne. 
        Dim pRelOp As IRelationalOperator = Nothing             'Interface qui permet de vérifier une relation spatiale.
        Dim pTopoOP As ITopologicalOperator2 = Nothing          'Interface qui permet d'effectuer des opérations spatiales.
        Dim pPoint As IPoint = Nothing                          'Interface contenant un point.
        Dim bOk As Boolean = False                              'Indique si un triangle a été trouvé.

        Try
            'Interface pour vérifier la relation spatiale
            pRelOp = CType(pPolyline, IRelationalOperator)

            'Créer la polyligne des triangles vides
            pTriangles = New Polyline
            pTriangles.SpatialReference = pPolyline.SpatialReference

            'Créer la polyligne des lignes vides
            pLignes = New Polyline
            pLignes.SpatialReference = pPolyline.SpatialReference

            'Interface pour extraire des lignes
            pGeomColl = CType(pPolylineGen, IGeometryCollection)
            'Interface pour ajouter des lignes
            pGeomCollAdd = CType(pLignes, IGeometryCollection)
            'Traiter toutes les lignes
            For i = 0 To pGeomColl.GeometryCount - 1
                'Définir la ligne à traiter
                pPath = CType(pGeomColl.Geometry(i), IPath)
                pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)
                'Vérifier si seulement 2 sommets
                If pPointColl.PointCount = 2 Then
                    'Vérifier si la ligne ne touche pas à la polyligne
                    If pRelOp.Disjoint(pPath.FromPoint) And pRelOp.Disjoint(pPath.ToPoint) Then
                        'Ajouter la ligne dans les polylignes à détruire
                        pGeomCollAdd.AddGeometry(pPath)
                    End If
                End If
            Next

            'Traiter tant que l'on trouve des triangles
            Do
                'Initialiser
                bOk = False
                'Simplifier les lignes
                pTopoOP = CType(pLignes, ITopologicalOperator2)
                pTopoOP.IsKnownSimple_2 = False
                pTopoOP.Simplify()
                'Interface pour extraire les lignes
                pGeomColl = CType(pLignes, IGeometryCollection)
                'Interface pour ajouter des triangles
                pGeomCollAdd = CType(pTriangles, IGeometryCollection)
                'Traiter toutes les lignes
                For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                    'Définir une ligne
                    pPath = CType(pGeomColl.Geometry(i), IPath)
                    pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)
                    'Vérifier si la ligne est un triangle fermée
                    If pPath.IsClosed And pPointColl.PointCount = 4 Then
                        'Enlever le triangle
                        pGeomColl.RemoveGeometries(i, 1)
                        'Ajouter le triangle
                        pGeomCollAdd.AddGeometry(pPath)
                        'Indiquer qu'un triangle a été trouvé
                        bOk = True
                    End If
                Next
            Loop While bOk

            'Simplifier les lignes
            pTopoOP = CType(pLignes, ITopologicalOperator2)
            pTopoOP.IsKnownSimple_2 = False
            pTopoOP.Simplify()
            'Interface pour extraire les lignes
            pGeomColl = CType(pLignes, IGeometryCollection)
            'Interface pour ajouter des triangles
            pGeomCollAdd = CType(pTriangles, IGeometryCollection)
            'Traiter toutes les lignes
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Définir une ligne
                pPath = CType(pGeomColl.Geometry(i), IPath)
                pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)
                'Vérifier si la ligne est un triangle non fermée
                If pPointColl.PointCount = 3 Then
                    'Fermer le triangle
                    pPointColl.AddPoint(pPointColl.Point(0))
                    pPath = CType(pPointColl, IPath)
                End If
                'Vérifier si la ligne est fermée
                If pPath.IsClosed Then
                    'Ajouter le triangle
                    pGeomCollAdd.AddGeometry(pPath)
                End If
            Next

            'Créer la polyligne des lignes vides
            pLignes = New Polyline
            pLignes.SpatialReference = pPolyline.SpatialReference

            'Interface pour extraire les triangles
            pGeomColl = CType(pTriangles, IGeometryCollection)
            'Interface pour ajouter des lignes
            pGeomCollAdd = CType(pLignes, IGeometryCollection)

            'Traiter tous les triangles
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Définir une ligne vide
                pPath = New Path
                pPath.SpatialReference = pPolyline.SpatialReference
                'Interface pour extraire un segment
                pSegColl = CType(pGeomColl.Geometry(i), ISegmentCollection)
                'Interface pour ajouter un segment
                pSegCollAdd = CType(pPath, ISegmentCollection)
                'Vérifier si le segment 0 est plus grand que le 1
                If pSegColl.Segment(0).Length > pSegColl.Segment(1).Length Then
                    'Vérifier si le segment 0 est plus grand que le 2
                    If pSegColl.Segment(0).Length > pSegColl.Segment(2).Length Then
                        'Ajouter le segment 0
                        pSegCollAdd.AddSegment(pSegColl.Segment(0))
                    Else
                        'Ajouter le segment 2
                        pSegCollAdd.AddSegment(pSegColl.Segment(2))
                    End If
                    'Si le segment 1 est plus grand que le 0
                Else
                    If pSegColl.Segment(1).Length > pSegColl.Segment(2).Length Then
                        'Ajouter le segment 1
                        pSegCollAdd.AddSegment(pSegColl.Segment(1))
                    Else
                        'Ajouter le segment 2
                        pSegCollAdd.AddSegment(pSegColl.Segment(2))
                    End If
                End If
                'Ajouter la ligne
                pGeomCollAdd.AddGeometry(pPath)
            Next

            ''Interface pour traiter la topologie
            'pTopoOP = CType(pPolylineGen, ITopologicalOperator2)
            ''Éliminer les lignes des triangles de la polyligne de généralisation
            'pPolylineGen = CType(pTopoOP.Difference(pLignes), IPolyline)

            'Définir un nouveau point vide
            pPoint = New Point
            pPoint.SpatialReference = pPolyline.SpatialReference
            'Interface pour vérifier une relation spatiale
            pRelOp = CType(pLignes, IRelationalOperator)
            'Interface pour extraire les lignes
            pGeomColl = CType(pLignes, IGeometryCollection)
            'Traiter toutes les lignes
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Définir une ligne
                pPath = CType(pGeomColl.Geometry(i), IPath)
                pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)
                'Vérifier si la ligne est une droite
                If pPointColl.PointCount = 2 Then
                    'Définir le centre de la droite
                    pPoint.X = (pPath.FromPoint.X + pPath.ToPoint.X) / 2
                    pPoint.Y = (pPath.FromPoint.Y + pPath.ToPoint.Y) / 2
                    'Vérifier si le point touche les lignes
                    If Not pRelOp.Disjoint(pPoint) Then
                        'Enlever la ligne
                        pGeomColl.RemoveGeometries(i, 1)
                    End If
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTriangles = Nothing
            pLignes = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pSegColl = Nothing
            pSegCollAdd = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pTopoOP = Nothing
            pPoint = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le polygone de généralisation.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Interface contenant le polygone à traiter.</param>
    ''' <param name="pBagDroites">Interface contenant les droites de Delaunay.</param>
    ''' <param name="dLargMin">Contient la largeur de généralisation minimum.</param>
    ''' <param name="dLongMin">Contient la longueur de généralisation minimum.</param>
    ''' <param name="dSupMin">Contient la superficie de généralisation minimum.</param>
    ''' <param name="pPolylineErr">Interface contenant la polyligne d'erreur de généralisation.</param>
    ''' <param name="pPolygonGen">Interface contenant le polygone de généralisation.</param>
    ''' 
    Private Shared Sub CreerPolygoneGeneralisation(ByVal pPolygon As IPolygon, pBagDroites As IGeometryBag, ByVal dLargMin As Double, ByVal dLongMin As Double, ByVal dSupMin As Double,
                                                   ByRef pPolylineErr As IPolyline, ByRef pPolygonGen As IPolygon)
        'Déclarer les variables de travail
        Dim pPolygonTmp As IPolygon = Nothing                   'Interface contenant un polygone de travail.
        Dim pBagDroitesTmp As IGeometryBag = Nothing            'Interface contenant les droites de Delaunay temporaire.
        Dim pBagSommetsPolygone As IGeometryBag = Nothing       'Interface contenant les sommets du polygone à traiter.
        Dim pDictLiens As New Dictionary(Of Integer, Noeud)     'Dictionnaire contenant l'information des liens entre les sommets du polygone à traiter.
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionnaire contenant les numéros des droites utilisé pour la généralisation.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface qui permet d'extraire des composantes de gémétrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface qui permet d'ajouter des composantes de gémétrie.
        Dim pRing As IRing = Nothing                            'Interface contenant un anneau. 
        Dim pArea As IArea = Nothing                            'Interface qui permet de calculer la superfie.
        Dim pTopoOP As ITopologicalOperator2 = Nothing          'Interface qui permet d'effectuer des opérations spatiales.

        Try
            'Vérifier s'il n'y a pas d'erreur de généralisation
            If pPolylineErr.IsEmpty Then
                'Redéfinir le polygon initiale
                pPolygonGen = pPolygon

                'S'il y a de la généralisation à effectuer
            Else
                'Créer le polygone de généralisation vide
                pPolygonGen = New Polygon
                pPolygonGen.SpatialReference = pPolygon.SpatialReference

                'Interface pour extraire les anneaux
                pGeomColl = CType(pPolygon, IGeometryCollection)

                'Traiter tous les anneaux du polygone
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Interface contenant un anneau du polygone 
                    pRing = CType(pGeomColl.Geometry(i), IRing)

                    'Créer le Bag des sommets de l'anneau et le dictionnaire des sommets contenant les sommets précédents et suivants de la ligne à traiter
                    Call CreerBagSommetsLigne(pRing, pBagSommetsPolygone, pDictLiens)
                    'Créer le dictionnaire contenant l'information des liens entre les sommets de la ligne
                    Call CreerDictLiensSommetsLignes(pBagSommetsPolygone, pBagDroites, pDictLiens, pBagDroitesTmp)
                    'Créer le dictionnaire des droites de généralisation utilisé pour généraliser le polygone
                    Call CreerDictDroiteGeneralisation(pPolylineErr, pBagDroites, pDictDroites)
                    'Créer le polygone de généralisation de l'anneau
                    Call CreerAnneauGeneralisation(pRing, pDictLiens, pDictDroites, pPolygonTmp)

                    'Vérifier si c'est le premier anneau
                    If i = 0 Then
                        'Définir le polygone extérieur
                        pPolygonGen = pPolygonTmp

                        'Si ce n'est pas le premier anneau
                    Else
                        'Interface pour enlever les polygones intérieurs
                        pTopoOP = CType(pPolygonGen, ITopologicalOperator2)
                        'Enlever le polygone intérieur au polygone extérieur
                        pPolygonGen = CType(pTopoOP.Difference(pPolygonTmp), IPolygon)
                    End If
                Next
            End If

            'Détruire les anneaux plus petites que les dimensions minimales
            '----------------------------
            'Créer un nouveau polygone vide
            pPolygonTmp = New Polygon
            pPolygonTmp.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter l'anneau dans le polygone
            pGeomCollAdd = CType(pPolygonTmp, IGeometryCollection)

            'Simplifier le polygone de gnéralisation
            pTopoOP = TryCast(pPolygonGen, ITopologicalOperator2)
            pTopoOP.IsKnownSimple_2 = False
            pTopoOP.Simplify()

            'Interface pour extraire les composantes du polygone de généralisation
            pGeomColl = CType(pPolygonGen, IGeometryCollection)

            'Traiter toutes les composantes du polygone de généralisation
            For i = 0 To pGeomColl.GeometryCount - 1
                'Interface pour calculer la superficie
                pArea = CType(pGeomColl.Geometry(i), IArea)
                'Vérifier si la superficie à la dimension minimale
                If Math.Abs(pArea.Area) > dSupMin Then
                    'Conserver l'anneau
                    pGeomCollAdd.AddGeometry(CType(pArea, IGeometry))
                End If
            Next

            'Redéfinir le polygone de généralisation
            pPolygonGen = pPolygonTmp

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPolygonTmp = Nothing
            pBagDroitesTmp = Nothing
            pBagSommetsPolygone = Nothing
            pDictLiens = Nothing
            pDictDroites = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pRing = Nothing
            pTopoOP = Nothing
            pArea = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le polygone de généralisation d'un anneau.
    ''' </summary>
    ''' 
    ''' <param name="pRing">Interface contenant l'anneau à généraliser.</param>
    ''' <param name="pDictLiens">Dictionnaire contenant l'information des sommets de la ligne à traiter.</param>
    ''' <param name="pDictDroites">Interface contenant la polyligne d'erreurs de généralisation.</param>
    ''' <param name="pPolygonGen">Interface contenant le polygone de généralisation.</param>
    ''' 
    Private Shared Sub CreerAnneauGeneralisation(ByVal pRing As IRing, ByVal pDictLiens As Dictionary(Of Integer, Noeud), ByVal pDictDroites As Dictionary(Of Integer, Integer), _
                                                 ByRef pPolygonGen As IPolygon)
        'Déclarer les variables de travail
        Dim pNoeudLien As Noeud = Nothing                   'Contient l'information d'un lien entre les sommets.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pPointCollAdd As IPointCollection = Nothing     'Interface utilisé pour ajouter les sommets d'une géométrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface qui permet d'ajouter des composantes de gémétrie.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne. 
        Dim pTopoOP As ITopologicalOperator2 = Nothing      'Interface qui permet d'effectuer des opérations spatiales.
        Dim pPoint As IPoint = Nothing                      'Interface contenant le point du centre de la droite.
        Dim pPointA As IPoint = Nothing                     'Interface contenant le premier sommet d'une droite.
        Dim pPointB As IPoint = Nothing                     'Interface contenant le deuxième sommet d'une droite.
        Dim pDroite As Droite = Nothing                     'Objet contenant la droite à traiter.

        Try
            'Créer un polygone vide
            pPolygonGen = New Polygon
            pPolygonGen.SpatialReference = pRing.SpatialReference
            'Interface pour ajouter les sommets de l'anneau
            pGeomCollAdd = CType(pPolygonGen, IGeometryCollection)

            'Créer un nouvel anneau vide
            pPath = New Ring
            pPath.SpatialReference = pRing.SpatialReference
            'Interface pour ajouter les sommets de l'anneau
            pPointCollAdd = CType(pPath, IPointCollection)

            'Interface pour extraire les sommets de l'anneau
            pPointColl = CType(pRing, IPointCollection)

            'Traiter tous les sommets
            For i = 0 To pPointColl.PointCount - 1
                'Définir l'information du lien
                pNoeudLien = pDictLiens.Item(i)

                'Vérifier si le sommet possède des droites
                If pNoeudLien.Droites.Count > 0 Then
                    'Traiter toutes les droites d'un sommet de l'anneau
                    For j = 0 To pNoeudLien.Droites.Count - 1
                        'Définir le numéro de sommet en lien
                        pDroite = pNoeudLien.Droites.Item(j)

                        'Vérifier si la droite doit être traité
                        If pDictDroites.ContainsKey(pDroite.No) Then
                            'Définir le point de début de la droite
                            pPointA = pDroite.PointDeb

                            'Définir le point de fin de la droite
                            pPointB = pDroite.PointFin

                            'Créer le point vide
                            pPoint = New Point
                            pPoint.SpatialReference = pRing.SpatialReference
                            'Définir le centre de la droite
                            pPoint.X = (pPointA.X + pPointB.X) / 2
                            pPoint.Y = (pPointA.Y + pPointB.Y) / 2
                            'Ajouter le sommet tel quel dans la ligne
                            pPointCollAdd.AddPoint(pPoint)

                            'Si la droite ne doit pas être traité
                        Else
                            'Ajouter le sommet tel quel dans l'anneau
                            pPointCollAdd.AddPoint(pPointColl.Point(i))
                        End If
                    Next

                    'Si le sommet ne possède pas de lien avec un autre sommet
                Else
                    'Ajouter le sommet tel quel dans l'anneau
                    pPointCollAdd.AddPoint(pPointColl.Point(i))
                End If
            Next

            'Ajouter la l'anneau de généralisation
            pGeomCollAdd.AddGeometry(pPath)

            'Simplifier le polygone de généralisation
            pTopoOP = CType(pPolygonGen, ITopologicalOperator2)
            pTopoOP.IsKnownSimple_2 = False
            pTopoOP.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pNoeudLien = Nothing
            pGeomCollAdd = Nothing
            pPath = Nothing
            pTopoOP = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pPoint = Nothing
            pPointA = Nothing
            pPointB = Nothing
            pDroite = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Function qui permet de créer et retourner le dictionnaire des droites des triangles de Delaunay à utiliser pour effectuer la généralisation.
    ''' </summary>
    ''' 
    ''' <param name="pPolylineErr">Interface contenant la polyligne d'erreur de généralisation.</param>
    ''' <param name="pBagLignesTriangles">Interface contenant les les droites des triangles de Delaunay.</param>
    ''' <param name="pDictDroites">Dictionnaire contenant les droites des triangles de Delaunay à traiter.</param>
    ''' 
    Private Shared Sub CreerDictDroiteGeneralisation(ByVal pPolylineErr As IPolyline, ByVal pBagLignesTriangles As IGeometryBag, _
                                                     ByRef pDictDroites As Dictionary(Of Integer, Integer))
        'Déclarer les variables de travail
        Dim pMultiPoint As IMultipoint = Nothing            'Interface contenant les points des lignes de généralisation.
        Dim pBagPoints As IGeometryBag = Nothing            'Interface contenant les points des lignes de généralisation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour simplifier.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter une relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement d'une relation spatiale.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes dans le Bag.
        Dim iSel As Integer = -1                            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1                            'Numéro de séquence de la géométrie en relation.

        Try
            'Créer le dictionnaire vide des droites
            pDictDroites = New Dictionary(Of Integer, Integer)

            'Créer un nouveau Bag vide des points de généralisation
            pMultiPoint = New Multipoint
            pMultiPoint.SpatialReference = pPolylineErr.SpatialReference
            'Interface pour ajouter les points dans le multipoint
            pPointColl = CType(pMultiPoint, IPointCollection)
            'Ajouter les points dans le multipoint
            pPointColl.AddPointCollection(CType(pPolylineErr, IPointCollection))
            'Simplifier le multipoint
            pTopoOp = CType(pMultiPoint, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Créer un nouveau Bag vide des points de généralisation
            pBagPoints = New GeometryBag
            pBagPoints.SpatialReference = pPolylineErr.SpatialReference
            'Interface pour ajouter les points dans le Bag
            pGeomColl = CType(pBagPoints, IGeometryCollection)
            'Ajouter les points dans le Bag
            pGeomColl.AddGeometryCollection(CType(pMultiPoint, IGeometryCollection))

            'Interface pour ajouter les lignes nulles dans le Bag des lignes des triangles
            pGeomColl = CType(pBagLignesTriangles, IGeometryCollection)

            'Interface pour traiter les relations spatiales entre les sommets du polygone et les lignes des triangles
            pRelOpNxM = CType(pBagPoints, IRelationalOperatorNxM)

            'Traiter la relation spatiale entre les sommets du polygone et les lignes des triangles
            'pRelResult = pRelOpNxM.Intersects(CType(pBagLignesTriangles, IGeometryBag))
            pRelResult = pRelOpNxM.Within(CType(pBagLignesTriangles, IGeometryBag))

            '---------------------------------------------------
            'Créer le dictionnaire des droites
            '---------------------------------------------------
            'Trier le résultat selon les lignes des triangles
            pRelResult.SortRight()

            'Traiter toutes les relations Lignes-Sommets
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)
                'Debug.Print(iSel.ToString & "-" & iRel.ToString)

                'Vérifier si la ligne existe
                If Not pDictDroites.ContainsKey(iRel) Then
                    'Ajouter la ligne dans le dictionnaire
                    pDictDroites.Add(iRel, iRel)
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pMultiPoint = Nothing
            pBagPoints = Nothing
            pTopoOp = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
        End Try
    End Sub
#End Region
End Class

''' <summary>
''' Classe qui permet de créer et gérer la triangulation de Delaunay.
''' </summary>
Public Class TriangulationDelaunay

#Region "Routines et fonctions publiques"
    ''' <summary>
    ''' Routine qui permet d'ajouter les sommets de connexion des lignes des triangles de Delaunay à la polyligne.
    ''' </summary>
    ''' 
    '''<param name="pLignes">Interface contenant les lignes des triangles de Delaunay.</param>
    '''<param name="pPolyline">Polyligne utilisée pour effectuer la généralisation de ligne.</param>
    ''' 
    Private Shared Sub ConnecterLignesPolyligne(ByVal pLignes As IPolyline, ByRef pPolyline As IPolyline)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface contenant les lignes des triangles de Delaunay.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets de connexion à la polyligne.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne des triangles de Delaunay.
        Dim pHitTest As IHitTest = Nothing                  'Interface pour tester la présence du sommet recherché
        Dim pPathColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes de la polyligne.
        Dim pNewPoint As IPoint = Nothing                   'Interface contenant le nouveau point trouvé (pas utilisé).
        Dim dDistance As Double = Nothing                   'Interface contenant la distance calculée entre le point de recherche et le sommet trouvé.
        Dim nNumeroPartie As Integer = Nothing              'Numéro de partie trouvée.
        Dim nNumeroSommet As Integer = Nothing              'Numéro de sommet de la partie trouvée.
        Dim bCoteDroit As Boolean = Nothing                 'Indiquer si le point trouvé est du côté droit de la géométrie.

        Try
            'Interface pour vérifier la connexion
            pHitTest = CType(pPolyline, IHitTest)

            'Interface pour extraire les anneaux du polygone
            pPathColl = CType(pPolyline, IGeometryCollection)

            'Interface pour extraire les lignes des triangles
            pGeomColl = CType(pLignes, IGeometryCollection)

            'Traiter toutes les lignes du triangle
            For i = 0 To pGeomColl.GeometryCount - 1
                'Interface contenant une ligne de triangle
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Vérifier si le premier point de la ligne n'intersecte pas avec un sommet de la polyligne
                If Not pHitTest.HitTest(pPath.FromPoint, 0.002, esriGeometryHitPartType.esriGeometryPartVertex, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Vérifier si le premier point de la ligne intersecte avec la limite de la polyligne
                    If pHitTest.HitTest(pPath.FromPoint, 0.002, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                        'Interface pour extraire le sommet de la composante de la polyligne
                        pPointColl = CType(pPathColl.Geometry(nNumeroPartie), IPointCollection)

                        'Insérer un nouveau sommet
                        pPointColl.InsertPoints(nNumeroSommet + 1, 1, pPath.FromPoint)
                    End If
                End If

                'Vérifier si le dernier point de la ligne n'intersecte pas avec un sommet de la polyligne
                If Not pHitTest.HitTest(pPath.ToPoint, 0.002, esriGeometryHitPartType.esriGeometryPartVertex, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                    'Vérifier si le dernier point de la ligne intersecte avec la limite de la polyligne
                    If pHitTest.HitTest(pPath.ToPoint, 0.002, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                    pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                        'Interface pour extraire le sommet de la composante de la polyligne
                        pPointColl = CType(pPathColl.Geometry(nNumeroPartie), IPointCollection)

                        'Insérer un nouveau sommet
                        pPointColl.InsertPoints(nNumeroSommet + 1, 1, pPath.ToPoint)
                    End If
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPointColl = Nothing
            pPath = Nothing
            pHitTest = Nothing
            pPathColl = Nothing
            pNewPoint = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer le squelette d'une ligne en utilisant les lignes des triangles de Delaunay.
    ''' </summary>
    ''' 
    '''<param name="pPath"> Ligne utilisée pour créer le squelette.</param>
    '''<param name="dLargMin"> Largeur minimum utilisée pour créer le squelette.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour créer le squelette.</param>
    '''<param name="pSquelette"> Interface contenant le squelette de la ligne.</param>
    '''<param name="pSqueletteEnv"> Interface contenant le squelette de la ligne avec son enveloppe.</param>
    '''<param name="pDictLiens"> Dictionnaire contenant l'information des sommets de la ligne à traiter.</param>
    '''<param name="pBagDroites"> Interface contenant le Bag des droites des triangles de Delaunay utilisées pour créer le squelette de la ligne.</param>
    '''<param name="pBagDroitesEnv"> Interface contenant les droites des triangles de Delaunay avec son enveloppe.</param>
    ''' 
    Public Shared Sub CreerSqueletteLigneDelaunay(ByVal pPath As IPath, ByVal dLargMin As Double, ByVal dLongMin As Double, _
                                                  ByRef pSquelette As IPolyline, ByRef pDictLiens As Dictionary(Of Integer, Noeud), _
                                                  ByRef pSqueletteEnv As IPolyline, ByRef pBagDroites As IGeometryBag, ByRef pBagDroitesEnv As IGeometryBag)
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                    'Interface contenant la polyline de la ligne à traiter.
        Dim pDroites As IPolyline = Nothing                     'Interface contenant les droites des triangles de Delaunay de la ligne à traiter.
        Dim pBagDroitesCoteDroit As IGeometryBag = Nothing      'Interface contenant le Bag des droites des triangles de Delaunay du côté droit de la ligne à traiter.
        Dim pBagSommets As IGeometryBag = Nothing               'Interface contenant le Bag des sommets de la ligne à traiter.
        Dim pBagSquelette As IGeometryBag = Nothing             'Interface contenant le Bag des lignes du squelette de la polyligne.
        Dim pBagSquelettePrimaire As IGeometryBag = Nothing     'Interface contenant le Bag des lignes du squelette primaire.
        Dim pBagSqueletteSimple As IGeometryBag = Nothing       'Interface contenant le Bag des lignes simples du squelette primaire.
        Dim pBagSqueletteBase As IGeometryBag = Nothing         'Interface contenant le Bag des lignes du squelette de base.
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionnaire contenant les numéros des droites à utiliser pour effectuer la généralisation.
        Dim pTopoOp As ITopologicalOperator5 = Nothing          'Interface pour effectuer des opérations spatiales.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les lignes de la polyligne.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface pour ajouter les lignes du squelette de la polyligne.
        Dim pRelOp As IRelationalOperator = Nothing             'Interface utiliser pour vérifier une relation spatiale.
        Dim pDroite As IPolyline = Nothing                      'Interface contenant une droite du Bag des droites.
        Dim pPointColl As IPointCollection = Nothing            'Interface pour traiter les sommets du squelette.
        Dim pClone As IClone = Nothing                          'Interface pour cloner une géométrie.
        Dim bIntersect As Boolean = False                       'Indiquer si un sommet du squelette intersecte le Bag des droites ou la ligne traitée.

        Try
            'Créer la polyligne vide
            pPolyline = New Polyline
            pPolyline.SpatialReference = pPath.SpatialReference
            'Interface pour ajouter des lignes dans une polyligne
            pGeomColl = CType(pPolyline, IGeometryCollection)
            'Ajouter la ligne dans la polyligne
            pGeomColl.AddGeometry(pPath)

            'Créer les droites des triangles de Delaunay de la ligne à traiter
            pDroites = CreerPolyligneTrianglesDelaunay(pPolyline)

            'Connecter les droites des triangles avec la ligne à traiter
            Call ConnecterLignesPolyligne(pDroites, pPolyline)

            'Créer le Bag des droites des triangles de Delaunay de la polyligne
            Call CreerBagLignesTriangles(pDroites, pBagDroitesEnv)

            '-----------------------------------------
            'Taiter la ligne-côté droit
            '-----------------------------------------
            'Créer le Bag des sommets et le dictionnaire des sommets contenant les sommets précédents et suivants de la ligne à traiter
            Call CreerBagSommetsLigne(pPath, pBagSommets, pDictLiens)
            'Créer le dictionnaire contenant l'information des liens entre les sommets de la ligne
            Call CreerDictLiensSommetsLignes(pBagSommets, pBagDroitesEnv, pDictLiens, pBagDroites)
            'Créer le Bag des lignes du squelette primaire et le Bag des droites du côté droit
            Call CreerBagLignesSquelettePrimaire(pPath, pDictLiens, pBagSquelettePrimaire, pBagDroitesCoteDroit)
            'Interface pour ajouter les lignes du squelette complet de la ligne à traiter
            pGeomCollAdd = CType(pBagSquelettePrimaire, IGeometryCollection)

            '-----------------------------------------
            'Créer le squelette de base
            '-----------------------------------------
            'Créer le Bag des lignes simples et doubles du squelette primaire
            'Le Bag des lignes doubles correspond aux lignes en double du squelette de base
            Call CreerBagLignesSimplesDoubles(pPolyline, pBagSquelettePrimaire, pBagSqueletteSimple, pBagSqueletteBase)
            'Traiter les lignes simples et les ajouter dans les lignes du squelette de base
            Call TraiterLignesSimples(pPolyline, pBagSqueletteSimple, pBagSqueletteBase)
            'Créer un nouveau squelette vide
            pSqueletteEnv = New Polyline
            pSqueletteEnv.SpatialReference = pPolyline.SpatialReference
            'Interface pour construire le squelette de base selon Delaunay
            pTopoOp = TryCast(pSqueletteEnv, ITopologicalOperator5)
            'Construire le squelette de base selon Delaunay
            pTopoOp.ConstructUnion(CType(pBagSqueletteBase, IEnumGeometry))

            '-----------------------------------------
            'Enlever les droites dont les 2 extrémités ne touchent pas à la ligne à traiter
            'du Bag des droites de la ligne-côté droit
            '-----------------------------------------
            'Interface pour vérifier la connexion avec la ligne à traiter
            pRelOp = CType(pPolyline, IRelationalOperator)
            'Interface pour extraire les lignes du Bag
            pGeomColl = CType(pBagDroites, IGeometryCollection)
            'Traiter toutes les droites du Bag
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Définir une droite du Bag des droites
                pDroite = CType(pGeomColl.Geometry(i), IPolyline)
                'Vérifier si une ou les 2 extrémités de la droite est disjoint de la ligne à traiter
                If pRelOp.Disjoint(pDroite.FromPoint) Or pRelOp.Disjoint(pDroite.ToPoint) Then
                    'Enlever la droite du Bag
                    pGeomColl.RemoveGeometries(i, 1)
                End If
            Next

            '-----------------------------------------
            'Créer le squelette de la ligne-côté droit
            '-----------------------------------------
            'Interface pour cloner le squelette
            pClone = CType(pSqueletteEnv, IClone)
            'Cloner le squelette
            pSquelette = CType(pClone.Clone, IPolyline)

            'Interface pour extraire les lignes du Bag
            pGeomColl = CType(pBagDroites, IGeometryCollection)
            'Interface pour extraire les sommets du squelette
            pPointColl = CType(pSquelette, IPointCollection)
            'Traiter tous les sommets du squelette
            For i = pPointColl.PointCount - 1 To 0 Step -1
                'Interface pour vérifier la relation spatiale
                pRelOp = CType(pPointColl.Point(i), IRelationalOperator)
                'Initialiser la variable d'intersection
                bIntersect = False
                'Traiter toutes les géométries du Bag
                For j = 0 To pGeomColl.GeometryCount - 1
                    'Définir une droite du Bag des droites
                    pDroite = CType(pGeomColl.Geometry(j), IPolyline)
                    'Vérifier si le point intersect la droite ou la ligne traitée
                    If Not (pRelOp.Disjoint(pDroite) And pRelOp.Disjoint(pPolyline)) Then
                        'Indiquer si le sommet intersect le Bag des droites ou la ligne traitée
                        bIntersect = True
                        'Sortir de la boucle
                        Exit For
                    End If
                Next
                'si le point n'intersecte pas le Bag des droites ou la ligne traitée
                If Not bIntersect Then
                    'Enlever le sommet
                    pPointColl.RemovePoints(i, 1)
                End If
            Next
            'Enlever les géométries invalides
            pTopoOp = CType(pSquelette, ITopologicalOperator5)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()
            'Enlever les géométries invalides
            pTopoOp = CType(pSqueletteEnv, ITopologicalOperator5)
            pSquelette = CType(pTopoOp.Intersect(pSquelette, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

            'pBagDroitesEnv = pBagSquelettePrimaire

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pBagDroitesCoteDroit = Nothing
            pBagSommets = Nothing
            pBagSquelette = Nothing
            pBagSquelettePrimaire = Nothing
            pBagSqueletteSimple = Nothing
            pBagSqueletteBase = Nothing
            pDictDroites = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPointColl = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pDroite = Nothing
            pClone = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer le squelette d'une ligne en utilisant les lignes des triangles de Delaunay.
    ''' </summary>
    ''' 
    '''<param name="pPolyline"> Polyligne utilisée pour créer le squelette.</param>
    '''<param name="dLargMin"> Largeur minimum utilisée pour créer le squelette.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour créer le squelette.</param>
    '''<param name="pSquelette"> Interface contenant le squelette de la ligne.</param>
    '''<param name="pSqueletteEnv"> Interface contenant le squelette de la ligne avec son enveloppe.</param>
    '''<param name="pDictLiens"> Dictionnaire contenant l'information des sommets de la ligne à traiter.</param>
    '''<param name="pBagDroites"> Interface contenant le Bag des droites des triangles de Delaunay utilisées pour créer le squelette de la ligne.</param>
    '''<param name="pBagDroitesEnv"> Interface contenant les droites des triangles de Delaunay avec son enveloppe.</param>
    ''' 
    Public Shared Sub CreerSqueletteLigneDelaunay(ByVal pPolyline As IPolyline, ByVal dLargMin As Double, ByVal dLongMin As Double, _
                                                  ByRef pSquelette As IPolyline, ByRef pDictLiens As Dictionary(Of Integer, Noeud), _
                                                  ByRef pSqueletteEnv As IPolyline, ByRef pBagDroites As IGeometryBag, ByRef pBagDroitesEnv As IGeometryBag)
        'Déclarer les variables de travail
        Dim pPolylineEnv As IPolyline = Nothing                 'Interface contenant la polyline et l'enveloppe de la polyligne à généraliser.
        Dim pSqueletteTmp As IPolyline = Nothing                'Interface contenant le squelette temporaire.
        Dim pLignesEnv As IPolyline = Nothing                   'Interface contenant les lignes des triangles de Delaunay pour la polyligne avec son enveloppe.
        Dim pBagDroitesTmp As IGeometryBag = Nothing            'Interface contenant le Bag des droites des triangles de Delaunay temporaires.
        Dim pBagSommets As IGeometryBag = Nothing               'Interface contenant le Bag des sommets de la ligne à traiter.
        Dim pBagSommetsTmp As IGeometryBag = Nothing            'Interface contenant le Bag des sommets temporaires de la ligne à traiter.
        Dim pBagSquelette As IGeometryBag = Nothing             'Interface contenant le Bag des lignes du squelette de la polyligne.
        Dim pBagSqueletteTmp As IGeometryBag = Nothing          'Interface contenant le Bag des lignes du squelette temporaire.
        Dim pBagSquelettePrimaire As IGeometryBag = Nothing     'Interface contenant le Bag des lignes du squelette primaire.
        Dim pBagSqueletteSimple As IGeometryBag = Nothing       'Interface contenant le Bag des lignes simples du squelette primaire.
        Dim pBagSqueletteBase As IGeometryBag = Nothing         'Interface contenant le Bag des lignes du squelette de base.
        Dim pDictLiensTmp As New Dictionary(Of Integer, Noeud)  'Dictionnaire contenant l'information des sommets de la polyligne temporaire.
        Dim pDictDroites As New Dictionary(Of Integer, Integer) 'Dictionnaire contenant les numéros des droites à utiliser pour effectuer la généralisation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour effectuer des opérations spatiales.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les lignes de la polyligne.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface pour ajouter les lignes du squelette de la polyligne.
        Dim pPath As IPath = Nothing                            'Interface contenant une ligne de la polyligne à traiter.
        Dim pRelOp As IRelationalOperator = Nothing             'Interface utiliser pour vérifier une relation spatiale.
        Dim pDroite As IPolyline = Nothing                      'Interface contenant une droite du Bag des droites.
        Dim pPointColl As IPointCollection = Nothing            'Interface pour traiter les sommets du squelette.
        Dim bIntersect As Boolean = False                       'Indiquer si un sommet du squelette intersecte le Bag des droites ou la ligne traitée.

        Try
            'Ajouter l'enveloppe à la polyligne à généraliser
            pPolylineEnv = AjouterEnveloppePolyligne(pPolyline, dLargMin)

            'Interface pour extraire les lignes de la polyligne
            pGeomColl = CType(pPolylineEnv, IGeometryCollection)

            'Créer les lignes des triangles de Delaunay de la polyligne avec son enveloppe
            pLignesEnv = CreerPolyligneTrianglesDelaunay(pPolylineEnv)

            'Connecter les lignes des triangles avec la polyligne à généraliser
            Call ConnecterLignesPolyligne(pLignesEnv, pPolylineEnv)

            'Créer le Bag des droites des triangles de Delaunay de la polyligne avec son enveloppe
            Call CreerBagLignesTriangles(pLignesEnv, pBagDroitesEnv)

            '-----------------------------------------
            'Traiter l'enveloppe de la ligne
            '-----------------------------------------
            'Définir la ligne de l'enveloppe à traiter
            pPath = CType(pGeomColl.Geometry(0), IPath)
            'Créer le Bag des sommets et le dictionnaire des sommets contenant les sommets précédents et suivants de la ligne à traiter
            Call CreerBagSommetsLigne(pPath, pBagSommetsTmp, pDictLiensTmp)
            'Créer le dictionnaire contenant l'information des liens entre les sommets de la ligne
            Call CreerDictLiensSommetsLignes(pBagSommetsTmp, pBagDroitesEnv, pDictLiensTmp, pBagDroitesTmp)
            'Créer le Bag des lignes du squelette primaire
            Call CreerBagLignesSquelettePrimaire(pPath, pDictLiensTmp, pBagSquelettePrimaire, pBagDroitesTmp)
            'Interface pour ajouter les lignes du squelette complet de la ligne avec l'enveloppe
            pGeomCollAdd = CType(pBagSquelettePrimaire, IGeometryCollection)

            '-----------------------------------------
            'Taiter la ligne - droite
            '-----------------------------------------
            'Définir la ligne à traiter
            pPath = CType(pGeomColl.Geometry(1), IPath)
            'Créer le Bag des sommets et le dictionnaire des sommets contenant les sommets précédents et suivants de la ligne à traiter
            Call CreerBagSommetsLigne(pPath, pBagSommets, pDictLiens)
            'Créer le dictionnaire contenant l'information des liens entre les sommets de la ligne
            Call CreerDictLiensSommetsLignes(pBagSommets, pBagDroitesEnv, pDictLiens, pBagDroites)
            'Créer le Bag des lignes du squelette primaire
            Call CreerBagLignesSquelettePrimaire(pPath, pDictLiens, pBagSquelette, pBagDroitesTmp)
            'Ajouter les lignes primaire temporaire
            pGeomCollAdd.AddGeometryCollection(CType(pBagSquelette, IGeometryCollection))

            '-----------------------------------------
            'Créer le Bag des droites de la ligne-droite
            '-----------------------------------------
            'Interface pour vérifier la connexion
            pRelOp = CType(pPolyline, IRelationalOperator)
            'Interface pour extraire les lignes du Bag
            pGeomColl = CType(pBagDroites, IGeometryCollection)
            'Traiter toutes les lignes du squelette
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Définir une droite du Bag des droites
                pDroite = CType(pGeomColl.Geometry(i), IPolyline)
                'Vérifier si le squelette est disjoint de la polyline
                If pRelOp.Disjoint(pDroite.FromPoint) Or pRelOp.Disjoint(pDroite.ToPoint) Then
                    'Enlever la ligne du squelette
                    pGeomColl.RemoveGeometries(i, 1)
                End If
            Next

            '-----------------------------------------
            'Traiter la ligne - gauche
            '-----------------------------------------
            'Inverser la ligne à traiter
            pPath.ReverseOrientation()
            'Créer le Bag des sommets et le dictionnaire des sommets contenant les sommets précédents et suivants de la ligne à traiter
            Call CreerBagSommetsLigne(pPath, pBagSommetsTmp, pDictLiensTmp)
            'Créer le dictionnaire contenant l'information des liens entre les sommets de la ligne
            Call CreerDictLiensSommetsLignes(pBagSommetsTmp, pBagDroitesEnv, pDictLiensTmp, pBagDroitesTmp)
            'Créer le Bag des lignes du squelette primaire
            Call CreerBagLignesSquelettePrimaire(pPath, pDictLiensTmp, pBagSqueletteTmp, pBagDroitesTmp)
            'Ajouter les lignes primaire temporaire
            pGeomCollAdd.AddGeometryCollection(CType(pBagSqueletteTmp, IGeometryCollection))

            '-----------------------------------------
            'Créer le squelette de base avec l'enveloppe
            '-----------------------------------------
            'Créer le Bag des lignes simples et doubles du squelette primaire
            'Le Bag des lignes doubles correspond aux lignes en double du squelette de base
            Call CreerBagLignesSimplesDoubles(pPolylineEnv, pBagSquelettePrimaire, pBagSqueletteSimple, pBagSqueletteBase)
            'Traiter les lignes simples et les ajouter dans les lignes du squelette de base
            Call TraiterLignesSimples(pPolylineEnv, pBagSqueletteSimple, pBagSqueletteBase)
            'Créer un nouveau squelette vide
            pSqueletteEnv = New Polyline
            pSqueletteEnv.SpatialReference = pPolyline.SpatialReference
            'Interface pour construire le squelette de base selon Delaunay
            pTopoOp = TryCast(pSqueletteEnv, ITopologicalOperator2)
            'Construire le squelette de base selon Delaunay
            pTopoOp.ConstructUnion(CType(pBagSqueletteBase, IEnumGeometry))

            '-----------------------------------------
            'Créer le squelette de la ligne-droite
            '-----------------------------------------
            'Créer un nouveau squelette vide
            pSqueletteTmp = New Polyline
            pSqueletteTmp.SpatialReference = pPolyline.SpatialReference
            'Interface pour construire le squelette de base selon Delaunay
            pTopoOp = TryCast(pSqueletteTmp, ITopologicalOperator2)
            'Construire le squelette de base selon Delaunay
            pTopoOp.ConstructUnion(CType(pBagSquelette, IEnumGeometry))

            'Définir le squelette finale
            'pSquelette = CType(pTopoOp.Intersect(pSqueletteEnv, esriGeometryDimension.esriGeometry1Dimension), IPolyline)
            pTopoOp = TryCast(pSqueletteEnv, ITopologicalOperator2)
            pSquelette = CType(pTopoOp.Intersect(pSqueletteTmp, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

            'Interface pour extraire les lignes du Bag
            pGeomColl = CType(pBagDroites, IGeometryCollection)
            'Interface pour extraire les sommets du squelette
            pPointColl = CType(pSquelette, IPointCollection)
            'Traiter tous les sommets du squelette
            For i = pPointColl.PointCount - 1 To 0 Step -1
                'Interface pour vérifier la relation spatiale
                pRelOp = CType(pPointColl.Point(i), IRelationalOperator)
                'Initialiser la variable d'intersection
                bIntersect = False
                'Traiter toutes les géométries du Bag
                For j = 0 To pGeomColl.GeometryCount - 1
                    'Définir une droite du Bag des droites
                    pDroite = CType(pGeomColl.Geometry(j), IPolyline)
                    'Vérifier si le point intersect la droite ou la ligne traitée
                    If Not (pRelOp.Disjoint(pDroite) And pRelOp.Disjoint(pPolyline)) Then
                        'Indiquer si le sommet intersect le Bag des droites ou la ligne traitée
                        bIntersect = True
                        'Sortir de la boucle
                        Exit For
                    End If
                Next
                'si le point n'intersecte pas le Bag des droites ou la ligne traitée
                If Not bIntersect Then
                    'Enlever le sommet
                    pPointColl.RemovePoints(i, 1)
                End If
            Next
            'Enlever les géométries invalides
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()
            'Enlever les géométries invalides
            pTopoOp = CType(pSqueletteEnv, ITopologicalOperator2)
            pSquelette = CType(pTopoOp.Intersect(pSquelette, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

            'Tester
            'pBagDroitesEnv = pBagSquelettePrimaire
            'pBagDroites = pBagSqueletteBase
            'pSquelette = pSqueletteEnv
            'pSquelette = pSqueletteTmp

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPolylineEnv = Nothing
            pSqueletteTmp = Nothing
            pLignesEnv = Nothing
            pBagDroitesTmp = Nothing
            pBagSommets = Nothing
            pBagSommetsTmp = Nothing
            pBagSquelette = Nothing
            pBagSqueletteTmp = Nothing
            pBagSquelettePrimaire = Nothing
            pBagSqueletteSimple = Nothing
            pBagSqueletteBase = Nothing
            pDictLiensTmp = Nothing
            pDictDroites = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPointColl = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pDroite = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner les lignes du squelette d'un polygone dans une Polyline.
    ''' Le squelette est créé à partir des lignes des triangles de Delaunay.
    ''' La liste des triangles de Delaunay est créée à partir de tous les sommets non duppliqués et triés de la géométrie spécifiée.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Polygone utilisée pour créer le squelette.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre le polygon et les éléments en relation.</param>
    '''<param name="dDistLat"> Distance latérale utilisée pour éliminer des sommets en trop.</param>
    '''<param name="dDistMin"> Distance minimum utilisée pour ajouter des sommets.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour éliminer des lignes du squelette trop petites.</param>
    ''' 
    Public Shared Sub CreerSquelettePolygoneDelaunay(ByVal pPolygon As IPolygon4, ByVal pPointsConnexion As IMultipoint, _
                                                     ByVal dDistLat As Double, ByVal dDistMin As Double, ByVal dLongMin As Double, _
                                                     ByRef pSquelette As IPolyline, ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pPointsNonConnexion As IMultipoint = Nothing    'Interface contenant les points du squelette qui doivent être connectés.
        Dim pLignes As IPolyline = Nothing                  'Interface contenant les lignes des triangles de Delaunay.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les anneaux extérieurs.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour ajouter les lignes du squelette.
        Dim pRingExt As IRing = Nothing                     'Interface contenant l'anneau extérieur.
        Dim pPolygonTmp As IPolygon = Nothing               'Interface contenant le polygone temporaire de traitement.
        Dim pSqueletteTmp As IPolyline = Nothing            'Interface contenant le squelette temporaire.
        Dim pGeomCollTmp As IGeometryCollection = Nothing   'Interface pour ajouter les anneaux extérieurs dans le polygone temporaire.
        Dim pPointsConnexionTmp As IMultipoint = Nothing    'Interface contenant les points de connexion temporaire.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface utilisé pour extraire les points de connexion temporaire.

        'Définir la valeur par défaut
        pSquelette = Nothing

        Try
            'Enlever les sommets en trop
            'pPolygon.Generalize(dDistLat)

            'Densifer les sommets du polygone
            pPolygon.Densify(dDistMin / 2, 0)

            'Créer la polyligne vide du squelette
            pSquelette = New Polyline
            pSquelette.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter les lignes du squelette
            pGeomCollAdd = CType(pSquelette, IGeometryCollection)

            'Interface pour extraire les anneaux extérieurs
            pGeomColl = CType(pPolygon.ExteriorRingBag, IGeometryCollection)

            'Traiter toutes les composantes
            For i = 0 To pGeomColl.GeometryCount - 1
                'Définir la composante
                pRingExt = CType(pGeomColl.Geometry(i), IRing)

                'Vérifier si l'anneau n'est pas vide
                If Not pRingExt.IsEmpty Then
                    'Créer un nouveau polygone vide
                    pPolygonTmp = New Polygon
                    pPolygonTmp.SpatialReference = pPolygon.SpatialReference
                    'Ajouter l'anneau extérieur
                    pGeomCollTmp = CType(pPolygonTmp, IGeometryCollection)
                    pGeomCollTmp.AddGeometry(pRingExt)
                    'Ajouter les anneaux intérieures
                    pGeomCollTmp.AddGeometryCollection(CType(pPolygon.InteriorRingBag(pRingExt), IGeometryCollection))

                    'Définir les lignes des triangles de Delaunay découpées selon le polygone
                    pLignes = CreerPolyligneTrianglesDelaunay(pPolygonTmp)

                    'Créer le squelette de base du polygone selon les lignes des triangles de Delaunay
                    Call CreerSqueletteBaseDelaunay(pPolygonTmp, pLignes, pSqueletteTmp, pBagDroites)

                    'Interface pour extraire les points d'intersection spécifique au polygone temporaire
                    pTopoOp = CType(pPointsConnexion, ITopologicalOperator2)

                    'Extraire les points d'intersection spécifique au polygone temporaire
                    pPointsConnexionTmp = CType(pTopoOp.Intersect(pPolygonTmp, esriGeometryDimension.esriGeometry0Dimension), IMultipoint)

                    'Connecter le squelette aux points de connexion des éléments en relation
                    'Call ConnecterSquelettePointsConnexion(pPointsConnexion, dDistMin, dLongMin, pSqueletteTmp)
                    'Call ConnecterSquelettePointsConnexion(pPointsConnexionTmp, pSqueletteTmp)
                    Call ConnecterSquelettePointsConnexion(pPolygon, pPointsConnexionTmp, pSqueletteTmp)

                    'Enlever les extrémités de lignes superflux dans le squelette
                    pPointsNonConnexion = EnleverExtremiteLigne(pPointsConnexion, pSqueletteTmp, dLongMin)

                    'Connecter les lignes non connectées du squelette avec la limite du polygone
                    Call ConnecterSquelettePolygone(pPolygonTmp, pPointsNonConnexion, pSqueletteTmp)

                    'Densifer les sommets du squelette
                    pSqueletteTmp.Densify(dDistMin, 0)

                    'Créer un nouveau squelette à partir des centres des droites du squelette initiale.
                    pSqueletteTmp = TraiterCentreDroites(pSqueletteTmp)

                    'Ajouter les lignes du squelette
                    pGeomCollAdd.AddGeometryCollection(CType(pSqueletteTmp, IGeometryCollection))
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLignes = Nothing
            pPointsNonConnexion = Nothing
            pGeomCollAdd = Nothing
            pRingExt = Nothing
            pPolygonTmp = Nothing
            pSqueletteTmp = Nothing
            pGeomCollTmp = Nothing
            pTopoOp = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Fonction qui permet de retourner le Polygone extérieur d'une géométrie à partir d'une largeur minimum.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour créer un polygone extérieure.</param>
    ''' <param name="dLargMin"> Largeur minimum utilisée pour créer le polygone.</param>
    ''' 
    ''' <returns>Polygon contenant le polygone extérieur à un polygone ou à une ligne.</returns>
    ''' 
    Public Shared Function CreerPolygoneExterieur(ByVal pGeometry As IGeometry, ByVal dLargMin As Double) As IPolygon4
        'Déclarer les variables de travail
        Dim pEnvelope As IEnvelope = Nothing            'Interface contenant l'enveloppe du polygone extérieur.
        Dim pPointColl As IPointCollection = Nothing    'Interface pour ajouter les sommets dans le polygone.
        Dim pTopoOp As ITopologicalOperator2 = Nothing  'Interface pour créer le polygone extérieure du polygone.

        'Créer le polygone extérieure vide
        CreerPolygoneExterieur = New Polygon
        CreerPolygoneExterieur.SpatialReference = pGeometry.SpatialReference

        Try
            'Définir l'enveloppe du polygone
            pEnvelope = pGeometry.Envelope

            'Agrandir l'enveloppe du polygone selon la largeur minimum doublée
            pEnvelope.Expand(dLargMin * 2, dLargMin * 2, False)

            'Interface pour ajouter les sommets dans le polygone
            pPointColl = CType(CreerPolygoneExterieur, IPointCollection)

            'Ajouter les sommets du polygone à partir de l'enveloppe
            pPointColl.AddPoint(pEnvelope.LowerLeft)
            pPointColl.AddPoint(pEnvelope.UpperLeft)
            pPointColl.AddPoint(pEnvelope.UpperRight)
            pPointColl.AddPoint(pEnvelope.LowerRight)
            pPointColl.AddPoint(pEnvelope.LowerLeft)

            'Vérifier si la géométrie est un polygone
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour créer le polygone extérieure du polygone
                pTopoOp = CType(CreerPolygoneExterieur, ITopologicalOperator2)

                'Créer le polygone extérieur
                CreerPolygoneExterieur = CType(pTopoOp.Difference(pGeometry), IPolygon4)
            End If

            'Densifier le polygone extérieur selon la largeur minimum
            CreerPolygoneExterieur.Densify(dLargMin, 0)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEnvelope = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Fonction qui permet de retourner la polyligne dans laquelle l'enveloppe a été ajoutée à partir d'une largeur minimum.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Polyligne utilisée pour ajouter son enveloppe.</param>
    ''' <param name="dLargMin"> Largeur minimum utilisée pour créer la polyligne avec son enveloppe.</param>
    ''' 
    ''' <returns>Polyline contenant la polyligne et son enveloppe.</returns>
    ''' 
    Public Shared Function AjouterEnveloppePolyligne(ByVal pPolyline As IPolyline, ByVal dLargMin As Double) As IPolyline
        'Déclarer les variables de travail
        Dim pEnvelope As IEnvelope = Nothing            'Interface contenant l'enveloppe de la polyligne.
        Dim pPointColl As IPointCollection = Nothing    'Interface pour ajouter les sommets dans le polygone.
        Dim pGeomColl As IGeometryCollection = Nothing  'Interface pour ajouter les composantes dans la polyligne.
        Dim pPath As IPath = Nothing                    'Interface contenant une composante de lignes

        'Créer le polygone extérieure vide
        AjouterEnveloppePolyligne = New Polyline
        AjouterEnveloppePolyligne.SpatialReference = pPolyline.SpatialReference

        Try
            'Interface pour ajouter les composantes dans la polyligne
            pGeomColl = CType(AjouterEnveloppePolyligne, IGeometryCollection)

            'Définir l'enveloppe de la polyligne
            pEnvelope = pPolyline.Envelope

            'Agrandir l'enveloppe de la polyligne selon la largeur minimum doublée
            pEnvelope.Expand(dLargMin * 2, dLargMin * 2, False)

            'Créer une nouvelle composante
            pPath = New Path
            pPath.SpatialReference = pPolyline.SpatialReference
            'Interface pour ajouter les sommets dans la polyligne
            pPointColl = CType(pPath, IPointCollection)
            'Ajouter les sommets de l'enveloppe
            pPointColl.AddPoint(pEnvelope.LowerLeft)
            pPointColl.AddPoint(pEnvelope.UpperLeft)
            pPointColl.AddPoint(pEnvelope.UpperRight)
            pPointColl.AddPoint(pEnvelope.LowerRight)
            pPointColl.AddPoint(pEnvelope.LowerLeft)
            'Ajouter la composante dans la polyligne
            pGeomColl.AddGeometry(pPath)

            'Densifier la polyligne selon la largeur minimum
            AjouterEnveloppePolyligne.Densify(dLargMin, 0)

            'Créer une nouvelle composante
            pPath = New Path
            pPath.SpatialReference = pPolyline.SpatialReference
            'Interface pour ajouter les sommets dans la polyligne
            pPointColl = CType(pPath, IPointCollection)
            'Ajouter les sommets de la polyligne
            pPointColl.AddPointCollection(CType(pPolyline, IPointCollection))
            'Ajouter la composante dans la polyligne
            pGeomColl.AddGeometry(pPath)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pEnvelope = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
            pPath = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Fonction qui permet d'identifier les limites du squelette qui correspondent aux sommets du squelettes qui sont non connectés entre eux.
    ''' </summary>
    ''' 
    ''' <param name="pSquelette">IPolyline contenant le squelette à traiter.</param>
    ''' 
    '''<returns>IMultiPoint contenant les limites du squelette.</returns>
    ''' 
    Public Shared Function IdentifierLimiteSquelette(ByRef pSquelette As IPolyline) As IMultipoint
        'Déclarer les variables de travail
        Dim pMultiPoint As IMultipoint = Nothing            'Interface contenant les limites de chaque ligne du squelette.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pGeomCollAdd As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les lignes d'extrémité.
        Dim pPath As IPath = Nothing                        'Interface contenant une composante de ligne.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier la connexion.
        Dim pPoint As IPoint = Nothing                      'Contient un sommet d'une ligne.
        Dim pListeSommets As New List(Of IPoint)            'Contient la liste des sommets des extrémités de ligne du squelette.
        Dim iNo As Integer = 0                              'Contient le numéro du sommet traité.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pGeometryBag As IGeometryBag = Nothing
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.
        Dim pDict As New Dictionary(Of Integer, Integer)

        'Par défaut, il n'y a pas de limite
        IdentifierLimiteSquelette = New Multipoint
        IdentifierLimiteSquelette.SpatialReference = pSquelette.SpatialReference

        Try
            'Vérifier si la ligne n'est pas vide
            If Not pSquelette.IsEmpty Then
                'Interface pour traiter toutes les lignes
                pGeomColl = CType(pSquelette, IGeometryCollection)

                'Créer un multipoint des limites des lignes vide
                pMultiPoint = New Multipoint
                pMultiPoint.SpatialReference = pSquelette.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pMultiPoint, IPointCollection)

                pGeometryBag = New GeometryBag
                pGeometryBag.SpatialReference = pSquelette.SpatialReference
                pGeomCollAdd = CType(pGeometryBag, IGeometryCollection)
                'Traiter toutes les lignes du squelette
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Définir la ligne
                    pPath = CType(pGeomColl.Geometry(i), IPath)

                    'Ajouter les extrémités de la ligne
                    pGeomCollAdd.AddGeometry(pPath.FromPoint)
                    pGeomCollAdd.AddGeometry(pPath.ToPoint)
                Next

                'Interface pour traiter les relations spatiales
                pRelOpNxM = CType(pGeometryBag, IRelationalOperatorNxM)
                'Traiter la relation spatiale entre les sommets du polygone et les lignes de généralisation en largeur
                pRelResult = pRelOpNxM.Intersects(CType(pGeometryBag, IGeometryBag))

                'Traiter toutes les relatins
                For i = 0 To pRelResult.RelationElementCount - 1
                    'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                    pRelResult.RelationElement(i, iSel, iRel)
                    'Vérifier si la géométrie en relation est différent de celle traitée
                    If iSel <> iRel Then
                        'Si la géométrie traitée n'est pas encore trouvée
                        If Not pDict.ContainsKey(iSel) Then
                            'Ajouté la géométrie trouvé avec sa relation
                            pDict.Add(iSel, iRel)
                            'Debug.Print(iSel.ToString & "-" & iRel.ToString)
                        End If
                    End If
                Next

                'Traiter toutes les extrémités de ligne du squelette
                For i = 0 To pGeomCollAdd.GeometryCount - 1
                    'Si l'extrémité de ligne n'a pas été trouvée
                    If Not pDict.ContainsKey(i) Then
                        'Ajouter l'extrémité comme un point de limite non connecté
                        pPointColl.AddPoint(CType(pGeomCollAdd.Geometry(i), IPoint))
                    End If
                Next

                'Retourner les limtes du squelette
                IdentifierLimiteSquelette = pMultiPoint
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPointColl = Nothing
            pMultiPoint = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pPoint = Nothing
            pListeSommets = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function qui permet de créer et retourner le GeometryBag contenant les polygones des triangles de Delaunay.
    ''' La liste des triangles de Delaunay est créée à partir de tous les sommets non duppliqués et triés de la géométrie spécifiée.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour créer le diagramme de Voronoi à partir de ses sommets.</param>
    ''' <param name="bDecouper">Indique si on doit découper les lignes du diagramme de Voronoi selon la zone de traitement.</param>
    ''' 
    ''' <returns>IGeometryBag contenant les polygones des triangles de Delaunay.</returns>
    ''' 
    Public Shared Function CreerBagTrianglesDelaunay(pGeometry As IGeometry, Optional ByVal bDecouper As Boolean = False) As IGeometryBag
        'Déclarer les variables de travail
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant les lignes du diagramme de Voronoi.
        Dim pTriangle As IPolygon = Nothing                 'Interface contenant un triangle de delaunay.
        Dim pZone As IGeometry = Nothing                    'Interface contenant la zone de traitement (Enveloppe ou Polygone de la géométrie).
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans le Bag.
        Dim pPointColl As IPointCollection = Nothing        'Interface contenant les sommets d'un triangle.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le polygone de Voronoi (ConvexHull).
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pListeTriangles As List(Of SimpleTriangle) = Nothing 'Liste des triangles de Delaunay pour créer le diagramme de Voronoi.
        Dim pArea As IArea = Nothing                        'Interface pour extraire la superficie d'un triangle.

        'Définir la valeur par défaut
        CreerBagTrianglesDelaunay = New GeometryBag
        CreerBagTrianglesDelaunay.SpatialReference = pGeometry.SpatialReference

        Try
            'Par défaut, la Zone de traitement est l'enveloppe de la géométrie
            pZone = pGeometry.Envelope

            'Vérifier si la géométrie est un polygone
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'La zone de traitement est le polygone de la géométrie
                pZone = pGeometry
            End If

            'Interface pour découper les triangles
            pTopoOp = TryCast(pZone, ITopologicalOperator2)

            'Créer la liste des sommets
            pListeSommets = CreerListeSommets(pGeometry)

            'Créer la liste des triangles de Delaunay pour créer le diagramme de Voronoi
            pListeTriangles = CreerListeTriangles(pGeometry.Envelope, pListeSommets)

            'Créer un nouveau Bag vide
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = pGeometry.SpatialReference

            'Interface pour ajouter des géométries dans le Bag
            pGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter tous les triangles
            For Each triangle As SimpleTriangle In pListeTriangles
                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pTriangle = New Polygon
                pTriangle.SpatialReference = pGeometry.SpatialReference

                'Interface pour ajouter des sommets
                pPointColl = CType(pTriangle, IPointCollection)

                'Ajouter le point A
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.A).X, .Y = pListeSommets.Item(triangle.A).Y})

                'Ajouter le point B
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.B).X, .Y = pListeSommets.Item(triangle.B).Y})

                'Ajouter le point C
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.C).X, .Y = pListeSommets.Item(triangle.C).Y})

                'Fermer le triangle
                pTriangle.Close()

                'Interface pour vérifier la superficie
                pArea = CType(pTriangle, IArea)
                'Inverser les coordonnées si la superficie est négative
                If pArea.Area < 0 Then pTriangle.ReverseOrientation()

                'Vérifier si on doit découper les lignes du diagramme de Voronoi
                If bDecouper Then
                    'Découper les lignes du polygone de Voronoi selon la zone de traitement
                    pTriangle = CType(pTopoOp.Intersect(pTriangle, esriGeometryDimension.esriGeometry2Dimension), IPolygon)
                End If

                'Ajouter le triangle
                pGeomColl.AddGeometry(pTriangle)
            Next

            'Définir le résultat
            CreerBagTrianglesDelaunay = pGeometryBag

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeometryBag = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pTriangle = Nothing
            pZone = Nothing
            pTopoOp = Nothing
            pListeSommets = Nothing
            pListeTriangles = Nothing
            pArea = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function qui permet de créer et retourner la polyligne des triangles de Delaunay.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour trouver les lignes des triangles de Delaunay.</param>
    ''' 
    ''' <returns>IPolyline contenant les lignes des triangles de Delaunay.</returns>
    ''' 
    Public Shared Function CreerPolyligneTrianglesDelaunay(ByVal pGeometry As IGeometry) As IPolyline
        'Déclarer les variables de travail
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant les lignes de généralisation de base.
        Dim pLimite As IPolyline = Nothing                  'Interface contenant la géométrie ou la limite du polygone.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans le Bag.
        Dim pPointColl As IPointCollection = Nothing        'Interface contenant les sommets d'un triangle.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour effectuer des opérations spatiales.
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pListeTriangles As List(Of SimpleTriangle) = Nothing 'Liste des triangles de Delaunay.

        'Définir la valeur par défaut
        CreerPolyligneTrianglesDelaunay = New Polyline
        CreerPolyligneTrianglesDelaunay.SpatialReference = pGeometry.SpatialReference

        Try
            'Vérifier si la géométrie n'est pas une ligne ou une surface
            If pGeometry.Dimension = esriGeometryDimension.esriGeometry0Dimension Then Exit Function

            'Créer la liste des sommets
            pListeSommets = CreerListeSommets(pGeometry)

            'Créer la liste des triangles de Delaunay pour créer le diagramme de Voronoi
            pListeTriangles = CreerListeTriangles(pGeometry.Envelope, pListeSommets)

            'Créer un nouveau Bag vide
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = pGeometry.SpatialReference

            'Interface pour ajouter des géométries dans le Bag
            pGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter tous les triangles
            For Each triangle As SimpleTriangle In pListeTriangles
                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point A
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.A).X, .Y = pListeSommets.Item(triangle.A).Y})
                'Ajouter le point B
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.B).X, .Y = pListeSommets.Item(triangle.B).Y})
                'Ajouter la ligne A-B
                pGeomColl.AddGeometry(pPolyline)

                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point B
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.B).X, .Y = pListeSommets.Item(triangle.B).Y})
                'Ajouter le point C
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.C).X, .Y = pListeSommets.Item(triangle.C).Y})
                'Ajouter la ligne B-C
                pGeomColl.AddGeometry(pPolyline)

                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point C
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.C).X, .Y = pListeSommets.Item(triangle.C).Y})
                'Ajouter le point A
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.A).X, .Y = pListeSommets.Item(triangle.A).Y})
                'Ajouter la ligne C-A
                pGeomColl.AddGeometry(pPolyline)
            Next

            'Interface pour construire les lignes des triangles
            pTopoOp = TryCast(CreerPolyligneTrianglesDelaunay, ITopologicalOperator2)
            'Construire l'union des lignes des triangles
            pTopoOp.ConstructUnion(CType(pGeomColl, IEnumGeometry))

            'Vérifier si la géométrie est une surface
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Conserver seulement les lignes à l'intérieur du polygone
                CreerPolyligneTrianglesDelaunay = CType(pTopoOp.Intersect(pGeometry, esriGeometryDimension.esriGeometry1Dimension), IPolyline)
                'Interface pour extraire la limite du polygone
                pTopoOp = TryCast(pGeometry, ITopologicalOperator2)
                'Définir la ligne de base
                pLimite = CType(pTopoOp.Boundary, IPolyline)

                'Si la géométrie est une ligne
            Else
                'Conserver seulement les lignes à l'intérieur de l'enveloppe
                CreerPolyligneTrianglesDelaunay = CType(pTopoOp.Intersect(pGeometry.Envelope, esriGeometryDimension.esriGeometry1Dimension), IPolyline)
                'Définir les lignes
                pLimite = CType(pGeometry, IPolyline)
            End If

            'Interface pour extraire la limite du polygone
            pTopoOp = TryCast(CreerPolyligneTrianglesDelaunay, ITopologicalOperator2)

            'Enlever les lignes à la limite de la géométrie
            CreerPolyligneTrianglesDelaunay = CType(pTopoOp.Difference(pLimite), IPolyline)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeometryBag = Nothing
            pPolyline = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
            pListeSommets = Nothing
            pListeTriangles = Nothing
            pLimite = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet de créer et retourner les lignes des triangles de Delaunay.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour trouver les lignes des triangles de Delaunay.</param>
    ''' <param name="pLignesTriangles">Interface contenant les lignes des triangles de Delaunay sans les limites des polygones ou sans la ligne.</param>
    ''' 
    ''' <returns>IPolyline contenant toutes les lignes des triangles de Delaunay</returns>
    Public Shared Function CreerPolyligneTrianglesDelaunay(ByVal pGeometry As IGeometry, ByRef pLignesTriangles As IPolyline) As IPolyline
        'Déclarer les variables de travail
        Dim pClone As IClone = Nothing                      'Interface pour cloner uen géométrie.
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant les lignes de généralisation de base.
        Dim pLimite As IPolyline = Nothing                  'Interface contenant la géométrie ou la limite du polygone.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans le Bag.
        Dim pPointColl As IPointCollection = Nothing        'Interface contenant les sommets d'un triangle.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour effectuer des opérations spatiales.
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pListeTriangles As List(Of SimpleTriangle) = Nothing 'Liste des triangles de Delaunay.

        'Définir la valeur par défaut
        CreerPolyligneTrianglesDelaunay = New Polyline
        CreerPolyligneTrianglesDelaunay.SpatialReference = pGeometry.SpatialReference

        'Définir la valeur par défaut
        pLignesTriangles = New Polyline
        pLignesTriangles.SpatialReference = pGeometry.SpatialReference

        Try
            'Vérifier si la géométrie n'est pas une ligne ou une surface
            If pGeometry.Dimension = esriGeometryDimension.esriGeometry0Dimension Then Exit Function

            'Créer la liste des sommets
            pListeSommets = CreerListeSommets(pGeometry)

            'Créer la liste des triangles de Delaunay pour créer le diagramme de Voronoi
            pListeTriangles = CreerListeTriangles(pGeometry.Envelope, pListeSommets)

            'Créer un nouveau Bag vide
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = pGeometry.SpatialReference

            'Interface pour ajouter des géométries dans le Bag
            pGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter tous les triangles
            For Each triangle As SimpleTriangle In pListeTriangles
                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point A
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.A).X, .Y = pListeSommets.Item(triangle.A).Y})
                'Ajouter le point B
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.B).X, .Y = pListeSommets.Item(triangle.B).Y})
                'Ajouter la ligne A-B
                pGeomColl.AddGeometry(pPolyline)

                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point B
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.B).X, .Y = pListeSommets.Item(triangle.B).Y})
                'Ajouter le point C
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.C).X, .Y = pListeSommets.Item(triangle.C).Y})
                'Ajouter la ligne B-C
                pGeomColl.AddGeometry(pPolyline)

                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point C
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.C).X, .Y = pListeSommets.Item(triangle.C).Y})
                'Ajouter le point A
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.A).X, .Y = pListeSommets.Item(triangle.A).Y})
                'Ajouter la ligne C-A
                pGeomColl.AddGeometry(pPolyline)
            Next

            'Interface pour construire les lignes des triangles
            pTopoOp = TryCast(pLignesTriangles, ITopologicalOperator2)
            'Construire l'union des lignes des triangles
            pTopoOp.ConstructUnion(CType(pGeometryBag, IEnumGeometry))

            'Interface pour cloner les lignes des triangles
            pClone = CType(pLignesTriangles, IClone)
            'Cloner les lignes des triangles
            CreerPolyligneTrianglesDelaunay = CType(pClone.Clone, IPolyline)

            'Vérifier si la géométrie est une surface
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour extraire la limite du polygone
                pTopoOp = TryCast(pGeometry, ITopologicalOperator2)
                'Définir la ligne de base
                pLimite = CType(pTopoOp.Boundary, IPolyline)

                'Interface pour construire les lignes des triangles
                pTopoOp = TryCast(pLignesTriangles, ITopologicalOperator2)
                'Enlever les lignes à la limite
                pLignesTriangles = CType(pTopoOp.Difference(pLimite), IPolyline)

                'Si la géométrie est une ligne
            Else
                'Enlever les lignes à la limite
                pLignesTriangles = CType(pTopoOp.Difference(pGeometry), IPolyline)
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pClone = Nothing
            pGeometryBag = Nothing
            pPolyline = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
            pListeSommets = Nothing
            pListeTriangles = Nothing
            pLimite = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet de créer et retourner les lignes des triangles de Delaunay qui sont à l'intérieure et l'extérieure de la géométrie.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour trouver les lignes des triangles de Delaunay.</param>
    ''' <param name="pLignesInt">Interface contenant les lignes des triangles de Delaunay qui sont à l'intérieure de la géométrie.</param>
    ''' <param name="pLignesExt">Interface contenant les lignes des triangles de Delaunay qui sont à l'extérieure de la géométrie.</param>
    ''' 
    Public Shared Sub CreerPolyligneTrianglesDelaunay(ByVal pGeometry As IGeometry, ByRef pLignesInt As IPolyline, ByRef pLignesExt As IPolyline)
        'Déclarer les variables de travail
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant les lignes de généralisation de base.
        Dim pLimite As IPolyline = Nothing                  'Interface contenant la géométrie ou la limite du polygone.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne.
        Dim pLignesTriangles As IPolyline = Nothing         'Interface contenant les lignes des triangles.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans le Bag.
        Dim pPointColl As IPointCollection = Nothing        'Interface contenant les sommets d'un triangle.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour effectuer des opérations spatiales.
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pListeTriangles As List(Of SimpleTriangle) = Nothing 'Liste des triangles de Delaunay.
        Dim pConvexHull As IPolygon = Nothing

        'Définir la valeur par défaut
        pLignesTriangles = New Polyline
        pLignesTriangles.SpatialReference = pGeometry.SpatialReference

        Try
            'Vérifier si la géométrie n'est pas une ligne ou une surface
            If pGeometry.Dimension = esriGeometryDimension.esriGeometry0Dimension Then Exit Sub

            'Interface pour extraire le ConvexHull de la géométrie
            pTopoOp = CType(pGeometry, ITopologicalOperator2)
            'Extraire le ConvexHull de la géométrie
            pConvexHull = CType(pTopoOp.ConvexHull, IPolygon)

            'Créer la liste des sommets
            pListeSommets = CreerListeSommets(pGeometry)

            'Créer la liste des triangles de Delaunay pour créer le diagramme de Voronoi
            pListeTriangles = CreerListeTriangles(pGeometry.Envelope, pListeSommets)

            'Créer un nouveau Bag vide
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = pGeometry.SpatialReference

            'Interface pour ajouter des géométries dans le Bag
            pGeomColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter tous les triangles
            For Each triangle As SimpleTriangle In pListeTriangles
                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point A
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.A).X, .Y = pListeSommets.Item(triangle.A).Y})
                'Ajouter le point B
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.B).X, .Y = pListeSommets.Item(triangle.B).Y})
                'Ajouter la ligne A-B
                pGeomColl.AddGeometry(pPolyline)

                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point B
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.B).X, .Y = pListeSommets.Item(triangle.B).Y})
                'Ajouter le point C
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.C).X, .Y = pListeSommets.Item(triangle.C).Y})
                'Ajouter la ligne B-C
                pGeomColl.AddGeometry(pPolyline)

                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pPolyline = New Polyline
                pPolyline.SpatialReference = pGeometry.SpatialReference
                'Interface pour ajouter des sommets
                pPointColl = CType(pPolyline, IPointCollection)
                'Ajouter le point C
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.C).X, .Y = pListeSommets.Item(triangle.C).Y})
                'Ajouter le point A
                pPointColl.AddPoint(New PointClass() With {.X = pListeSommets.Item(triangle.A).X, .Y = pListeSommets.Item(triangle.A).Y})
                'Ajouter la ligne C-A
                pGeomColl.AddGeometry(pPolyline)
            Next

            'Interface pour construire les lignes des triangles
            pTopoOp = TryCast(pLignesTriangles, ITopologicalOperator2)
            'Construire l'union des lignes des triangles
            pTopoOp.ConstructUnion(CType(pGeometryBag, IEnumGeometry))

            'Vérifier si la géométrie est une surface
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'Interface pour extraire la limite du polygone
                pTopoOp = TryCast(pGeometry, ITopologicalOperator2)
                'Définir la ligne de base
                pLimite = CType(pTopoOp.Boundary, IPolyline)

                'Interface pour construire les lignes des triangles
                pTopoOp = TryCast(pLignesTriangles, ITopologicalOperator2)
                'Enlever les lignes à la limite
                pTopoOp = CType(pTopoOp.Difference(pLimite), ITopologicalOperator2)

                'Conserver seulement les lignes à l'intérieur du polygone
                pLignesInt = CType(pTopoOp.Intersect(pGeometry, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

                'Conserver seulement les lignes à l'extérieur du polygone
                pLignesExt = CType(pTopoOp.Difference(pLignesInt), IPolyline)

                'Interface pour extraire les lignes à l'intérieures du ConvexHull de la géométrie
                pTopoOp = CType(pLignesExt, ITopologicalOperator2)
                'Extraire les lignes à l'intérieures du ConvexHull de la géométrie
                pLignesExt = CType(pTopoOp.Intersect(pConvexHull, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

                'Si la géométrie est une ligne
            Else
                'Enlever les lignes à la limite
                pTopoOp = CType(pTopoOp.Difference(pGeometry), ITopologicalOperator2)

                'Conserver seulement les lignes à l'intérieur du ConvexHull
                pLignesInt = CType(pTopoOp.Intersect(pConvexHull, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

                'Conserver seulement les lignes à l'extérieur de l'enveloppe
                pLignesExt = New Polyline
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeometryBag = Nothing
            pPolyline = Nothing
            pLignesTriangles = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
            pListeSommets = Nothing
            pListeTriangles = Nothing
            pLimite = Nothing
            pConvexHull = Nothing
        End Try
    End Sub
#End Region

#Region "Routines et fonctions privées"
    ''' <summary>
    ''' Function qui permet de créer et retourner la liste des sommets d'une géométrie.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour créer la liste des sommets.</param>
    ''' 
    ''' <returns>List(Of SimplePoint) contenant les sommets d'une géométrie, Nothing sinon.</returns>
    ''' 
    Protected Friend Shared Function CreerListeSommets(pGeometry As IGeometry) As List(Of SimplePoint)
        'Déclarer les variables de travail
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets d'une géométrie.
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pSimplePoint As SimplePoint = Nothing           'Sommet utilisé pour créer le diagramme de Voronoi.
        Dim iNbSommets As Integer = 0                       'Contient le nombre de sommets non dupliqués de la géométrie.

        'Définir la valeur par défaut
        CreerListeSommets = Nothing

        Try
            'Vérifier si la géométrie est valide
            If pGeometry Is Nothing Then
                'Retourner une erreur
                Throw New ArgumentException("ERREUR : La géométrie est invalide.")
            End If

            'Vérifier si la géométrie est un point
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPoint Then
                'Retourner une erreur
                Throw New ArgumentException("ERREUR : La géométrie ne doit pas être un point.")
            End If

            'Interface pour extraire les sommets de la géométrie
            pPointColl = TryCast(pGeometry, IPointCollection)
            'Vérifier si la géométrie possède au moins 3 sommets
            If pPointColl.PointCount < 3 Then
                'Retourner une erreur
                Throw New ArgumentException("ERREUR : La géométrie ne possède pas au moins 3 sommets.")
            End If

            'Initialiser la liste des sommets
            pListeSommets = New List(Of SimplePoint)

            'Traiter tous les sommets de la géométrie
            For i As Integer = 0 To pPointColl.PointCount - 1
                'Créer un nouveau sommet afin de pouvoir effectuer le tri
                pSimplePoint = New SimplePoint(pPointColl.Point(i).X, pPointColl.Point(i).Y)

                'Vérifier si le sommet est déjà présent 
                'car il ne doit pas avoir de sommets en double lors de la triangulation de Delaunay
                If Not pListeSommets.Contains(pSimplePoint) Then
                    'Ajouter le sommet dans la liste des sommets
                    pListeSommets.Add(pSimplePoint)
                End If
            Next

            'Vérifier si le nombre de sommets est inférieure à 3
            If pListeSommets.Count < 3 Then
                'Retourner une erreur
                Throw New ArgumentException("ERREUR : Les sommets non dupliqués de la géométrie ne possède pas au moins 3 sommets.")
            End If

            'Important: Trier la liste des sommets afin que les sommets se balayent de gauche à droite
            pListeSommets.Sort()

            'Définir la liste des sommets
            CreerListeSommets = pListeSommets

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPointColl = Nothing
            pListeSommets = Nothing
            pSimplePoint = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function qui permet de créer et retourner la liste des triangles de Delaunay d'une géométrie.
    ''' </summary>
    ''' 
    ''' <param name="pEnvelope">Enveloppe de la liste des sommets non dupliqués.</param>
    ''' <param name="pListeSommets">Liste des sommets non dupliqués pour créer la liste des triangles de Delaunay.</param>
    ''' 
    ''' <returns>List(Of SimpleTriangle) contenant les triangles de Delaunay d'une géométrie, Nothing sinon.</returns>
    ''' 
    Protected Friend Shared Function CreerListeTriangles(ByVal pEnvelope As IEnvelope, pListeSommets As IList(Of SimplePoint)) As List(Of SimpleTriangle)
        'Déclarer les variables de travail
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets d'une géométrie.
        Dim iNbSommets As Integer = 0                       'Contient le nombre de sommets non dupliqués de la géométrie.
        Dim dMax As Double = 0

        'Définir la valeur par défaut
        CreerListeTriangles = Nothing

        Try
            'Important: Conserver le nombre de sommets dans le tableau car certains points en double ont été supprimés
            iNbSommets = pListeSommets.Count

            'Définir la largeur ou hauteur maximum
            If pEnvelope.Width > pEnvelope.Height Then
                'Définir la largeur maximum
                dMax = pEnvelope.Width
            Else
                'Définir la hauteur maximum
                dMax = pEnvelope.Height
            End If

            'Définir la valeur X du centre de l'enveloppe
            Dim avgx As Double = ((pEnvelope.XMax - pEnvelope.XMin) / 2) + pEnvelope.XMin
            'Définir la valeur Y du centre de l'enveloppe
            Dim avgy As Double = ((pEnvelope.YMax - pEnvelope.YMin) / 2) + pEnvelope.YMin

            'Créer le point A du SuperTraingle de la liste des sommets
            Dim a As New SimplePoint(avgx - (2 * dMax), avgy - dMax)
            'Créer le point B du SuperTraingle de la liste des sommets
            Dim b As New SimplePoint(avgx + (2 * dMax), avgy - dMax)
            'Créer le point C du SuperTraingle de la liste des sommets
            Dim c As New SimplePoint(avgx, avgy + (2 * dMax))

            'Ajouter les sommets du SuperTriangle à la fin de la liste des sommets
            pListeSommets.Add(a)
            pListeSommets.Add(b)
            pListeSommets.Add(c)

            'Définir le centre du cercle et du rayon en utilisant les trois sommets du SuperTriangle
            Dim dRadius As Double
            Dim pSimplePointCentre As SimplePoint
            clsGeneraliserGeometrie.CalculateCircumcircle(a, b, c, pSimplePointCentre, dRadius)

            'Créer le Super triangle
            Dim pSuperTriangle As New SimpleTriangle(iNbSommets, iNbSommets + 1, iNbSommets + 2, pSimplePointCentre, dRadius)

            'Initialiser la liste des triangles temporaires
            Dim pListeTrianglesTemp As New List(Of SimpleTriangle)
            'Ajouter le super trianle à la liste des triangles
            pListeTrianglesTemp.Add(pSuperTriangle)

            'Initialiser la liste des triangles complets
            Dim pListeTrianglesComplets As New List(Of SimpleTriangle)

            'Traiter tous les sommets de la liste
            For i = 0 To iNbSommets - 1
                'Initialiser la liste des cotés de triangle
                Dim pListeCotes As New List(Of Integer())

                'Traiter tous les triangles
                For j As Integer = pListeTrianglesTemp.Count - 1 To 0 Step -1
                    ' Si le point se trouve dans la circonférence de ce triangle
                    If Distance(pListeTrianglesTemp(j).CircumCentre, pListeSommets(i)) < pListeTrianglesTemp(j).Radius Then
                        'Ajouter les côtés de triangle
                        pListeCotes.Add(New Integer() {pListeTrianglesTemp(j).A, pListeTrianglesTemp(j).B})
                        pListeCotes.Add(New Integer() {pListeTrianglesTemp(j).B, pListeTrianglesTemp(j).C})
                        pListeCotes.Add(New Integer() {pListeTrianglesTemp(j).C, pListeTrianglesTemp(j).A})

                        'Détruire le triangle de la liste
                        pListeTrianglesTemp.RemoveAt(j)

                        'Sinon
                    ElseIf pListeSommets(i).X > pListeTrianglesTemp(j).CircumCentre.X + pListeTrianglesTemp(j).Radius Then
                        'Vérifier si le triangle est complet
                        If True Then
                            'Ajouter le triangle complet
                            pListeTrianglesComplets.Add(pListeTrianglesTemp(j))
                        End If
                        'Détruire le triangle de la liste
                        pListeTrianglesTemp.RemoveAt(j)
                    End If
                Next

                'Traiter tous les côtés du début vers la fin
                For j As Integer = pListeCotes.Count - 1 To 1 Step -1
                    'Traiter dans l'autre sens
                    For k As Integer = j - 1 To 0 Step -1
                        'Comparez si ce le côté correspond dans l'une ou l'autre direction
                        If pListeCotes(j)(0).Equals(pListeCotes(k)(1)) AndAlso pListeCotes(j)(1).Equals(pListeCotes(k)(0)) Then
                            'Détruire les côtés en double
                            pListeCotes.RemoveAt(j)
                            pListeCotes.RemoveAt(k)

                            'Supprimer un élément de la liste inférieure à celle où j est maintenant, donc mise à jour j
                            j -= 1

                            'Sortir de la boucle
                            Exit For
                        End If
                    Next
                Next

                'Traiter tous les côtés
                For j = 0 To pListeCotes.Count - 1
                    'Définir le centre du cercle et du rayon en utilisant trois sommets
                    clsGeneraliserGeometrie.CalculateCircumcircle(pListeSommets(pListeCotes(j)(0)), pListeSommets(pListeCotes(j)(1)), pListeSommets(i), pSimplePointCentre, dRadius)
                    'Créer un nouveau triangle à partir du point actuel
                    Dim t As New SimpleTriangle(pListeCotes(j)(0), pListeCotes(j)(1), i, pSimplePointCentre, dRadius)
                    'Ajouter le nouveau triangle
                    pListeTrianglesTemp.Add(t)
                Next
            Next

            'Ajouter les triangles restants sur la liste complétée
            pListeTrianglesComplets.AddRange(pListeTrianglesTemp)

            'Définir la liste des triangles
            CreerListeTriangles = pListeTrianglesComplets

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPointColl = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet de créer et retourner les lignes du squelette de base d'un polygone dans une Polyline.
    ''' La fonction utilise les lignes des triangles de Delaunay pour créer le squelette de base.
    ''' Le squelette de base est créé à partir d'un squelette primaire.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Interface contenant le polygone utilisée pour trouver son squelette.</param>
    ''' <param name="pLignesTriangles">Interface contenant les lignes des triangles de Delaunay.</param>
    ''' <param name="pSquelette">Interface contenant les lignes du squelette du polygone.</param>
    ''' <param name="pBagDroites">Interface contenant le Bag des droites des triangles de Delaunay.</param>
    ''' 
    Protected Friend Shared Sub CreerSqueletteBaseDelaunay(ByVal pPolygon As IPolygon, ByVal pLignesTriangles As IPolyline, _
                                                           ByRef pSquelette As IPolyline, ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pBagDroitesTmp As IGeometryBag = Nothing            'Interface contenant le Bag des droites de Delaunay temporaire.
        Dim pBagSommetsPolygone As IGeometryBag = Nothing       'Interface contenant les sommets du polygone à traiter.
        Dim pBagSquelettePrimaire As IGeometryBag = Nothing     'Interface contenant les lignes du squelette primaire.
        Dim pBagSqueletteTmp As IGeometryBag = Nothing          'Interface contenant les lignes du squelette primaire temporaire.
        Dim pBagSqueletteBase As IGeometryBag = Nothing         'Interface contenant les lignes du squelette de base.
        Dim pBagSqueletteSimple As IGeometryBag = Nothing       'Interface contenant les lignes simples du squelette primaire.
        Dim pDictLiens As New Dictionary(Of Integer, Noeud)     'Dictionnaire contenant l'information des liens entre les sommets du polygone à traiter.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour effectuer des opérations spatiales.
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les anneaux du polygone.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface pour ajouter les lignes du squelettes primaire.
        Dim pPath As IPath = Nothing                            'Interface contenant un anneau du polygone.
        Dim pPointColl As IPointCollection = Nothing            'Interface pour extraire les sommets du polygone.

        Try
            'Créer le Bag des droites des triangles de Delaunay
            Call CreerBagLignesTriangles(pLignesTriangles, pBagDroites)

            'Créer le Bag du squelette promaire vide
            pBagSquelettePrimaire = New GeometryBag
            pBagSquelettePrimaire.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter les lignes du squelettes primaire
            pGeomCollAdd = CType(pBagSquelettePrimaire, IGeometryCollection)

            'Interface pour extraire les anneaux
            pGeomColl = CType(pPolygon, IGeometryCollection)

            'Traiter tous les anneaux du polygone
            For i = 0 To pGeomColl.GeometryCount - 1
                'Interface contenant un anneau du polygone 
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Créer le Bag des sommets de l'anneau et le dictionnaire des sommets contenant les sommets précédents et suivants de la ligne à traiter
                Call CreerBagSommetsLigne(pPath, pBagSommetsPolygone, pDictLiens)
                'Créer le dictionnaire contenant l'information des liens entre les sommets de la ligne
                Call CreerDictLiensSommetsLignes(pBagSommetsPolygone, pBagDroites, pDictLiens, pBagDroitesTmp)
                'Créer le Bag des lignes du squelette primaire
                Call CreerBagLignesSquelettePrimaire(pPath, pDictLiens, pBagSqueletteTmp, pBagDroitesTmp)

                'Ajouter les lignes du squelette primaire
                pGeomCollAdd.AddGeometryCollection(CType(pBagSqueletteTmp, IGeometryCollection))
            Next

            'Créer le Bag des lignes simples et doubles du squelette primaire
            'Le Bag des lignes doubles correspond aux lignes en double du squelette de base
            Call CreerBagLignesSimplesDoubles(pPolygon, pBagSquelettePrimaire, pBagSqueletteSimple, pBagSqueletteBase)

            'Traiter les lignes simples et les ajouter dans les lignes du squelette de base
            Call TraiterLignesSimples(pPolygon, pBagSqueletteSimple, pBagSqueletteBase)
            'pBagDroites = pBagSqueletteBase

            'Créer le squelette de base vide selon Delaunay
            pSquelette = New Polyline
            pSquelette.SpatialReference = pPolygon.SpatialReference

            'Interface pour construire le squelette de base selon Delaunay
            pTopoOp = TryCast(pSquelette, ITopologicalOperator2)
            'Construire le squelette de base selon Delaunay
            pTopoOp.ConstructUnion(CType(pBagSqueletteBase, IEnumGeometry))

            'Vérifier si le squelette est vide
            If pSquelette.IsEmpty Then
                'Interface pour extraire les sommets du polygone
                pPointColl = CType(pPolygon, IPointCollection)
                'Définir le squelette par défaut
                pSquelette.FromPoint = pPointColl.Point(0)
                pSquelette.ToPoint = pPointColl.Point(pPointColl.PointCount - 2)
            End If

            'pBagDroites = pBagSqueletteBase

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pBagDroitesTmp = Nothing
            pBagSommetsPolygone = Nothing
            pBagSquelettePrimaire = Nothing
            pBagSqueletteTmp = Nothing
            pBagSqueletteBase = Nothing
            pBagSqueletteSimple = Nothing
            pDictLiens = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPath = Nothing
            pPointColl = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le Bag et le dictionnaire des sommets d'une ligne ou d'un anneau à traiter.
    ''' Le dictionnaire des sommets contient les numéros de sommets précédents et suivants dans des noeuds de sommets.
    ''' </summary>
    ''' 
    ''' <param name="pPath">Interface contenant la ligne ou l'anneau à traiter.</param>
    ''' <param name="pBagSommetsLigne">Interface contenant le Bag des sommets de la ligne à traiter.</param>
    ''' <param name="pDictSommetsLigne">Dictionnaire contenant l'information de base pour chaque sommet de la ligne.</param>
    ''' 
    Protected Friend Shared Sub CreerBagSommetsLigne(ByVal pPath As IPath, ByRef pBagSommetsLigne As IGeometryBag, ByRef pDictSommetsLigne As Dictionary(Of Integer, Noeud))
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter des géométries dans le Bag.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pNoeudLien As Noeud = Nothing                   'Objet contenant l'information de base d'un sommet de ligne ou d'anneau.

        Try
            'Créer le GeometryBag vide des sommets
            pBagSommetsLigne = New GeometryBag
            pBagSommetsLigne.SpatialReference = pPath.SpatialReference
            'Interface pour ajouter les sommets de la ligne
            pGeomColl = CType(pBagSommetsLigne, IGeometryCollection)

            'Créer la liste vide des sommets précédents de la polyligne
            pDictSommetsLigne = New Dictionary(Of Integer, Noeud)

            'Interface pour extraire les sommets de la ligne
            pPointColl = CType(pPath, IPointCollection)

            'Traiter tous les sommets de la ligne
            For j = 0 To pPointColl.PointCount - 1
                'Créer le noeud pour le sommet de la ligne
                pNoeudLien = New Noeud(j)

                'Vérifier si c'est un début de ligne
                If j = 0 Then
                    'Définir le numéro du sommet suivant pour les débuts de ligne
                    pNoeudLien.NoSuiv = j + 1
                    'Définir l'angle maximum
                    pNoeudLien.AngleMax = Angle(pPointColl.Point(j), pPointColl.Point(pNoeudLien.NoSuiv))

                    'Vérifier si la ligne est fermée
                    If pPath.IsClosed Then
                        'Définir le numéro du sommet précédent pour les débuts de ligne
                        pNoeudLien.NoPrec = pPointColl.PointCount - 2
                        'Définir l'angle minimum
                        pNoeudLien.AngleMin = Angle(pPointColl.Point(j), pPointColl.Point(pNoeudLien.NoPrec))

                        'Vérifier si l'angle minimum est supérieure à l'angle maximum
                        If pNoeudLien.AngleMin > pNoeudLien.AngleMax Then
                            'Ajouter 360 à l'angle maximum
                            pNoeudLien.AngleMax = pNoeudLien.AngleMax + 360
                        End If

                        'Si la ligne n'est pas fermée
                    Else
                        'Définir le numéro du sommet précédent pour les débuts de ligne
                        pNoeudLien.NoPrec = -1

                        'Vérifier si l'angle maximum est inférieure à 180
                        If pNoeudLien.AngleMax < 180 Then
                            'Définir l'angle maximum
                            pNoeudLien.AngleMin = pNoeudLien.AngleMax + 180
                            pNoeudLien.AngleMax = pNoeudLien.AngleMin + 180
                        End If

                        'Définir l'angle minimum
                        pNoeudLien.AngleMin = pNoeudLien.AngleMax - 180
                    End If

                    'Vérifier si c'est une fin de ligne
                ElseIf j = pPointColl.PointCount - 1 Then
                    'Définir le numéro du sommet précédent pour les fins de ligne
                    pNoeudLien.NoPrec = j - 1
                    'Définir l'angle minimum
                    pNoeudLien.AngleMin = Angle(pPointColl.Point(j), pPointColl.Point(pNoeudLien.NoPrec))

                    'Vérifier si la ligne est fermée
                    If pPath.IsClosed Then
                        'Définir le numéro du sommet suivant pour les fins de ligne
                        pNoeudLien.NoSuiv = 1
                        'Définir l'angle maximum
                        pNoeudLien.AngleMax = Angle(pPointColl.Point(j), pPointColl.Point(pNoeudLien.NoSuiv), pNoeudLien.AngleMin)

                        'Si la ligne n'est pas fermée
                    Else
                        'Définir le numéro du sommet suivant pour les fins de ligne
                        pNoeudLien.NoSuiv = -1
                        'Définir l'angle maximum
                        pNoeudLien.AngleMax = pNoeudLien.AngleMin + 180
                    End If

                    'Si ce n'est pas un début ou une fin de ligne
                Else
                    'Définir le numéro du sommet précédent
                    pNoeudLien.NoPrec = j - 1
                    'Définir l'angle minimum
                    pNoeudLien.AngleMin = Angle(pPointColl.Point(j), pPointColl.Point(pNoeudLien.NoPrec))

                    'Définir le numéro du sommet suivant
                    pNoeudLien.NoSuiv = j + 1
                    'Définir l'angle maximum
                    pNoeudLien.AngleMax = Angle(pPointColl.Point(j), pPointColl.Point(pNoeudLien.NoSuiv), pNoeudLien.AngleMin)
                End If

                'Ajouter le sommet de la ligne dans le Bag
                pGeomColl.AddGeometry(pPointColl.Point(j))

                'Ajouter l'information du sommet/noeud dans le dictionnaire
                pDictSommetsLigne.Add(j, pNoeudLien)
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPointColl = Nothing
            pNoeudLien = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le Bag des lignes des triangles de Delaunay.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Interface contenant les lignes des triangles de Delaunay.</param>
    ''' <param name="pBagLignesTriangles">Interface contenant les lignes des triangles de Delaunay.</param>
    ''' 
    Protected Friend Shared Sub CreerBagLignesTriangles(ByVal pPolyline As IPolyline, ByRef pBagLignesTriangles As IGeometryBag)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes dans le Bag.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segment des lignes.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les sommets d'une géométrie.

        Try
            'Créer le GeometryBag vide des lignes
            pBagLignesTriangles = New GeometryBag
            pBagLignesTriangles.SpatialReference = pPolyline.SpatialReference

            'Interface pour ajouter les lignes
            pGeomColl = CType(pBagLignesTriangles, IGeometryCollection)

            'Interface pour extraire les lignes
            pSegColl = CType(pPolyline, ISegmentCollection)

            'Traiter toutes les lignes
            For i = 0 To pSegColl.SegmentCount - 1
                'Créer une ligne
                pPolyline = New Polyline
                pPolyline.SpatialReference = pPolyline.SpatialReference

                'Interface pour ajouter les sommets de la ligne
                pPointColl = CType(pPolyline, IPointCollection)

                'Ajouter les sommets de la ligne
                pPointColl.AddPoint(pSegColl.Segment(i).FromPoint)
                pPointColl.AddPoint(pSegColl.Segment(i).ToPoint)

                'Ajouter la ligne dans le Bag
                pGeomColl.AddGeometry(pPolyline)
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pSegColl = Nothing
            pPointColl = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le dictionnaire contenant l'information des liens entre les sommets d'une ligne ou d'un anneau et les lignes de Delaunay.
    ''' </summary>
    ''' 
    ''' <param name="pBagSommetsLigne">Interface contenant les sommets de la ligne à traiter.</param>
    ''' <param name="pBagLignesTriangles">Interface contenant les lignes des triangles de Delaunay.</param>
    ''' <param name="pDictLiens">Dictionnaire contenant l'information des sommets de la ligne à traiter.</param>
    ''' <param name="pBagDroites">Interface contenant les lignes des triangles de Delaunay utilisées.</param>
    ''' 
    Protected Friend Shared Sub CreerDictLiensSommetsLignes(ByVal pBagSommetsLigne As IGeometryBag, ByVal pBagLignesTriangles As IGeometryBag, _
                                                            ByRef pDictLiens As Dictionary(Of Integer, Noeud), ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pTest As New List(Of Integer)                   'Liste des numéros de ligne traitée.
        Dim pNoeudLien As Noeud = Nothing                   'Objet contenant l'information d'un sommet de la ligne à traiter.
        Dim pLigne As IPolyline = Nothing                   'Interface contenant la ligne de Delaunay à traiter.
        Dim pPointColl As IGeometryCollection = Nothing     'Interface utilisé pour extraire les points dans le Bag.
        Dim pPointA As IPoint = Nothing                     'Interface contenant le premier sommet d'une droite.
        Dim pPointB As IPoint = Nothing                     'Interface contenant le deuxième sommet d'une droite.
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface pour extraire les lignes dans le Bag.
        Dim pLigneCollAdd As IGeometryCollection = Nothing  'Interface pour ajouter les lignes dans le Bag.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter une relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement d'une relation spatiale.
        Dim iSel As Integer = -1                            'Numéro de sommet.
        Dim iRel As Integer = -1                            'Numéro de la ligne.
        Dim dAngleDroite As Double = 0                      'Contient l'angle de la droite.

        Try
            'Créer le GeometryBag vide des droites
            pBagDroites = New GeometryBag
            pBagDroites.SpatialReference = pBagLignesTriangles.SpatialReference
            'Interface pour ajouter les lignes dans le Bag
            pLigneCollAdd = CType(pBagDroites, IGeometryCollection)
            'Interface pour extraire les lignes dans le Bag
            pLigneColl = CType(pBagLignesTriangles, IGeometryCollection)

            'Interface pour extraire les points de la ligne
            pPointColl = CType(pBagSommetsLigne, IGeometryCollection)

            'Interface pour traiter les relations spatiales entre les sommets de la ligne et les lignes des triangles
            pRelOpNxM = CType(pBagSommetsLigne, IRelationalOperatorNxM)

            'Traiter la relation spatiale entre les sommets du polygone et les lignes des triangles
            pRelResult = pRelOpNxM.Intersects(CType(pBagLignesTriangles, IGeometryBag))

            'Traiter toutes les relations Sommets-Lignes
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Définir le point de début la ligne
                pPointA = CType(pPointColl.Geometry(iSel), IPoint)

                'Définir la ligne de Delaunay
                pLigne = CType(pLigneColl.Geometry(iRel), IPolyline)

                'Définir le noeud du sommet de la ligne
                pNoeudLien = pDictLiens.Item(iSel)

                'Vérifier si le point est égal au premier point de la ligne
                If pPointA.Compare(pLigne.FromPoint) = 0 Then
                    'Définir le point de fin de ligne
                    pPointB = pLigne.ToPoint

                    'Si le point est égal au dernier point de la ligne
                Else
                    'Définir le point de fin de ligne
                    pPointB = pLigne.FromPoint
                End If

                'Définir l'angle de la ligne
                dAngleDroite = Angle(pPointA, pPointB, pNoeudLien.AngleMin)
                'Vérifier si l'angle de la droite est valide
                If dAngleDroite < pNoeudLien.AngleMax Then
                    'Debug.Print(iSel.ToString & "-" & pPointA.X.ToString & "," & pPointA.Y.ToString & ": " & dAngleDroite.ToString & ": " & pNoeudLien.AngleMin.ToString & ": " & pNoeudLien.AngleMax.ToString)
                    'Ajouter le numéro de la droite, le point de début et fin de la droite, l'angle de la droite et si elle est traitée
                    pNoeudLien.Add(iRel, pPointA, pPointB, dAngleDroite, pTest.Contains(iRel))

                    'Ajouter la ligne au Bag
                    If pTest.Contains(iRel) = False Then
                        'Ajouter la ligne utilisée
                        pLigneCollAdd.AddGeometry(pLigneColl.Geometry(iRel))

                        'Indiquer que la ligne a été traitée
                        pTest.Add(iRel)
                    End If
                End If

                'Mettre à jour le noeud du sommet de la ligne dans le dictionnaire
                pDictLiens.Item(iSel) = pNoeudLien
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTest = Nothing
            pNoeudLien = Nothing
            pLigne = Nothing
            pPointA = Nothing
            pPointB = Nothing
            pPointColl = Nothing
            pLigneColl = Nothing
            pLigneCollAdd = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le Bag des lignes du squelette primaire d'une ligne ou d'un anneau.
    ''' </summary>
    ''' 
    ''' <param name="pPath">Interface contenant la ligne ou d'un anneau à traiter.</param>
    ''' <param name="pDictLiens">Dictionnaire contenant l'information des liens entre les sommets de la géométrie et les droites de Delaunay.</param>
    ''' <param name="pBagLignesPrimaire">Interface contenant les lignes du squelette primaire de la polyligne.</param>
    ''' <param name="pBagDroites">Interface contenant les droites de Delaunay utilisées.</param>
    ''' 
    Protected Friend Shared Sub CreerBagLignesSquelettePrimaire(ByVal pPath As IPath, ByVal pDictLiens As Dictionary(Of Integer, Noeud), _
                                                                ByRef pBagLignesPrimaire As IGeometryBag, ByRef pBagDroites As IGeometryBag)
        'Déclarer les variables de travail
        Dim pListeDroites As New List(Of Droite)            'Liste des droites ordonnées.
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface pour ajouter des lignes.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour ajouter les lignes dans le Bag.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour ajouter les lignes dans le Bag.
        Dim pLigne As IPolyline = Nothing                   'Interface contenant une ligne.
        Dim pPointColl As IPointCollection = Nothing        'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pPointCollAdd As IPointCollection = Nothing     'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pPoint As IPoint = Nothing                      'Interface contenant le point du centre de la droite.
        Dim pPointA As IPoint = Nothing                     'Interface contenant le premier sommet d'une droite.
        Dim pPointB As IPoint = Nothing                     'Interface contenant le deuxième sommet d'une droite.
        Dim pDroiteDeb As Droite = Nothing                  'Objet contenant la droite de début.
        Dim pDroite As Droite = Nothing                     'Objet contenant la droite à traiter.
        Dim pNoeudLien As Noeud = Nothing                   'Objet contenant l'information d'un sommet à traiter.
        Dim iNbDroites As Integer = 0       'Compteur des droites.
        Dim iNbPoints As Integer = 0        'Compteur de points.

        Try
            'Créer le GeometryBag vide des lignes
            pBagDroites = New GeometryBag
            pBagDroites.SpatialReference = pPath.SpatialReference
            'Interface pour ajouter les lignes du squelette primaire
            pGeomCollAdd = CType(pBagDroites, IGeometryCollection)

            'Créer le GeometryBag vide des lignes
            pBagLignesPrimaire = New GeometryBag
            pBagLignesPrimaire.SpatialReference = pPath.SpatialReference
            'Interface pour ajouter les lignes du squelette primaire
            pGeomColl = CType(pBagLignesPrimaire, IGeometryCollection)

            'Interface pour extraire les sommets de la polyligne
            pPointColl = CType(pPath, IPointCollection)

            'Vérifier si la ligne est fermée
            If pPath.IsClosed Then
                'Définir le nombre de points à traiter
                iNbPoints = pPointColl.PointCount - 2

                'Si la ligne est pas fermée
            Else
                'Définir le nombre de points à traiter
                iNbPoints = pPointColl.PointCount - 1
            End If

            '---------------------------------------------------
            'Créer la liste des droites ordonnées
            '---------------------------------------------------
            'Traiter tous les sommets
            For i = 0 To iNbPoints
                'Vérifier si le sommet possède un lien
                If pDictLiens.ContainsKey(i) Then
                    'Définir le noeud du lien entre les sommets
                    pNoeudLien = pDictLiens.Item(i)

                    'Définir le nombre de droites
                    iNbDroites = pNoeudLien.Droites.Count

                    'Si le sommet ne possède pas un lien
                Else
                    'Définir le nombre de droites
                    iNbDroites = 0
                End If

                'Vérifier la présence des droites pour le sommet
                If iNbDroites > 0 Then
                    'Traiter tous les sommets des liens
                    For j = 0 To pNoeudLien.Droites.Count - 1
                        'Définir le numéro de sommet en lien
                        pDroite = pNoeudLien.Droites.Item(j)
                        'Ajouter la droite dans la liste
                        pListeDroites.Add(pDroite)

                        'Créer une ligne
                        pLigne = New Polyline
                        pLigne.SpatialReference = pPath.SpatialReference
                        'Interface pour ajouter les sommets de la ligne
                        pPointCollAdd = CType(pLigne, IPointCollection)
                        pPointCollAdd.AddPoint(pDroite.PointDeb)
                        pPointCollAdd.AddPoint(pDroite.PointFin)
                        'Ajouter la droite dans le Bag
                        pGeomCollAdd.AddGeometry(pLigne)
                    Next

                    'Si le sommet ne possède pas de droite
                Else
                    'Créer une nouvelle droite avec le même sommet
                    pDroite = New Droite
                    pDroite.PointDeb = pPointColl.Point(i)
                    pDroite.PointFin = pPointColl.Point(i)
                    pDroite.Angle = -1
                    pDroite.No = -1
                    'Ajouter la droite dans la liste
                    pListeDroites.Add(pDroite)

                    'Créer une ligne
                    pLigne = New Polyline
                    pLigne.SpatialReference = pPath.SpatialReference
                    'Interface pour ajouter les sommets de la ligne
                    pPointCollAdd = CType(pLigne, IPointCollection)
                    pPointCollAdd.AddPoint(pDroite.PointDeb)
                    pPointCollAdd.AddPoint(pDroite.PointFin)
                    'Ajouter la droite dans le Bag
                    pGeomCollAdd.AddGeometry(pLigne)
                End If
            Next

            'Vérifier si la ligne est fermée
            If pPath.IsClosed Then
                'Définir la première droite 
                pDroite = pListeDroites.Item(0)

                'Ajouter la droite dans la liste
                pListeDroites.Add(pDroite)
            End If

            '---------------------------------------------------
            'Créer le squelette primaire
            '---------------------------------------------------
            'Traiter tous les liens entre sommets
            For i = 1 To pListeDroites.Count - 1
                'Définir la droite à traiter
                pDroiteDeb = pListeDroites.Item(i - 1)
                'Définir la droite à traiter
                pDroite = pListeDroites.Item(i)

                'Créer une ligne
                pLigne = New Polyline
                pLigne.SpatialReference = pPath.SpatialReference
                'Interface pour ajouter les sommets de la ligne
                pPointCollAdd = CType(pLigne, IPointCollection)

                'Définir le point de début de la droite
                pPointA = pDroiteDeb.PointDeb
                'Définir le point de fin de la droite
                pPointB = pDroiteDeb.PointFin
                'Définir le centre de la droite
                pPoint = New Point
                pPoint.SpatialReference = pPath.SpatialReference
                pPoint.X = (pPointA.X + pPointB.X) / 2
                pPoint.Y = (pPointA.Y + pPointB.Y) / 2
                'Ajouter le premier sommet de la droite
                pPointCollAdd.AddPoint(pPoint)

                'Définir le point de début de la droite
                pPointA = pDroite.PointDeb
                'Définir le point de fin de la droite
                pPointB = pDroite.PointFin
                'Définir le centre de la droite
                pPoint = New Point
                pPoint.SpatialReference = pPath.SpatialReference
                pPoint.X = (pPointA.X + pPointB.X) / 2
                pPoint.Y = (pPointA.Y + pPointB.Y) / 2
                'Ajouter le deuxième sommet de la droite
                pPointCollAdd.AddPoint(pPoint)

                'Si la longueur n'est pas zéro et si c'est un sommet précédent 
                If pLigne.Length > 0 Then
                    'Ajouter la ligne dans le bag
                    pGeomColl.AddGeometry(CType(pLigne, IGeometry))
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pListeDroites = Nothing
            pLigneColl = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pLigne = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pPoint = Nothing
            pPointA = Nothing
            pPointB = Nothing
            pDroiteDeb = Nothing
            pDroite = Nothing
            pNoeudLien = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le Bag des lignes de jonction (non doublées) du squelette primaire d'un polygone.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Interface contenant le polygone à traiter.</param>
    ''' <param name="pBagLignesPrimaire">Interface contenant les lignes du squelette primaire.</param>
    ''' <param name="pBagLignesSimples">Interface contenant les lignes simples du squelette primaire.</param>
    ''' <param name="pBagLignesDoubles">Interface contenant les lignes doubles du squelette de base.</param>
    ''' 
    Protected Friend Shared Sub CreerBagLignesSimplesDoubles(ByVal pPolygon As IPolygon, pBagLignesPrimaire As IGeometryBag, _
                                                             ByRef pBagLignesSimples As IGeometryBag, ByRef pBagLignesDoubles As IGeometryBag)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les lignes du squelette primaire.
        Dim pBagPolygone As IGeometryBag = Nothing              'Interface contenant le polygone à traiter.
        Dim pGeomCollSimple As IGeometryCollection = Nothing    'Interface pour ajouter les lignes simples dans le Bag.
        Dim pGeomCollDouble As IGeometryCollection = Nothing    'Interface pour ajouter les lignes doubles dans le Bag.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing       'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing             'Interface contenant le résultat du traitement de la relation spatiale.
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.
        Dim iDeb As Integer = -2            'Dernier numéro de début de droite traité.
        Dim iNo As Integer = -2             'Contient le nombre de relation.

        Try
            'Créer le GeometryBag vide des lignes simples
            pBagLignesSimples = New GeometryBag
            pBagLignesSimples.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter les lignes simples
            pGeomCollSimple = CType(pBagLignesSimples, IGeometryCollection)

            'Créer le GeometryBag vide des lignes doubles
            pBagLignesDoubles = New GeometryBag
            pBagLignesDoubles.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter les lignes doubles
            pGeomCollDouble = CType(pBagLignesDoubles, IGeometryCollection)

            'Interface pour extraire les lignes du squelette primaire
            pGeomColl = CType(pBagLignesPrimaire, IGeometryCollection)

            'Interface pour traiter les relations spatiales
            pRelOpNxM = CType(pBagLignesPrimaire, IRelationalOperatorNxM)
            'Traiter la relation spatiale entre les lignes
            pRelResult = pRelOpNxM.Contains(CType(pBagLignesPrimaire, IGeometryBag))
            'Trier le résultat
            pRelResult.SortRight()

            'Initialiser le nombrede relation de la droite
            iNo = 0
            'Initialiser le no de la droite de début
            iDeb = -1

            'Traiter toutes les relations
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iRel, iSel)
                'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                'Vérifier si c'est le même numéro
                If iSel = iDeb Then
                    'Compter le nombre de relation
                    iNo = iNo + 1

                    'Si ce ne sont pas les mêmes numéros
                Else
                    'Vérifier si une seule relation est présente
                    If iNo = 1 Then
                        'Ajouter la droite simple
                        pGeomCollSimple.AddGeometry(pGeomColl.Geometry(iDeb))

                        'Si plusieurs relations sont présentes
                    ElseIf iNo > 0 Then
                        'Ajouter la droite double
                        pGeomCollDouble.AddGeometry(pGeomColl.Geometry(iDeb))
                    End If

                    'Définir l'ancien numéro traité
                    iDeb = iSel

                    'Initialiser le nombre de relation
                    iNo = 1
                End If
            Next

            'Créer le GeometryBag vide du polygone à traiter
            pBagPolygone = New GeometryBag
            pBagPolygone.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter le polygone dans le Bag
            pGeomColl = CType(pBagPolygone, IGeometryCollection)
            'Ajouter le polygone dans le Bag
            pGeomColl.AddGeometry(pPolygon)

            'Interface pour extraire les lignes simples
            pGeomColl = CType(pBagLignesSimples, IGeometryCollection)

            'Créer le GeometryBag vide des lignes simples
            pBagLignesSimples = New GeometryBag
            pBagLignesSimples.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter les lignes simples
            pGeomCollSimple = CType(pBagLignesSimples, IGeometryCollection)

            'Interface pour traiter les relations spatiales
            pRelOpNxM = CType(pGeomColl, IRelationalOperatorNxM)
            'Traiter la relation spatiale entre les lignes
            pRelResult = pRelOpNxM.Within(CType(pBagPolygone, IGeometryBag))

            'Vérifier si toutes les lignes simple sont à l'intérieure du polygone 
            If pRelResult.RelationElementCount = pGeomColl.GeometryCount Then
                'Définiur le Bag des lignes simples originales
                pBagLignesSimples = CType(pGeomColl, IGeometryBag)

                'S'il y a des lignes simple qui ne sont pas à l'intérieure du polygone 
            Else
                'Traiter toutes les relations
                For i = 0 To pRelResult.RelationElementCount - 1
                    'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                    pRelResult.RelationElement(i, iSel, iRel)
                    'Ajouter la droite simple
                    pGeomCollSimple.AddGeometry(pGeomColl.Geometry(iSel))
                Next
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pBagPolygone = Nothing
            pGeomColl = Nothing
            pGeomCollSimple = Nothing
            pGeomCollDouble = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de créer et retourner le Bag des lignes de jonction (non doublées) du squelette primaire d'une polyligne.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Interface contenant la polyligne à traiter.</param>
    ''' <param name="pBagLignesPrimaire">Interface contenant les lignes du squelette primaire.</param>
    ''' <param name="pBagLignesSimples">Interface contenant les lignes simples du squelette primaire.</param>
    ''' <param name="pBagLignesDoubles">Interface contenant les lignes doubles du squelette de base.</param>
    ''' 
    Protected Friend Shared Sub CreerBagLignesSimplesDoubles(ByVal pPolyline As IPolyline, pBagLignesPrimaire As IGeometryBag, _
                                                             ByRef pBagLignesSimples As IGeometryBag, ByRef pBagLignesDoubles As IGeometryBag)
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les lignes du squelette primaire.
        Dim pGeomCollSimple As IGeometryCollection = Nothing    'Interface pour ajouter les lignes simples dans le Bag.
        Dim pGeomCollDouble As IGeometryCollection = Nothing    'Interface pour ajouter les lignes doubles dans le Bag.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing       'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing             'Interface contenant le résultat du traitement de la relation spatiale.
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.
        Dim iDeb As Integer = -2            'Dernier numéro de début de droite traité.
        Dim iNo As Integer = -2             'Contient le nombre de relation.

        Try
            'Créer le GeometryBag vide des lignes simples
            pBagLignesSimples = New GeometryBag
            pBagLignesSimples.SpatialReference = pPolyline.SpatialReference
            'Interface pour ajouter les lignes simples
            pGeomCollSimple = CType(pBagLignesSimples, IGeometryCollection)

            'Créer le GeometryBag vide des lignes doubles
            pBagLignesDoubles = New GeometryBag
            pBagLignesDoubles.SpatialReference = pPolyline.SpatialReference
            'Interface pour ajouter les lignes doubles
            pGeomCollDouble = CType(pBagLignesDoubles, IGeometryCollection)

            'Interface pour extraire les lignes du squelette primaire
            pGeomColl = CType(pBagLignesPrimaire, IGeometryCollection)

            'Interface pour traiter les relations spatiales
            pRelOpNxM = CType(pBagLignesPrimaire, IRelationalOperatorNxM)
            'Traiter la relation spatiale entre les lignes
            pRelResult = pRelOpNxM.Contains(CType(pBagLignesPrimaire, IGeometryBag))
            'Trier le résultat
            pRelResult.SortRight()

            'Initialiser le nombrede relation de la droite
            iNo = 0
            'Initialiser le no de la droite de début
            iDeb = -1

            'Traiter toutes les relations
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iRel, iSel)
                'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                'Vérifier si c'est le même numéro
                If iSel = iDeb Then
                    'Compter le nombre de relation
                    iNo = iNo + 1

                    'Si ce ne sont pas les mêmes numéros
                Else
                    'Vérifier si une seule relation est présente
                    If iNo = 1 Then
                        'Ajouter la droite simple
                        pGeomCollSimple.AddGeometry(pGeomColl.Geometry(iDeb))

                        'Si plusieurs relations sont présentes
                    ElseIf iNo > 0 Then
                        'Ajouter la droite double
                        pGeomCollDouble.AddGeometry(pGeomColl.Geometry(iDeb))
                    End If

                    'Définir l'ancien numéro traité
                    iDeb = iSel

                    'Initialiser le nombre de relation
                    iNo = 1
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pGeomCollSimple = Nothing
            pGeomCollDouble = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de traiter les lignes simples et les ajouter dans les lignes du squelette de base d'un polygone.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Interface contenant le polygone à traiter.</param>
    ''' <param name="pBagLignesSimples">Interface contenant les lignes simples du squelette primaire.</param>
    ''' <param name="pBagLignesBase">Interface contenant les lignes en double du squelette de base.</param>
    ''' 
    Protected Friend Shared Sub TraiterLignesSimples(ByVal pPolygon As IPolygon, ByVal pBagLignesSimples As IGeometryBag, ByRef pBagLignesBase As IGeometryBag)
        'Déclarer les variables de travail
        Dim pTriangle As IPolygon = Nothing                     'Interface contenant un triangle.
        Dim pListeLignes As List(Of Integer)                    'Liste des numéros de ligne traitées.
        Dim pLigneColl As IGeometryCollection = Nothing         'Interface pour ajouter des lignes.
        Dim pGeomCollSimple As IGeometryCollection = Nothing    'Interface pour extraire les lignes simples dans le Bag.
        Dim pPolyline As IPolyline = Nothing                    'Interface contenant une ligne.
        Dim pPointColl As IPointCollection = Nothing            'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour effectuer des opérations spatiales.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing       'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing             'Interface contenant le résultat du traitement de la relation spatiale.
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.
        Dim iNo As Integer = -2             'Contient le nombre de relation.
        Dim iDeb As Integer = -2            'Dernier numéro de ligne traitée.

        Try
            'Interface pour extraire les lignes simples
            pGeomCollSimple = CType(pBagLignesSimples, IGeometryCollection)

            'Traiter tant qu'il reste des lignes
            Do
                'Initialiser la liste des numéros de lignes traitées
                pListeLignes = New List(Of Integer)

                'Interface pour traiter les relations spatiales des lignes simples
                pRelOpNxM = CType(pBagLignesSimples, IRelationalOperatorNxM)
                'Traiter la relation spatiale entre les lignes simples
                pRelResult = pRelOpNxM.Touches(CType(pBagLignesSimples, IGeometryBag))
                'Trier le résultat
                pRelResult.SortRight()

                'Initialiser le nombre de connexion
                iNo = 0
                'Initialiser le numéro de ligne précédent
                iDeb = -1
                Dim i As Integer
                'Traiter toutes les relations
                For i = 0 To pRelResult.RelationElementCount - 1
                    'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                    pRelResult.RelationElement(i, iRel, iSel)
                    'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                    'Vérifier si c'est le même numéro
                    If iSel = iDeb Then
                        'Compter le nombre de relation
                        iNo = iNo + 1

                        'Si ce ne sont pas les mêmes numéros
                    Else
                        'Vérifier si une ou deux relation sont présentes
                        If iNo = 1 Or iNo = 2 Then
                            'Créer un triagle vide
                            pTriangle = New Polygon
                            pTriangle.SpatialReference = pPolygon.SpatialReference
                            'Interface pour extraire les sommets du triangle
                            pPointColl = CType(pTriangle, IPointCollection)

                            'Extraire le résultat de la relation précédente
                            pRelResult.RelationElement(i - 1, iRel, iSel)
                            'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                            'Définir la ligne traitée
                            pPolyline = CType(pGeomCollSimple.Geometry(iSel), IPolyline)
                            'Ajouter les sommets de la ligne dans le triangle 
                            pPointColl.AddPoint(pPolyline.FromPoint)
                            pPointColl.AddPoint(pPolyline.ToPoint)

                            'Définir la ligne en relation
                            pPolyline = CType(pGeomCollSimple.Geometry(iRel), IPolyline)
                            'Ajouter les sommets de la ligne dans le triangle 
                            pPointColl.AddPoint(pPolyline.FromPoint)
                            pPointColl.AddPoint(pPolyline.ToPoint)

                            'Définir que la ligne traité a été traitée
                            If Not pListeLignes.Contains(iSel) Then pListeLignes.Add(iSel)
                            'Définir que la ligne en relation a été traitée
                            If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)

                            'Si le nombre de relation est 2
                            If iNo = 2 Then
                                'Extraire la deuxième relation
                                pRelResult.RelationElement(i - 2, iRel, iSel)
                                'Définir que la ligne en relation a été traitée
                                If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)
                            End If

                            'Revenir à la relation courante
                            pRelResult.RelationElement(i, iRel, iSel)

                            'Fermer le triangle
                            pTriangle.Close()
                            'Simplifier le triangle
                            pTopoOp = CType(pTriangle, ITopologicalOperator2)
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()

                            'Ajouter les 2 plus petites longueur du triangle dans les lignes du squelette de base
                            Call TraiterLongueurTriangle(pTriangle, pBagLignesBase)
                        End If

                        'Initialiser le dernier numéro traité
                        iDeb = iSel
                        'Initialiser le nombre de relation
                        iNo = 1
                    End If
                Next

                'Vérifier si une ou deux relation sont présentes
                If iNo = 1 Or iNo = 2 Then
                    'Créer un triagle vide
                    pTriangle = New Polygon
                    pTriangle.SpatialReference = pPolygon.SpatialReference
                    'Interface pour extraire les sommets du triangle
                    pPointColl = CType(pTriangle, IPointCollection)

                    'Extraire le résultat de la relation précédente
                    pRelResult.RelationElement(pRelResult.RelationElementCount - 1, iRel, iSel)
                    'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                    'Définir la ligne traitée
                    pPolyline = CType(pGeomCollSimple.Geometry(iSel), IPolyline)
                    'Ajouter les sommets de la ligne dans le triangle
                    pPointColl.AddPoint(pPolyline.FromPoint)
                    pPointColl.AddPoint(pPolyline.ToPoint)

                    'Définir la ligne en relation
                    pPolyline = CType(pGeomCollSimple.Geometry(iRel), IPolyline)
                    'Ajouter les sommets de la ligne dans le triangle
                    pPointColl.AddPoint(pPolyline.FromPoint)
                    pPointColl.AddPoint(pPolyline.ToPoint)

                    'Définir que la ligne traité a été traitée
                    If Not pListeLignes.Contains(iSel) Then pListeLignes.Add(iSel)
                    'Définir que la ligne en relation a été traitée
                    If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)

                    'Si le nombre de relation est 2
                    If iNo = 2 Then
                        'Extraire la deuxième relation
                        pRelResult.RelationElement(pRelResult.RelationElementCount - 2, iRel, iSel)
                        'Définir que la ligne en relation a été traitée
                        If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)
                    End If

                    'Fermer le triangle
                    pTriangle.Close()
                    'Interface pour simplifier le triangle
                    pTopoOp = CType(pTriangle, ITopologicalOperator2)
                    'Simplifier le triangle
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()

                    'Ajouter les 2 plus petites longueur du triangle dans les lignes du squelette de base
                    Call TraiterLongueurTriangle(pTriangle, pBagLignesBase)
                End If

                'Créer un Bag vide de lignes simples
                pBagLignesSimples = New GeometryBag
                pBagLignesSimples.SpatialReference = pPolygon.SpatialReference
                'Interface pour ajouter les lignes simples
                pLigneColl = CType(pBagLignesSimples, IGeometryCollection)

                'Traiter toutes les lignes simples
                For i = 0 To pGeomCollSimple.GeometryCount - 1
                    'Si la ligne n'a pas été traitée
                    If Not pListeLignes.Contains(i) Then
                        'Ajouter la ligne simple à traiter
                        pLigneColl.AddGeometry(pGeomCollSimple.Geometry(i))
                    End If
                Next

                'Interface pour extraire les lignes simples
                pGeomCollSimple = pLigneColl

            Loop While pGeomCollSimple.GeometryCount > 0 And pRelResult.RelationElementCount > 0

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTriangle = Nothing
            pLigneColl = Nothing
            pGeomCollSimple = Nothing
            pListeLignes = Nothing
            pPolyline = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de traiter les lignes simples et les ajouter dans les lignes du squelette de base d'une polyligne.
    ''' </summary>
    ''' 
    ''' <param name="pPolyligne">Interface contenant la polyligne à traiter.</param>
    ''' <param name="pBagLignesSimples">Interface contenant les lignes simples du squelette primaire.</param>
    ''' <param name="pBagLignesBase">Interface contenant les lignes en double du squelette de base.</param>
    ''' 
    Protected Friend Shared Sub TraiterLignesSimples(ByVal pPolyligne As IPolyline, ByVal pBagLignesSimples As IGeometryBag, ByRef pBagLignesBase As IGeometryBag)
        'Déclarer les variables de travail
        Dim pTriangle As IPolygon = Nothing                     'Interface contenant un triangle.
        Dim pListeLignes As List(Of Integer)                    'Liste des numéros de ligne traitées.
        Dim pLigneColl As IGeometryCollection = Nothing         'Interface pour ajouter des lignes.
        Dim pGeomCollSimple As IGeometryCollection = Nothing    'Interface pour extraire les lignes simples dans le Bag.
        Dim pLigne As IPolyline = Nothing                       'Interface contenant une ligne.
        Dim pPointColl As IPointCollection = Nothing            'Interface utilisé pour extraire les sommets d'une géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour effectuer des opérations spatiales.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing       'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing             'Interface contenant le résultat du traitement de la relation spatiale.
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.
        Dim iNo As Integer = -2             'Contient le nombre de relation.
        Dim iDeb As Integer = -2            'Dernier numéro de ligne traitée.

        Try
            'Interface pour extraire les lignes simples
            pGeomCollSimple = CType(pBagLignesSimples, IGeometryCollection)

            'Traiter tant qu'il reste des lignes
            Do
                'Initialiser la liste des numéros de lignes traitées
                pListeLignes = New List(Of Integer)

                'Interface pour traiter les relations spatiales des lignes simples
                pRelOpNxM = CType(pBagLignesSimples, IRelationalOperatorNxM)
                'Traiter la relation spatiale entre les lignes simples
                pRelResult = pRelOpNxM.Touches(CType(pBagLignesSimples, IGeometryBag))
                'Trier le résultat
                pRelResult.SortRight()

                'Initialiser le nombre de connexion
                iNo = 0
                'Initialiser le numéro de ligne précédent
                iDeb = -1
                Dim i As Integer
                'Traiter toutes les relations
                For i = 0 To pRelResult.RelationElementCount - 1
                    'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                    pRelResult.RelationElement(i, iRel, iSel)
                    'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                    'Vérifier si c'est le même numéro
                    If iSel = iDeb Then
                        'Compter le nombre de relation
                        iNo = iNo + 1

                        'Si ce ne sont pas les mêmes numéros
                    Else
                        'Vérifier si une ou deux relation sont présentes
                        If iNo = 1 Or iNo = 2 Then
                            'Créer un triagle vide
                            pTriangle = New Polygon
                            pTriangle.SpatialReference = pPolyligne.SpatialReference
                            'Interface pour extraire les sommets du triangle
                            pPointColl = CType(pTriangle, IPointCollection)

                            'Extraire le résultat de la relation précédente
                            pRelResult.RelationElement(i - 1, iRel, iSel)
                            'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                            'Définir la ligne traitée
                            pLigne = CType(pGeomCollSimple.Geometry(iSel), IPolyline)
                            'Ajouter les sommets de la ligne dans le triangle 
                            pPointColl.AddPoint(pLigne.FromPoint)
                            pPointColl.AddPoint(pLigne.ToPoint)

                            'Définir la ligne en relation
                            pLigne = CType(pGeomCollSimple.Geometry(iRel), IPolyline)
                            'Ajouter les sommets de la ligne dans le triangle 
                            pPointColl.AddPoint(pLigne.FromPoint)
                            pPointColl.AddPoint(pLigne.ToPoint)

                            'Définir que la ligne traité a été traitée
                            If Not pListeLignes.Contains(iSel) Then pListeLignes.Add(iSel)
                            'Définir que la ligne en relation a été traitée
                            If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)

                            'Si le nombre de relation est 2
                            If iNo = 2 Then
                                'Extraire la deuxième relation
                                pRelResult.RelationElement(i - 2, iRel, iSel)
                                'Définir que la ligne en relation a été traitée
                                If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)
                            End If

                            'Revenir à la relation courante
                            pRelResult.RelationElement(i, iRel, iSel)

                            'Fermer le triangle
                            pTriangle.Close()
                            'Simplifier le triangle
                            pTopoOp = CType(pTriangle, ITopologicalOperator2)
                            pTopoOp.IsKnownSimple_2 = False
                            pTopoOp.Simplify()

                            'Ajouter les 2 plus petites longueur du triangle dans les lignes du squelette de base
                            Call TraiterLongueurTriangle(pTriangle, pBagLignesBase)
                        End If

                        'Initialiser le dernier numéro traité
                        iDeb = iSel
                        'Initialiser le nombre de relation
                        iNo = 1
                    End If
                Next

                'Vérifier si une ou deux relation sont présentes
                If iNo = 1 Or iNo = 2 Then
                    'Créer un triagle vide
                    pTriangle = New Polygon
                    pTriangle.SpatialReference = pPolyligne.SpatialReference
                    'Interface pour extraire les sommets du triangle
                    pPointColl = CType(pTriangle, IPointCollection)

                    'Extraire le résultat de la relation précédente
                    pRelResult.RelationElement(pRelResult.RelationElementCount - 1, iRel, iSel)
                    'Debug.Print((iSel + 1).ToString & "-" & (iRel + 1).ToString)

                    'Définir la ligne traitée
                    pLigne = CType(pGeomCollSimple.Geometry(iSel), IPolyline)
                    'Ajouter les sommets de la ligne dans le triangle
                    pPointColl.AddPoint(pLigne.FromPoint)
                    pPointColl.AddPoint(pLigne.ToPoint)

                    'Définir la ligne en relation
                    pLigne = CType(pGeomCollSimple.Geometry(iRel), IPolyline)
                    'Ajouter les sommets de la ligne dans le triangle
                    pPointColl.AddPoint(pLigne.FromPoint)
                    pPointColl.AddPoint(pLigne.ToPoint)

                    'Définir que la ligne traité a été traitée
                    If Not pListeLignes.Contains(iSel) Then pListeLignes.Add(iSel)
                    'Définir que la ligne en relation a été traitée
                    If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)

                    'Si le nombre de relation est 2
                    If iNo = 2 Then
                        'Extraire la deuxième relation
                        pRelResult.RelationElement(pRelResult.RelationElementCount - 2, iRel, iSel)
                        'Définir que la ligne en relation a été traitée
                        If Not pListeLignes.Contains(iRel) Then pListeLignes.Add(iRel)
                    End If

                    'Fermer le triangle
                    pTriangle.Close()
                    'Interface pour simplifier le triangle
                    pTopoOp = CType(pTriangle, ITopologicalOperator2)
                    'Simplifier le triangle
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()

                    'Ajouter les 2 plus petites longueur du triangle dans les lignes du squelette de base
                    Call TraiterLongueurTriangle(pTriangle, pBagLignesBase)
                End If

                'Créer un Bag vide de lignes simples
                pBagLignesSimples = New GeometryBag
                pBagLignesSimples.SpatialReference = pPolyligne.SpatialReference
                'Interface pour ajouter les lignes simples
                pLigneColl = CType(pBagLignesSimples, IGeometryCollection)

                'Traiter toutes les lignes simples
                For i = 0 To pGeomCollSimple.GeometryCount - 1
                    'Si la ligne n'a pas été traitée
                    If Not pListeLignes.Contains(i) Then
                        'Ajouter la ligne simple à traiter
                        pLigneColl.AddGeometry(pGeomCollSimple.Geometry(i))
                    End If
                Next

                'Interface pour extraire les lignes simples
                pGeomCollSimple = pLigneColl

            Loop While pGeomCollSimple.GeometryCount > 0 And pRelResult.RelationElementCount > 0 And pListeLignes.Count > 0

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTriangle = Nothing
            pLigneColl = Nothing
            pGeomCollSimple = Nothing
            pListeLignes = Nothing
            pLigne = Nothing
            pPointColl = Nothing
            pTopoOp = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet d'ajouter les deux plus petites lignes d'un triangle dans le Bag des lignes du squelette de base.
    ''' </summary>
    ''' 
    ''' <param name="pTriangle">Interface contenant un triangle.</param>
    ''' <param name="pBagLignesBase">Interface contenant les lignes du squelette de base.</param>
    ''' 
    Private Shared Sub TraiterLongueurTriangle(ByVal pTriangle As IPolygon, ByRef pBagLignesBase As IGeometryBag)
        'Déclarer les variables de travail
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface pour ajouter des lignes.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments des lignes du triangle.
        Dim pSegCollAdd As ISegmentCollection = Nothing     'Interface pour ajouter des segments dans une ligne.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne.

        Try
            'Interface pour ajouter des lignes
            pLigneColl = CType(pBagLignesBase, IGeometryCollection)

            'Interface pour extraire les segments de lignes du triangle
            pSegColl = CType(pTriangle, ISegmentCollection)

            'Vérifier si c'est un triangle
            If pSegColl.SegmentCount = 3 Then
                'Vérifier si la longueur du premier segment est inférieure au deuxième
                If pSegColl.Segment(0).Length < pSegColl.Segment(1).Length Then
                    'Créer une ligne vide
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = pTriangle.SpatialReference
                    'Interface pour ajouter un segment dans la ligne
                    pSegCollAdd = CType(pPolyline, ISegmentCollection)
                    'Ajouter un segment dans la ligne
                    pSegCollAdd.AddSegment(pSegColl.Segment(0))
                    'Ajouter la ligne dans le Bag du squelette de base
                    pLigneColl.AddGeometry(pPolyline)

                    'Créer une ligne vide
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = pTriangle.SpatialReference
                    'Interface pour ajouter un segment dans la ligne
                    pSegCollAdd = CType(pPolyline, ISegmentCollection)

                    'Vérifier si la longueur du premier segment est inférieure au troisième
                    If pSegColl.Segment(1).Length < pSegColl.Segment(2).Length Then
                        'Ajouter un segment dans la ligne
                        pSegCollAdd.AddSegment(pSegColl.Segment(1))

                        'Si la longueur du premier segment est supérieure au troisième
                    Else
                        'Ajouter un segment dans la ligne
                        pSegCollAdd.AddSegment(pSegColl.Segment(2))
                    End If

                    'Ajouter la ligne dans le Bag du squelette de base
                    pLigneColl.AddGeometry(pPolyline)

                    'Si la longueur du premier segment est supérieure ou égale au deuxième
                Else
                    'Créer une ligne vide
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = pTriangle.SpatialReference
                    'Interface pour ajouter un segment dans la ligne
                    pSegCollAdd = CType(pPolyline, ISegmentCollection)
                    'Ajouter un segment dans la ligne
                    pSegCollAdd.AddSegment(pSegColl.Segment(1))
                    'Ajouter la ligne dans le Bag du squelette de base
                    pLigneColl.AddGeometry(pPolyline)

                    'Créer une ligne vide
                    pPolyline = New Polyline
                    pPolyline.SpatialReference = pTriangle.SpatialReference
                    'Interface pour ajouter un segment dans la ligne
                    pSegCollAdd = CType(pPolyline, ISegmentCollection)
                    'Vérifier si la longueur du premier segment est inférieure au troisième
                    If pSegColl.Segment(0).Length < pSegColl.Segment(2).Length Then
                        'Ajouter un segment dans la ligne
                        pSegCollAdd.AddSegment(pSegColl.Segment(0))

                        'Si la longueur du premier segment est supérieure au troisième
                    Else
                        'Ajouter un segment dans la ligne
                        pSegCollAdd.AddSegment(pSegColl.Segment(2))
                    End If

                    'Ajouter la ligne dans le Bag du squelette de base
                    pLigneColl.AddGeometry(pPolyline)
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLigneColl = Nothing
            pSegColl = Nothing
            pSegCollAdd = Nothing
            pPolyline = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet d'ajouter 3 lignes à partir du centre d'un triangle dans le Bag des lignes du squelette de base.
    ''' </summary>
    ''' 
    ''' <param name="pTriangle">Interface contenant un triangle.</param>
    ''' <param name="pBagLignesBase">Interface contenant les lignes du squelette de base.</param>
    ''' 
    Private Shared Sub TraiterCentreTriangle(ByVal pTriangle As IPolygon, ByRef pBagLignesBase As IGeometryBag)
        'Déclarer les variables de travail
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface pour ajouter des lignes.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets du triangle.
        Dim pPointCollAdd As IPointCollection = Nothing     'Interface pour ajouter les sommets dans la ligne.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne.
        Dim pArea As IArea = Nothing                        'Interface utilisé pour extraire le centre d'un triangle
        Dim pPoint As IPoint = Nothing                      'Interface contenant le centre du triangle.

        Try
            'Interface pour ajouter des lignes
            pLigneColl = CType(pBagLignesBase, IGeometryCollection)

            'Interface utilisé pour extraire le centre d'un triangle
            pArea = CType(pTriangle, IArea)
            'Extraire le centre d'un triangle
            pPoint = pArea.LabelPoint()

            'Interface pour extraire les sommets du triangle
            pPointColl = CType(pTriangle, IPointCollection)

            'Traiter les sommets du triangle
            For j = 0 To pPointColl.PointCount - 2
                'Créer une nouvelle ligne vide
                pPolyline = New Polyline
                pPolyline.SpatialReference = pTriangle.SpatialReference
                'Interface pour ajouter des sommets dans la ligne
                pPointCollAdd = CType(pPolyline, IPointCollection)

                'Ajouter les sommets dans la ligne
                pPointCollAdd.AddPoint(pPointColl.Point(j))
                pPointCollAdd.AddPoint(pPoint)

                'Ajouter la ligne dans le Bag des ligne du squelette de base
                pLigneColl.AddGeometry(pPolyline)
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLigneColl = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
            pPolyline = Nothing
            pArea = Nothing
            pPoint = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Function qui permet de créer une nouvelle Polyline à partir des centres des droites de la polyligne d'entrée.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Interface contenant les lignes des triangles de Delaunay.</param>
    ''' 
    ''' <returns>IPolyline contenant le squelette de base du polygone.</returns>
    ''' 
    Private Shared Function TraiterCentreDroites(ByVal pPolyline As IPolyline) As IPolyline
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing          'Interface pour extraire les lignes de la polyligne à traiter.
        Dim pGeomCollAdd As IGeometryCollection = Nothing       'Interface pour ajouter les lignes de la polyligne à retourner.
        Dim pPath As IPath = Nothing                            'Interface contenant la nouvelle ligne.
        Dim pPointColl As IPointCollection = Nothing            'Interface pour extraire les sommets de la ligne.
        Dim pPointCollAdd As IPointCollection = Nothing         'Interface pour ajouter les sommets de la nouvelle ligne.
        Dim pPoint As IPoint = Nothing
        Dim pPointA As IPoint = Nothing
        Dim pPointB As IPoint = Nothing

        'Définir la valeur par défaut
        TraiterCentreDroites = New Polyline
        TraiterCentreDroites.SpatialReference = pPolyline.SpatialReference

        Try
            'Interface pour extraire les composantes de la polyligne
            pGeomCollAdd = CType(TraiterCentreDroites, IGeometryCollection)

            'Interface pour extraire les composantes de la polyligne
            pGeomColl = CType(pPolyline, IGeometryCollection)

            'Traiter toutes les composantes de la polyligne
            For i = 0 To pGeomColl.GeometryCount - 1
                'Interface pour extraire les sommets de la ligne
                pPointColl = CType(pGeomColl.Geometry(i), IPointCollection)

                'Vérifier si seulement 2 sommets
                If pPointColl.PointCount = 2 Then
                    'Ajouter la ligne originale
                    pGeomCollAdd.AddGeometry(pGeomColl.Geometry(i))

                    'Si plus de deux sommets
                Else
                    'Créer une nouvelle ligne vide
                    pPath = New Path
                    pPath.SpatialReference = pPolyline.SpatialReference

                    'Interface pour ajouter les sommets de la nouvelle ligne
                    pPointCollAdd = CType(pPath, IPointCollection)

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
                        pPoint.SpatialReference = pPolyline.SpatialReference
                        pPoint.X = (pPointA.X + pPointB.X) / 2
                        pPoint.Y = (pPointA.Y + pPointB.Y) / 2
                        'Ajouter le premier sommet
                        pPointCollAdd.AddPoint(pPoint)
                    Next

                    'Ajouter le dernier sommet
                    pPointCollAdd.AddPoint(pPointB)

                    'Ajouter la nouvelle ligne
                    pGeomCollAdd.AddGeometry(pPath)
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pPath = Nothing
            pPointColl = Nothing
            pPointCollAdd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet d'enlever les lignes inutiles d'un squelette et retourner les extrémités non connectés.
    ''' Les extrémités de ligne qui possède une longueur inférieure à la longueur minimale seront enlever du squelette s'ils ne sont pas connectés. 
    ''' </summary>
    ''' 
    ''' <param name="pMultiPoint">Sommets pour lesquels le squelette doit être connecté.</param>
    ''' <param name="pSquelette">Lignes contenant le squelette de départ.</param>
    ''' <param name="dLongMin">Contient la longueur minimale des lignes à conserver dans le squelette.</param>
    ''' 
    ''' <returns>IMultipoint contenant les extrémités du squelette non connectées.</returns>
    ''' 
    Protected Friend Shared Function EnleverExtremiteLigne(ByVal pMultipoint As IMultipoint, ByRef pSquelette As IPolyline, Optional ByVal dLongMin As Double = 50) As IMultipoint
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes du nouveau squelette.
        Dim pLimite As IMultipoint = Nothing                'Interface contenant les limites du squelette.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire les différences.
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface contenant une ligne.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPath As IPath = Nothing                        'Interface contenant une composante de ligne.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier les extrémité de ligne.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour extraire les segments.
        Dim dLong As Double = 0

        'Définir la valeur par défaut
        EnleverExtremiteLigne = Nothing

        Try
            'Traiter tout
            Do
                'Vérifier si la longueur est 0
                If dLong = 0 Then
                    'On défini la première vérification de la longueur minimale à moitié
                    dLong = dLongMin / 2
                    'Si la longueur n'est pas à 0, ce n'est pas la première vérification
                Else
                    'On défini la longueur minimale demandé si ce n'est pas la première vérification
                    dLong = dLongMin
                End If

                'Identifier les limites du squelette
                pLimite = IdentifierLimiteSquelette(pSquelette)

                'Interface pour enlever les sommets de conexion des limites
                pTopoOp = CType(pLimite, ITopologicalOperator2)
                pTopoOp.IsKnownSimple_2 = False
                pTopoOp.Simplify()

                'Interface pour trouver les extrémités de ligne non connecté
                pRelOp = CType(pTopoOp.Difference(pMultipoint), IRelationalOperator)

                'Créer le nouveau squelette sans les lignes qui ne sont pas une extrémité à enlever
                '---------------------------------------------------------
                'Créer une nouvelle polyligne vide
                pPolyline = New Polyline
                pPolyline.SpatialReference = pSquelette.SpatialReference
                'Interface pour ajouter les lignes qui ne sont pas les extrémités de lignes
                pLigneColl = CType(pPolyline, IGeometryCollection)

                'Interface pour traiter toutes les lignes
                pGeomColl = CType(pSquelette, IGeometryCollection)
                'Traiter tous les composantes de géométrie
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Définir la ligne à traiter
                    pPath = CType(pGeomColl.Geometry(i), IPath)
                    'Interface pour extraire les segments de lignes
                    pSegColl = CType(pPath, ISegmentCollection)

                    'Vérifier si la ligne est une extrémité au premier sommet
                    If pRelOp.Contains(pPath.FromPoint) Then
                        'Vérifier si la longueur de la ligne est supérieure à la longueur minimale
                        If pPath.Length - pSegColl.Segment(pSegColl.SegmentCount - 1).Length > dLong Then
                            'Conserver la ligne dans le squelette
                            pLigneColl.AddGeometry(pGeomColl.Geometry(i))
                        End If

                        'Si la ligne est une extrémité au dernier sommet
                    ElseIf pRelOp.Contains(pPath.ToPoint) Then
                        'Vérifier si la longueur de la ligne est supérieure à la longueur minimale
                        If pPath.Length - pSegColl.Segment(0).Length > dLong Then
                            'Conserver la ligne dans le squelette
                            pLigneColl.AddGeometry(pGeomColl.Geometry(i))
                        End If

                        'Si la ligne n'est pas une extrémité
                    Else
                        'Conserver la ligne dans le squelette
                        pLigneColl.AddGeometry(pGeomColl.Geometry(i))
                    End If
                Next

                'Vérifier si la géométrie n'est pas vide
                If pLigneColl.GeometryCount > 0 Then
                    'Redéfinir la polyligne d'entrée
                    pSquelette = pPolyline
                    'Interface pour simplifier le résultat
                    pTopoOp = TryCast(pSquelette, ITopologicalOperator2)
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()
                End If

                'Traiter tant qu'il y a des changements
            Loop While (pGeomColl.GeometryCount <> pLigneColl.GeometryCount And pLigneColl.GeometryCount > 0) Or dLong <> dLongMin

            'Définir les extrémités du squelette non connectés
            '---------------------------------------------------------
            EnleverExtremiteLigne = CType(pRelOp, IMultipoint)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pTopoOp = Nothing
            pRelOp = Nothing
            pLigneColl = Nothing
            pGeomColl = Nothing
            pPath = Nothing
            pLimite = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function qui permet de retourner les sommets pour lesquels le squelette du polygone doit être connecté.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Polygone utilisé pour créer le squelette.</param>
    ''' <param name="pGeometryBagRel">Sommets pour lesquels le squelette doit être connecté.</param>
    ''' 
    ''' <returns>IMultipoint contenant les sommets pour lesquels le squelette doit être connecté.</returns>
    ''' 
    Protected Friend Shared Function ExtraireSommetsRelation(ByVal pPolygon As IPolygon, ByVal pGeometryBagRel As IGeometryBag) As IMultipoint
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pPointsIntersect As IGeometry = Nothing         'Interface contenant les points d'intersections.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets en relation.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pGeometry As IGeometry = Nothing                'Interface contenant la géométrie en relation.

        Try
            'Créer la géométrie des sommets en relation par défaut
            ExtraireSommetsRelation = New Multipoint
            ExtraireSommetsRelation.SpatialReference = pPolygon.SpatialReference

            'Interface pour extraire le sommets en relation
            pPointColl = CType(ExtraireSommetsRelation, IPointCollection)

            'Interface pour extraire le sommets en relation
            pTopoOp = CType(pPolygon, ITopologicalOperator2)

            'Interface pour extraire les géométries en relation
            pGeomColl = CType(pGeometryBagRel, IGeometryCollection)

            'Traiter toutes les géométries en relation
            For i = 0 To pGeomColl.GeometryCount - 1
                'Définir la géométrie en relation
                pGeometry = pGeomColl.Geometry(i)

                'Définir les points d'intersection
                pPointsIntersect = pTopoOp.Intersect(pGeometry, esriGeometryDimension.esriGeometry0Dimension)

                'Ajouter les sommets en relation
                pPointColl.AddPointCollection(CType(pPointsIntersect, IPointCollection))
            Next

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(ExtraireSommetsRelation, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pPointsIntersect = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
            pGeometry = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function qui permet de retourner les points d'intersection entre les composantes du polygone et entre le polygone et ses géométries en relation.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Polygone utilisé pour créer le squelette.</param>
    ''' <param name="pGeometryBagRel">Sommets pour lesquels le squelette doit être connecté.</param>
    ''' 
    ''' <returns>IMultipoint contenant les sommets pour lesquels le squelette doit être connecté.</returns>
    ''' 
    Protected Friend Shared Function ExtrairePointsIntersection(ByVal pPolygon As IPolygon, ByVal pGeometryBagRel As IGeometryBag) As IMultipoint
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pPointsIntersect As IGeometry = Nothing         'Interface contenant les points d'intersections.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets en relation.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pGeomCollAdd As IGeometryCollection = Nothing   'Interface pour ajouter les géométries dans une ligne.
        Dim pLigneA As IPolyline = Nothing                  'Interface contenant la ligne A des limites du polygone.
        Dim pLigneB As IPolyline = Nothing                  'Interface contenant la ligne A des limites du polygone.

        Try
            'Créer la géométrie des sommets en relation par défaut
            ExtrairePointsIntersection = New Multipoint
            ExtrairePointsIntersection.SpatialReference = pPolygon.SpatialReference

            'Interface pour extraire le sommets en relation
            pPointColl = CType(ExtrairePointsIntersection, IPointCollection)

            'Interface pour extraire le sommets en relation
            pTopoOp = CType(pPolygon, ITopologicalOperator2)

            'Interface pour extraire les géométries en relation
            pGeomColl = CType(pGeometryBagRel, IGeometryCollection)

            'Traiter toutes les géométries en relation
            For i = 0 To pGeomColl.GeometryCount - 1
                'Définir les points d'intersection
                pPointsIntersect = pTopoOp.Intersect(pGeomColl.Geometry(i), esriGeometryDimension.esriGeometry0Dimension)

                'Ajouter les sommets en relation
                pPointColl.AddPointCollection(CType(pPointsIntersect, IPointCollection))
            Next

            'Interface pour extraire les limites du polygone
            pTopoOp = CType(pPolygon, ITopologicalOperator2)
            'Interface pour extraire les points d'intersections entre les limites du polygone
            pGeomColl = CType(pTopoOp.Boundary, IGeometryCollection)

            'Traiter toutes les lignes des limites du polygone
            For i = 0 To pGeomColl.GeometryCount - 2
                'Définir la ligne A
                pLigneA = New Polyline
                pLigneA.SpatialReference = pPolygon.SpatialReference

                'Interface pour ajouter les géométries
                pGeomCollAdd = CType(pLigneA, IGeometryCollection)
                'Ajouter les géométries
                pGeomCollAdd.AddGeometry(pGeomColl.Geometry(i))

                'Interface pour extraire les points d'intersection
                pTopoOp = CType(pLigneA, ITopologicalOperator2)

                'Avec toutes les lignes des limites du polygone
                For j = i + 1 To pGeomColl.GeometryCount - 1
                    'Définir la ligne A
                    pLigneB = New Polyline
                    pLigneB.SpatialReference = pPolygon.SpatialReference

                    'Interface pour ajouter les géométries
                    pGeomCollAdd = CType(pLigneB, IGeometryCollection)
                    'Ajouter les géométries
                    pGeomCollAdd.AddGeometry(pGeomColl.Geometry(j))

                    'Définir les points d'intersection
                    pPointsIntersect = pTopoOp.Intersect(pLigneB, esriGeometryDimension.esriGeometry0Dimension)

                    'Ajouter les sommets en relation
                    pPointColl.AddPointCollection(CType(pPointsIntersect, IPointCollection))
                Next
            Next

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(ExtrairePointsIntersection, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pPointsIntersect = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
            pGeomCollAdd = Nothing
            pLigneA = Nothing
            pLigneB = Nothing
        End Try
    End Function
    ''' <summary>
    ''' Function qui permet de retourner les points d'intersection entre un élément et ses éléments en relation.
    ''' </summary>
    ''' 
    ''' <param name="pFeature">Interface contenant l'élément pour lequel on veut extraire ses points de connexion.</param>
    ''' <param name="pTopologyGraph">Interface contenant la topologie d'un élément et ses éléments en relation.</param>
    ''' 
    ''' <returns>IMultipoint contenant les sommets pour lesquels le squelette doit être connecté.</returns>
    ''' 
    Protected Friend Shared Function ExtrairePointsIntersection(ByVal pFeature As IFeature, ByVal pTopologyGraph As ITopologyGraph) As IMultipoint
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing          'Interface pour extraire le sommets en relation.
        Dim pPointColl As IPointCollection = Nothing            'Interface pour ajouter les sommets en relation.
        Dim pEnumTopoParents As IEnumTopologyParent = Nothing   'Interface contenant le nombre d'élément au point d'intersection.
        Dim pEnumTopoNode As IEnumTopologyNode = Nothing        'Interface pour extraire les points de connexion.
        Dim pTopoNode As ITopologyNode = Nothing                'Interface contenant un point de connexion.
        Dim pEnumTopoEdge As IEnumTopologyEdge = Nothing        'Interface pour extraire les lignes de connexion.
        Dim pTopoEdge As ITopologyEdge = Nothing                'Interface contenant une ligne de connexion.
        Dim pPolyline As IPolyline = Nothing                    'Interface pour extraire le centre de la ligne.
        Dim pPointCollDiff As IPointCollection = Nothing        'Interface pour ajouter les sommets de différence.
        Dim pPointDiff As IMultipoint = Nothing                 'Interface contenant les sommets de différence.
        Dim pPoint As IPoint = Nothing                          'Interface contenant un point d'intersection.

        Try
            'Créer la géométrie des sommets en relation par défaut
            ExtrairePointsIntersection = New Multipoint
            ExtrairePointsIntersection.SpatialReference = pFeature.Shape.SpatialReference

            'Interface pour extraire le sommets en relation
            pPointColl = CType(ExtrairePointsIntersection, IPointCollection)

            'Créer un multipoint vide
            pPointDiff = New Multipoint
            pPointDiff.SpatialReference = pFeature.Shape.SpatialReference

            'Interface pour extraire le sommets en relation
            pPointCollDiff = CType(pPointDiff, IPointCollection)

            'Extraire les lignes de connexion de la topologie
            pEnumTopoEdge = pTopologyGraph.GetParentEdges(CType(pFeature.Class, IFeatureClass), pFeature.OID)
            'Initialiser l'extraction des lignes de connexion
            pEnumTopoEdge.Reset()
            'Extraire la première ligne de connexion
            pTopoEdge = pEnumTopoEdge.Next

            'Extraire toutes les lignes de connexion
            Do Until pTopoEdge Is Nothing
                'Extraire le nombre d'éléments
                pEnumTopoParents = pTopoEdge.Parents()
                'Vérifier si plusieurs Edges
                If pEnumTopoParents.Count > 1 Then
                    'vérifier si la ligne est valide
                    If pTopoEdge.Geometry IsNot Nothing Then
                        'Définir la ligne
                        pPolyline = CType(pTopoEdge.Geometry, IPolyline)
                        'Créer un nouveau point vide
                        pPoint = New Point
                        pPoint.SpatialReference = pPolyline.SpatialReference
                        'Extraire le centre de la ligne
                        pPolyline.QueryPoint(esriSegmentExtension.esriNoExtension, pPolyline.Length / 2, False, pPoint)
                        'Ajouter le point de connexion
                        pPointColl.AddPoint(pPoint)
                        'Ajouter le point de connexion
                        pPointCollDiff.AddPoint(pPolyline.FromPoint)
                        'Ajouter le point de connexion
                        pPointCollDiff.AddPoint(pPolyline.ToPoint)
                    End If
                End If

                'Extraire la prochaine ligne de connexion
                pTopoEdge = pEnumTopoEdge.Next
            Loop

            'Extraire les points de connexion de la topologie
            pEnumTopoNode = pTopologyGraph.GetParentNodes(CType(pFeature.Class, IFeatureClass), pFeature.OID)
            'Initialiser l'extraction des points de connexion
            pEnumTopoNode.Reset()
            'Extraire le premier point de connexion
            pTopoNode = pEnumTopoNode.Next

            'Extraire tous les points de connexion
            Do Until pTopoNode Is Nothing
                'Extraire le nombre d'éléments
                pEnumTopoParents = pTopoNode.Parents()
                'Vérifier si plus de 2 Edges ou si plus de 1 élément
                If pTopoNode.Degree > 2 Or pEnumTopoParents.Count > 1 Then
                    'Ajouter le point de connexion
                    pPointColl.AddPoint(CType(pTopoNode.Geometry, IPoint))
                End If

                'Extraire le prochain point de connexion
                pTopoNode = pEnumTopoNode.Next
            Loop

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(ExtrairePointsIntersection, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

            'Enlever les points de différence
            ExtrairePointsIntersection = CType(pTopoOp.Difference(pPointDiff), IMultipoint)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pPointColl = Nothing
            pEnumTopoParents = Nothing
            pEnumTopoNode = Nothing
            pTopoNode = Nothing
            pEnumTopoEdge = Nothing
            pTopoEdge = Nothing
            pPolyline = Nothing
            pPointCollDiff = Nothing
            pPointDiff = Nothing
            pPoint = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet de connecter le squelette aux points de connexion des éléments en relation selon les points les plus proche du squelette simplifier.
    ''' </summary>
    ''' 
    ''' <param name="pPointsConnexion">Points pour lesquels le squelette doit être connecté.</param>
    ''' <param name="dLargMin">Largeur minimun de connexion utilisé pour éliminer les extrémité du squelette temporaire.</param>
    ''' <param name="dLongMin">Contient la longueur minimale des lignes à conserver dans le squelette.</param>
    ''' <param name="pSquelette">Polyline contenant le squelette utilisé pour se connecter aux points des éléments en relation.</param>
    ''' 
    Protected Friend Shared Sub ConnecterSquelettePointsConnexion(ByVal pPointsConnexion As IMultipoint, ByVal dLargMin As Double, ByVal dLongMin As Double, _
                                                                  ByRef pSquelette As IPolyline)
        'Déclarer les variables de travail
        Dim pSqueletteTmp As IPolyline = Nothing            'Interface contenant le squelette simplifié temporaire de connexion.
        Dim pClone As IClone = Nothing                      'Interface utiliser pour clone le squelette.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets en relation.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pPoint As IPoint = Nothing                      'Interface ESRI contenant le sommet trouvé.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface utilisé pour vérifier la connexion.
        Dim pProxOp As IProximityOperator = Nothing         'Interface utilisé pour trouver le sommet le plus proche.
        Dim pPath As IPath = Nothing                        'Interface contenant la ligne à ajouter au squellette pour effectuer la connexion.
        Dim pLigne As IPolyline = Nothing                   'Interface contenant la ligne de connexion.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant les points d'intersection.
        Dim dLong As Double = 0                             'Contient la longueur minimum.

        Try
            'Interface utilisé pour extraire le point le plus proche
            pRelOp = CType(pSquelette, IRelationalOperator)

            'Interface utilisé pour ajouter une nouvelle ligne
            pGeomColl = CType(pSquelette, IGeometryCollection)

            'Interface pour cloner le squelette
            pClone = CType(pSquelette, IClone)
            'Cloner le squelette
            pSqueletteTmp = CType(pClone.Clone, IPolyline)

            'vérifier si la longueur minimum est plus grande que la largueur minimum
            If dLongMin > dLargMin Then
                'Définir la longueur minimum selon la longueur minimum
                dLong = dLongMin
            Else
                'Définir la longueur minimum selon la largeur minimum
                dLong = dLargMin
            End If

            'Enlever les extrémités de lignes superflux dans le squelette et retourner les points du squelette non-connectés
            Call EnleverExtremiteLigne(pPointsConnexion, pSqueletteTmp, dLong)

            'Interface pour extraire le sommets en relation
            pPointColl = CType(pPointsConnexion, IPointCollection)

            'Traiter tous les points en relation
            For i = 0 To pPointColl.PointCount - 1
                'Vérifier si le point n'est pas déjà connecté
                If pRelOp.Disjoint(pPointColl.Point(i)) Then
                    'Interface utilisé pour extraire le point le plus proche
                    pProxOp = CType(pSqueletteTmp, IProximityOperator)

                    'Créer un nouveau sommet vide
                    pPoint = New Point
                    pPoint.SpatialReference = pSquelette.SpatialReference

                    'Trouver le sommet le plus proche
                    pProxOp.QueryNearestPoint(pPointColl.Point(i), esriSegmentExtension.esriNoExtension, pPoint)

                    'Creér la ligne de connexion
                    pLigne = New Polyline
                    pLigne.SpatialReference = pSquelette.SpatialReference
                    pLigne.FromPoint = pPoint
                    pLigne.ToPoint = pPointColl.Point(i)

                    'Vérifier si la ligne croise le squelette
                    If pRelOp.Crosses(pLigne) Then
                        'Interface pour extraire les points de connexion
                        pTopoOp = CType(pLigne, ITopologicalOperator2)
                        'Extraire les points de connexion
                        pMultipoint = CType(pTopoOp.Intersect(pSquelette, esriGeometryDimension.esriGeometry0Dimension), IMultipoint)

                        'Interface utilisé pour extraire le point le plus proche
                        pProxOp = CType(pMultipoint, IProximityOperator)

                        'Trouver le sommet le plus proche dans les points de connexion
                        pProxOp.QueryNearestPoint(pPointColl.Point(i), esriSegmentExtension.esriNoExtension, pPoint)
                    End If

                    'Créer une nouvelle ligne
                    pPath = New Path
                    pPath.SpatialReference = pSquelette.SpatialReference

                    'Définir les sommets de la ligne
                    pPath.FromPoint = pPoint
                    pPath.ToPoint = pPointColl.Point(i)

                    'Ajouter la nouvelle ligne au squellette
                    pGeomColl.AddGeometry(pPath)
                End If
            Next

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pSqueletteTmp = Nothing
            pClone = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
            pPoint = Nothing
            pProxOp = Nothing
            pPath = Nothing
            pRelOp = Nothing
            pLigne = Nothing
            pMultipoint = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de connecter le squelette aux points de connexion des éléments en relation selon les points les plus proche du squelette non simplifier.
    ''' </summary>
    ''' 
    ''' <param name="pPointsConnexion">Points pour lesquels le squelette doit être connecté.</param>
    ''' <param name="pSquelette">Polyline contenant le squelette utilisé pour se connecter aux points des éléments en relation.</param>
    ''' 
    Protected Friend Shared Sub ConnecterSquelettePointsConnexion(ByVal pPointsConnexion As IMultipoint, ByRef pSquelette As IPolyline)
        'Déclarer les variables de travail
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets en relation.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pPoint As IPoint = Nothing                      'Interface ESRI contenant le sommet trouvé.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface utilisé pour vérifier la connexion.
        Dim pProxOp As IProximityOperator = Nothing         'Interface utilisé pour trouver le sommet le plus proche.
        Dim pPath As IPath = Nothing                        'Interface contenant la ligne à ajouter au squellette pour effectuer la connexion.

        Try
            'Interface utilisé pour ajouter une nouvelle ligne
            pGeomColl = CType(pSquelette, IGeometryCollection)

            'Interface utilisé pour extraire le point le plus proche
            pRelOp = CType(pSquelette, IRelationalOperator)

            'Interface utilisé pour extraire le point le plus proche
            pProxOp = CType(pSquelette, IProximityOperator)

            'Interface pour extraire le sommets en relation
            pPointColl = CType(pPointsConnexion, IPointCollection)

            'Traiter tous les sommets en relation
            For i = 0 To pPointColl.PointCount - 1
                'Vérifier si le point n'est pas déjà connecté
                If pRelOp.Disjoint(pPointColl.Point(i)) Then
                    'Créer un nouveau sommet vide
                    pPoint = New Point
                    pPoint.SpatialReference = pSquelette.SpatialReference

                    'Trouver le sommet le plus proche
                    pProxOp.QueryNearestPoint(pPointColl.Point(i), esriSegmentExtension.esriNoExtension, pPoint)

                    'Créer une nouvelle ligne
                    pPath = New Path
                    pPath.SpatialReference = pSquelette.SpatialReference

                    'Définir les sommets de la ligne
                    pPath.FromPoint = pPoint
                    pPath.ToPoint = pPointColl.Point(i)

                    'Ajouter la nouvelle ligne au squellette
                    pGeomColl.AddGeometry(pPath)
                End If
            Next

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pTopoOp = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
            pPoint = Nothing
            pProxOp = Nothing
            pPath = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de connecter le squelette aux points de connexion des éléments en relation en fonction de l'angle des points de connexion du polygone.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Interface contenant le polygone à traiter.</param>
    ''' <param name="pPointsConnexion">Interface contenant les points pour lesquels le squelette doit être connecté.</param>
    ''' <param name="pSquelette">Polyline contenant le squelette utilisé pour se connecter aux points des éléments en relation.</param>
    ''' 
    Protected Friend Shared Sub ConnecterSquelettePointsConnexion(ByVal pPolygon As IPolygon, ByVal pPointsConnexion As IMultipoint, ByRef pSquelette As IPolyline)
        'Déclarer les variables de travail
        Dim pGeometry As IGeometry = Nothing                'Interface contenant une géométrie.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets en relation.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pPoint As IPoint = Nothing                      'Interface ESRI contenant le sommet trouvé.
        Dim pRelOp As IRelationalOperator = Nothing         'Interface utilisé pour vérifier la connexion.
        Dim pProxOp As IProximityOperator = Nothing         'Interface utilisé pour trouver le sommet le plus proche.
        Dim pPath As IPath = Nothing                        'Interface contenant la ligne à ajouter au squellette pour effectuer la connexion.
        Dim pLigneMax As IPolyline = Nothing                'Interface contenant la ligne maximum.
        Dim pMultipoint As IMultipoint = Nothing            'Interface contenant les points de connexion entre la ligne maximum et le squelette.
        Dim pHitTest As IHitTest = Nothing                  'Interface pour tester la présence du sommet recherché
        Dim pRingColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes du polygone.
        Dim pPntColl As IPointCollection = Nothing          'Interface pour extraire les sommets du polygone.
        Dim pNewPoint As IPoint = Nothing                   'Interface contenant le nouveau point trouvé (pas utilisé).
        Dim dDistance As Double = Nothing                   'Interface contenant la distance calculée entre le point de recherche et le sommet trouvé
        Dim nNumeroPartie As Integer = Nothing              'Numéro de partie trouvée
        Dim nNumeroSommet As Integer = Nothing              'Numéro de sommet de la partie trouvée
        Dim bCoteDroit As Boolean = Nothing                 'Indiquer si le point trouvé est du côté droit de la géométrie
        Dim dAngle As Double = 0                            'Angle du sommet trouvé dans le polygone.
        Dim dAngle1 As Double = 0                           'Angle de la droite précédente du sommet trouvé dans le polygone.
        Dim dAngle2 As Double = 0                           'Angle de la droite suivante du sommet trouvé dans le polygone.

        Try
            'Interface pour trouver un point sur le polygone
            pHitTest = CType(pPolygon, IHitTest)
            'Interface pour extraire les composantes du polygone
            pRingColl = CType(pPolygon, IGeometryCollection)

            'Interface utilisé pour ajouter une nouvelle ligne
            pGeomColl = CType(pSquelette, IGeometryCollection)

            'Interface pour extraire le sommets en relation
            pPointColl = CType(pPointsConnexion, IPointCollection)

            'Traiter tous les sommets en relation
            For i = 0 To pPointColl.PointCount - 1
                'Interface utilisé pour extraire le point le plus proche
                pRelOp = CType(pSquelette, IRelationalOperator)

                'Vérifier si le point n'est pas déjà connecté
                If pRelOp.Disjoint(pPointColl.Point(i)) Then
                    'Rechercher le point du polygone par rapport à chaque point d'intersection
                    If pHitTest.HitTest(pPointColl.Point(i), 0.001, esriGeometryHitPartType.esriGeometryPartBoundary, _
                                        pNewPoint, dDistance, nNumeroPartie, nNumeroSommet, bCoteDroit) Then
                        'Interface pour extraire le sommet de la composante du polygone
                        pPntColl = CType(pRingColl.Geometry(nNumeroPartie), IPointCollection)

                        'Vérifier si c'est le premier sommet à extraire
                        If nNumeroSommet = 0 Then
                            'Définir l'angle de la ligne de connexion
                            dAngle1 = clsGeneraliserGeometrie.Angle(pPntColl.Point(nNumeroSommet), pPntColl.Point(pPntColl.PointCount - 2))
                            dAngle2 = clsGeneraliserGeometrie.Angle(pPntColl.Point(nNumeroSommet), pPntColl.Point(nNumeroSommet + 1))
                            'Si ce n'est pas le premier sommet
                        Else
                            'Définir l'angle de la ligne de connexion
                            dAngle1 = clsGeneraliserGeometrie.Angle(pPntColl.Point(nNumeroSommet), pPntColl.Point(nNumeroSommet - 1))
                            dAngle2 = clsGeneraliserGeometrie.Angle(pPntColl.Point(nNumeroSommet), pPntColl.Point(nNumeroSommet + 1))
                        End If

                        'S'assurer que l'angle2 est supérieure à l'angle1
                        If dAngle2 < dAngle1 Then dAngle2 = dAngle2 + 360
                        'Calculer l'angle du point trouvé
                        dAngle = (dAngle1 + dAngle2) / 2
                        'S'assurer que l'angle est supérieure à 0
                        If dAngle < 0 Then dAngle = dAngle + 360
                        'Sassurer que l'angle est inférieure à 360
                        If dAngle >= 360 Then dAngle = dAngle - 360

                        'Définir la distance de la nouvelle ligne maximum
                        dDistance = Math.Sqrt(Math.Pow(pPolygon.Envelope.Width, 2) + Math.Pow(pPolygon.Envelope.Height, 2))

                        'Créer un nouveau point pour l'extrémité de la ligne maximum
                        pPoint = PointConstructAngleDistance(pPointColl.Point(i), dAngle, dDistance)
                        'Créer la ligne maximum vide
                        pLigneMax = New Polyline
                        pLigneMax.SpatialReference = pPolygon.SpatialReference
                        'Définir la ligne maximum
                        pLigneMax.FromPoint = pPointColl.Point(i)
                        pLigneMax.ToPoint = pPoint
                        'Interface pour extraire les points d'intersection entre la ligne maximum et le squelette
                        pTopoOp = CType(pLigneMax, ITopologicalOperator2)
                        'Extraire les points d'intersection entre la ligne maximum et le squelette
                        pGeometry = pTopoOp.Intersect(pSquelette, esriGeometryDimension.esriGeometry0Dimension)

                        'Créer un nouveau point vide
                        pPoint = New Point
                        pPoint.SpatialReference = pSquelette.SpatialReference

                        'Vérifier si aucun point d'intersection n'est trouvé
                        If Not pGeometry.IsEmpty Then
                            'Définir le multipoint
                            pMultipoint = CType(pGeometry, IMultipoint)
                            'Interface pour extraire le point d'intersection le plus proche du point du polygone trouvé
                            pProxOp = CType(pMultipoint, IProximityOperator)
                            'Extraire le point d'intersection le plus proche du point du polygone trouvé
                            pProxOp.QueryNearestPoint(pPointColl.Point(i), esriSegmentExtension.esriNoExtension, pPoint)

                            'Si aucun point d'intersection n'est trouvé
                        Else
                            'Interface utilisé pour extraire le point le plus proche
                            pProxOp = CType(pSquelette, IProximityOperator)
                            'Extraire le point d'intersection le plus proche du point du polygone trouvé
                            pProxOp.QueryNearestPoint(pPointColl.Point(i), esriSegmentExtension.esriNoExtension, pPoint)
                        End If

                        'Interface utilisé pour extraire le point le plus proche
                        pRelOp = CType(pPolygon, IRelationalOperator)
                        'Créer la ligne vide
                        pLigneMax = New Polyline
                        pLigneMax.SpatialReference = pPolygon.SpatialReference
                        'Définir la ligne
                        pLigneMax.FromPoint = pPointColl.Point(i)
                        pLigneMax.ToPoint = pPoint
                        'Vérifier si la ligne est à l'extérieur du polygone
                        If Not pRelOp.Contains(pLigneMax) Then
                            'Interface utilisé pour extraire le point le plus proche
                            pProxOp = CType(pSquelette, IProximityOperator)
                            'Extraire le point d'intersection le plus proche du point du polygone trouvé
                            pProxOp.QueryNearestPoint(pPointColl.Point(i), esriSegmentExtension.esriNoExtension, pPoint)
                        End If

                        'Créer une nouvelle ligne vide de connexion entre le polygone et le squelette
                        pPath = New Path
                        pPath.SpatialReference = pSquelette.SpatialReference
                        'Définir les sommets de la ligne de connexion
                        pPath.FromPoint = pPointColl.Point(i)
                        pPath.ToPoint = pPoint
                        'Ajouter la nouvelle ligne de connexion au squellette
                        pGeomColl.AddGeometry(pPath)
                    End If
                End If
            Next

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeometry = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
            pGeomColl = Nothing
            pPoint = Nothing
            pProxOp = Nothing
            pPath = Nothing
            pLigneMax = Nothing
            pMultipoint = Nothing
            pHitTest = Nothing
            pRingColl = Nothing
            pPntColl = Nothing
            pNewPoint = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Fonction qui permet de créer un nouveau point à partir d'un point de départ d'un angle et d'une distance.
    ''' </summary>
    ''' 
    ''' <param name="pPoint">Interface contenant le point de départ.</param>
    ''' <param name="dAngle">Angle en degrés.</param>
    ''' <param name="dDist">Distance utilisé pour construire le nouveau point.</param>
    ''' 
    ''' <returns>IPoint contenant le nouveau point calculé.</returns>
    ''' 
    Protected Friend Shared Function PointConstructAngleDistance(ByVal pPoint As IPoint, ByVal dAngle As Double, ByVal dDist As Double) As IPoint
        'Déclarer les variable de travail
        Dim pConstPoint As IConstructPoint = Nothing    'Interface utilisé pour calculer le nouveau point.
        Dim dAngleRad As Double                         'Angle en radiant utilisé pour calculer le nouveau point.

        'Créer le nouveau point vide
        PointConstructAngleDistance = New Point
        PointConstructAngleDistance.SpatialReference = pPoint.SpatialReference

        Try
            'Interface pour construire le nouveau point
            pConstPoint = CType(PointConstructAngleDistance, IConstructPoint)

            'Calculer l'angle en radiant
            dAngleRad = dAngle * 2 * Math.PI / 360

            'Contruire le nouveau point
            pConstPoint.ConstructAngleDistance(pPoint, dAngleRad, dDist)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pConstPoint = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet de connecter les extrémités des lignes du squelette en les prolongeant jusqu'au polygone.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Polygone utilisé pour créer le squelette de départ.</param>
    ''' <param name="pMultipoint">Interface contenant les sommets des extrémités de ligne du squelette.</param>
    ''' <param name="pSquelette">Polyline contenant le squelette utilisé pour connecter ses extrémités de ligne au polygone.</param>
    ''' 
    Protected Friend Shared Sub ConnecterSquelettePolygone(ByVal pPolygon As IPolygon, ByVal pMultipoint As IMultipoint, ByRef pSquelette As IPolyline)
        'Déclarer les variables de travail
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier la relation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour traiter tous les segments du squelette.
        Dim pLine As ILine = Nothing                        'Interface contenant une extrémité de ligne.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne du squellette.
        Dim dAngle As Double = 0        'Contient l'angle en radiant de l'extrémité de la ligne.
        Dim dDist As Double = 0         'Contient la distance utilisée pour créer la ligne de connexion.
        Dim pPointColl As IPointCollection = Nothing
        Dim pPointsPolygon As IMultipoint = Nothing

        Try
            'Sortir si aucune extrémité de lignes
            If pMultipoint.IsEmpty Then Exit Sub

            pPointColl = CType(pMultipoint, IPointCollection)
            pPointsPolygon = New Multipoint
            pPointsPolygon.SpatialReference = pPolygon.SpatialReference
            pPointColl = CType(pPointsPolygon, IPointCollection)
            pPointColl.AddPointCollection(CType(pPolygon, IPointCollection))
            pTopoOp = CType(pPointColl, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()
            pTopoOp = CType(pMultipoint, ITopologicalOperator2)
            pMultipoint = CType(pTopoOp.Difference(pPointsPolygon), IMultipoint)
            pPointColl = CType(pMultipoint, IPointCollection)

            'Interface pour vérifier la relation spatiale
            pRelOp = CType(pMultipoint, IRelationalOperator)

            'Interface pour traiter toutes les lignes
            pGeomColl = CType(pSquelette, IGeometryCollection)

            'Traiter toutes les lignes du squelette
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Définir la ligne du squellette à traiter
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Vérifier si la ligne est une extrémité de début de ligne
                If Not pRelOp.Disjoint(pPath.FromPoint) Then
                    'Interface pour extraire les segments
                    pSegColl = CType(pPath, ISegmentCollection)
                    'Définir la ligne
                    pLine = CType(pSegColl.Segment(0), ILine)
                    'Changer l'orientation
                    pLine.ReverseOrientation()
                    'Définir l'angle de la nouvelle ligne
                    dAngle = pLine.Angle
                    'Définir la distance de la nouvelle ligne
                    dDist = Math.Sqrt(Math.Pow(pPolygon.Envelope.Width, 2) + Math.Pow(pPolygon.Envelope.Height, 2))

                    'Ajouter la ligne de connexion
                    pGeomColl.AddGeometry(CreerLigneConnexion(pPolygon, pPath.FromPoint, dAngle, dDist))
                End If

                'Si la ligne est une extrémité de fin de ligne
                If Not pRelOp.Disjoint(pPath.ToPoint) Then
                    'Interface pour extraire les segments
                    pSegColl = CType(pPath, ISegmentCollection)
                    'Définir la ligne
                    pLine = CType(pSegColl.Segment(pSegColl.SegmentCount - 1), ILine)
                    'Définir l'angle de la nouvelle ligne
                    dAngle = pLine.Angle
                    'Définir la distance de la nouvelle ligne
                    dDist = Math.Sqrt(Math.Pow(pPolygon.Envelope.Width, 2) + Math.Pow(pPolygon.Envelope.Height, 2))

                    'Ajouter la ligne de connexion
                    pGeomColl.AddGeometry(CreerLigneConnexion(pPolygon, pPath.ToPoint, dAngle, dDist))
                End If
            Next

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pRelOp = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLine = Nothing
            pPath = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet de connecter les extrémités des lignes du squelette en les prolongeant jusqu'à une polyligne.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Polyligne utilisé pour créer le squelette de départ.</param>
    ''' <param name="pMultipoint">Interface contenant les sommets des extrémités de ligne du squelette.</param>
    ''' <param name="pSquelette">Polyline contenant le squelette utilisé pour connecter ses extrémités de ligne au polygone.</param>
    ''' 
    Protected Friend Shared Sub ConnecterSquelettePolyline(ByVal pPolyline As IPolyline, ByVal pMultipoint As IMultipoint, ByRef pSquelette As IPolyline)
        'Déclarer les variables de travail
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier la relation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pSegColl As ISegmentCollection = Nothing        'Interface pour traiter tous les segments du squelette.
        Dim pLine As ILine = Nothing                        'Interface contenant une extrémité de ligne.
        Dim pPath As IPath = Nothing                        'Interface contenant une ligne du squellette.
        Dim dAngle As Double = 0        'Contient l'angle en radiant de l'extrémité de la ligne.
        Dim dDist As Double = 0         'Contient la distance utilisée pour créer la ligne de connexion.
        Dim pPointColl As IPointCollection = Nothing
        Dim pPointsPolyline As IMultipoint = Nothing

        Try
            'Sortir si aucune extrémité de lignes
            If pMultipoint.IsEmpty Then Exit Sub

            pPointColl = CType(pMultipoint, IPointCollection)
            pPointsPolyline = New Multipoint
            pPointsPolyline.SpatialReference = pPolyline.SpatialReference
            pPointColl = CType(pPointsPolyline, IPointCollection)
            pPointColl.AddPointCollection(CType(pPolyline, IPointCollection))
            pTopoOp = CType(pPointColl, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()
            pTopoOp = CType(pMultipoint, ITopologicalOperator2)
            pMultipoint = CType(pTopoOp.Difference(pPointsPolyline), IMultipoint)
            pPointColl = CType(pMultipoint, IPointCollection)

            'Interface pour vérifier la relation spatiale
            pRelOp = CType(pMultipoint, IRelationalOperator)

            'Interface pour traiter toutes les lignes
            pGeomColl = CType(pSquelette, IGeometryCollection)

            'Traiter toutes les lignes du squelette
            For i = pGeomColl.GeometryCount - 1 To 0 Step -1
                'Définir la ligne du squellette à traiter
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Vérifier si la ligne est une extrémité de début de ligne
                If Not pRelOp.Disjoint(pPath.FromPoint) Then
                    'Interface pour extraire les segments
                    pSegColl = CType(pPath, ISegmentCollection)
                    'Définir la ligne
                    pLine = CType(pSegColl.Segment(0), ILine)
                    'Changer l'orientation
                    pLine.ReverseOrientation()
                    'Définir l'angle de la nouvelle ligne
                    dAngle = pLine.Angle
                    'Définir la distance de la nouvelle ligne
                    dDist = Math.Sqrt(Math.Pow(pPolyline.Envelope.Width, 2) + Math.Pow(pPolyline.Envelope.Height, 2))

                    'Ajouter la ligne de connexion
                    pGeomColl.AddGeometry(CreerLigneConnexion(pPolyline, pPath.FromPoint, dAngle, dDist))
                End If

                'Si la ligne est une extrémité de fin de ligne
                If Not pRelOp.Disjoint(pPath.ToPoint) Then
                    'Interface pour extraire les segments
                    pSegColl = CType(pPath, ISegmentCollection)
                    'Définir la ligne
                    pLine = CType(pSegColl.Segment(pSegColl.SegmentCount - 1), ILine)
                    'Définir l'angle de la nouvelle ligne
                    dAngle = pLine.Angle
                    'Définir la distance de la nouvelle ligne
                    dDist = Math.Sqrt(Math.Pow(pPolyline.Envelope.Width, 2) + Math.Pow(pPolyline.Envelope.Height, 2))

                    'Ajouter la ligne de connexion
                    pGeomColl.AddGeometry(CreerLigneConnexion(pPolyline, pPath.ToPoint, dAngle, dDist))
                End If
            Next

            'Interface pour simplifier les sommets en relation
            pTopoOp = CType(pSquelette, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pRelOp = Nothing
            pTopoOp = Nothing
            pGeomColl = Nothing
            pSegColl = Nothing
            pLine = Nothing
            pPath = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Fonction qui permet de retourner la ligne de connexion jusqu'à la limite d'un polygone.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Polygone utiliser pour prolonger la ligne.</param>
    ''' <param name="pPoint">Extremité de la ligne à prolonger.</param>
    ''' <param name="dAngle">Contient l'angle en radiant de l'extrémité de la ligne.</param>
    ''' <param name="dDist">Contient la distance utilisée pour créer la ligne de connexion.</param>
    ''' 
    '''<returns>IPath contenant la ligne de connexion jusqu'à la limite d'un polygone, Nothing sinon.</returns>
    ''' 
    Private Shared Function CreerLigneConnexion(ByVal pPolygon As IPolygon, ByVal pPoint As IPoint, ByVal dAngle As Double, ByVal dDist As Double) As IPath
        'Déclarer les variables de travail
        Dim pRelOp As IRelationalOperator = Nothing         'Interface pour vérifier la relation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets en relation.
        Dim pLineColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pPath As IPath = Nothing                        'Interface contenant la ligne à ajouter au squellette pour effectuer la connexion.
        Dim pPointNew As IPoint = Nothing                   'Interface ESRI contenant le sommet trouvé.
        Dim pConstPoint As IConstructPoint = Nothing        'Interface utiliser pour construire la ligne de connexion.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant la ligne de connexion.

        'La ligne est invalide par défaut
        CreerLigneConnexion = New Path
        CreerLigneConnexion.SpatialReference = pPolygon.SpatialReference

        Try
            'Définir un nouveaqu point vide
            pConstPoint = New Point
            'Interface pour définir le nouveau point
            pConstPoint.ConstructAngleDistance(pPoint, dAngle, dDist)
            'Définir le nouveau point
            pPointNew = CType(pConstPoint, IPoint)

            'Créer une nouvelle ligne
            pPolyline = New Polyline
            pPolyline.SpatialReference = pPolygon.SpatialReference
            'Interface pour ajouter des points
            pPointColl = CType(pPolyline, IPointCollection)
            'ajouter le point de début
            pPointColl.AddPoint(pPoint)
            'ajouter le point de fin
            pPointColl.AddPoint(pPointNew)

            'Interface pour découper la ligne
            pTopoOp = CType(pPolygon, ITopologicalOperator2)
            'Découper la ligne
            pLineColl = CType(pTopoOp.Intersect(pPolyline, esriGeometryDimension.esriGeometry1Dimension), IGeometryCollection)

            'Si seulement une composante est trouvée
            If pLineColl.GeometryCount = 1 Then
                'Retourner la nouvelle ligne au squellette
                CreerLigneConnexion = CType(pLineColl.Geometry(0), IPath)

                'Si plusieurs composantes sont trouvéee
            ElseIf pLineColl.GeometryCount > 1 Then
                'Interface pour vérifier la relation spatiale
                pRelOp = CType(pPoint, IRelationalOperator)

                'Traiter toutes les conposantes trouvées
                For j = 0 To pLineColl.GeometryCount - 1
                    'Définir la ligne du squellette à traiter
                    pPath = CType(pLineColl.Geometry(j), IPath)

                    'Vérifier si c'est la ligne de connexion
                    If Not (pRelOp.Disjoint(pPath.FromPoint) And pRelOp.Disjoint(pPath.ToPoint)) Then
                        'Retourner la nouvelle ligne au squellette
                        CreerLigneConnexion = CType(pLineColl.Geometry(j), IPath)
                        'Sortir de la boucle
                        Exit For
                    End If
                Next
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pRelOp = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
            pLineColl = Nothing
            pPath = Nothing
            pPointNew = Nothing
            pConstPoint = Nothing
            pPolyline = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Fonction qui permet de retourner la ligne de connexion jusqu'au sommet le prus proche de la ligne.
    ''' </summary>
    ''' 
    ''' <param name="pPolyline">Polyligone utiliser pour prolonger la ligne.</param>
    ''' <param name="pPoint">Extremité de la ligne à prolonger.</param>
    ''' <param name="dAngle">Contient l'angle en radiant de l'extrémité de la ligne.</param>
    ''' <param name="dDist">Contient la distance utilisée pour créer la ligne de connexion.</param>
    ''' 
    '''<returns>IPath contenant la ligne de connexion jusqu'à la limite d'un polygone, Nothing sinon.</returns>
    ''' 
    Private Shared Function CreerLigneConnexion(ByVal pPolyline As IPolyline, ByVal pPoint As IPoint, ByVal dAngle As Double, ByVal dDist As Double) As IPath
        'Déclarer les variables de travail
        Dim pProxOp As IProximityOperator = Nothing         'Interface pour extraire le sommet le plus proche.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le sommets en relation.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les sommets en relation.
        Dim pLineColl As IGeometryCollection = Nothing      'Interface pour extraire les géométries en relation.
        Dim pPath As IPath = Nothing                        'Interface contenant la ligne à ajouter au squellette pour effectuer la connexion.
        Dim pPointNew As IPoint = Nothing                   'Interface ESRI contenant le sommet trouvé.
        Dim pConstPoint As IConstructPoint = Nothing        'Interface utiliser pour construire la ligne de connexion.
        Dim pLigne As IPolyline = Nothing                   'Interface contenant la ligne de connexion.
        Dim pMultiPoint As IMultipoint = Nothing
        Dim pRelOp As IRelationalOperator = Nothing

        'La ligne est invalide par défaut
        CreerLigneConnexion = New Path
        CreerLigneConnexion.SpatialReference = pPolyline.SpatialReference

        Try
            'Définir un nouveaqu point vide
            pConstPoint = New Point
            'Interface pour définir le nouveau point
            pConstPoint.ConstructAngleDistance(pPoint, dAngle, dDist)
            'Définir le nouveau point
            pPointNew = CType(pConstPoint, IPoint)

            'Créer une nouvelle ligne
            pLigne = New Polyline
            pLigne.SpatialReference = pPolyline.SpatialReference
            'Interface pour ajouter des points
            pPointColl = CType(pLigne, IPointCollection)
            'ajouter le point de début
            pPointColl.AddPoint(pPoint)
            'ajouter le point de fin
            pPointColl.AddPoint(pPointNew)

            'Interface pour vérifier si les lignes s'intersecte
            pRelOp = CType(pLigne, IRelationalOperator)
            'Si les lignes s'intersecte
            If pRelOp.Disjoint(pPolyline) = False Then
                'Interface pour extraire les points d'intersection
                pTopoOp = CType(pLigne, ITopologicalOperator2)
                'Extraire les points d'intersection
                pMultiPoint = CType(pTopoOp.Intersect(pPolyline, esriGeometryDimension.esriGeometry0Dimension), IMultipoint)
                'Vérifier un point d'intersection est présent
                If pMultiPoint.IsEmpty = False Then
                    'Interface pour extraire le point d'intersection
                    pProxOp = CType(pMultiPoint, IProximityOperator)
                    'Définir le point le plus proche
                    pPointNew = pProxOp.ReturnNearestPoint(pPoint, esriSegmentExtension.esriNoExtension)

                    'Interface pour ajouter des points
                    pPointColl = CType(CreerLigneConnexion, IPointCollection)
                    'ajouter le point de début
                    pPointColl.AddPoint(pPoint)
                    'ajouter le point de fin
                    pPointColl.AddPoint(pPointNew)
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pProxOp = Nothing
            pTopoOp = Nothing
            pPointColl = Nothing
            pLineColl = Nothing
            pPath = Nothing
            pPointNew = Nothing
            pConstPoint = Nothing
            pLigne = Nothing
        End Try
    End Function
#End Region

#Region "Structures"
    ''' <summary>
    ''' Structure d'un noeud contenant le numéro de point et ses liens avec les autres numéros de point.
    ''' </summary>
    Public Structure Noeud
        Implements IComparable

        ''' <summary>
        ''' Numéro du sommet/noeud
        ''' </summary>
        Public No As Integer

        ''' <summary>
        ''' Numéro du sommet précédent
        ''' </summary>
        Public NoPrec As Integer

        ''' <summary>
        ''' Numéro du sommet suivant
        ''' </summary>
        Public NoSuiv As Integer

        ''' <summary>
        ''' Angle minimum de la droite du noeud
        ''' </summary>
        Public AngleMin As Double

        ''' <summary>
        ''' Angle maximum de la droite du noeud
        ''' </summary>
        Public AngleMax As Double

        ''' <summary>
        ''' Liste des numéros de points en lien
        ''' </summary>
        Public Liens As List(Of Integer)

        ''' <summary>
        ''' Liste des droites en lien
        ''' </summary>
        Public Droites As List(Of Droite)

        ''' <summary>
        ''' Initialiser le noeud
        ''' </summary>
        ''' 
        ''' <param name="no">Numéro du sommet/noeud</param>
        ''' 
        Public Sub New(no As Integer)
            Me.No = no
            Me.AngleMin = -1
            Me.AngleMax = -1
            Me.Liens = New List(Of Integer)
            Me.Droites = New List(Of Droite)
        End Sub

        ''' <summary>
        ''' Ajouter un numéro de point en lien avec le point du noeud.
        ''' </summary>
        ''' 
        ''' <param name="lien">Numéro du point en lien</param>
        ''' 
        Public Sub Add(lien As Integer)
            Me.Liens.Add(lien)
            Me.Liens.Sort()
        End Sub

        ''' <summary>
        ''' Ajouter une droite en lien avec le point du noeud.
        ''' </summary>
        ''' 
        ''' <param name="no">Numéro de la droite</param>
        ''' <param name="deb">Numéro du point de début de la droite</param>
        ''' <param name="fin">Numéro du point de fin de la droite</param>
        ''' <param name="angle">Angle de la droite en lien</param>
        ''' <param name="traiter">Indiquer si la droite a déjà été traitée</param>
        ''' 
        Public Sub Add(no As Integer, deb As Integer, fin As Integer, angle As Double, traiter As Boolean)
            Dim pDroite As New Droite(no, deb, fin, angle, traiter)
            Me.Droites.Add(pDroite)
            Me.Droites.Sort()
        End Sub

        ''' <summary>
        ''' Ajouter une droite en lien avec le point du noeud.
        ''' </summary>
        ''' 
        ''' <param name="no">Numéro de la droite</param>
        ''' <param name="deb">Point de début de la droite</param>
        ''' <param name="fin">Point de fin de la droite</param>
        ''' <param name="angle">Angle de la droite en lien</param>
        ''' <param name="traiter">Indiquer si la droite a déjà été traitée</param>
        ''' 
        Public Sub Add(no As Integer, deb As IPoint, fin As IPoint, angle As Double, traiter As Boolean)
            Dim pDroite As New Droite(no, deb, fin, angle, traiter)
            Me.Droites.Add(pDroite)
            Me.Droites.Sort()
        End Sub

        ''' <summary>
        ''' Implement IComparable CompareTo method to enable sorting
        ''' </summary>
        ''' <param name="obj">object Noeud</param>
        ''' <returns>value of compare</returns>
        Private Function IComparable_CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Dim other As Noeud = CType(obj, Noeud)
            If Me.No > other.No Then
                Return 1
            ElseIf Me.No < other.No Then
                Return -1
            Else
                Return 0
            End If
        End Function
    End Structure

    ''' <summary>
    ''' Déclarer la structure d'une droite.
    ''' </summary>
    ''' 
    Public Structure Droite
        Implements IComparable

        ''' <summary>
        ''' Numéro de la droite en lien
        ''' </summary>
        Public No As Integer

        ''' <summary>
        ''' Numéro du point de début de la droite
        ''' </summary>
        Public Deb As Integer

        ''' <summary>
        ''' Numéro du point de fin de la droite
        ''' </summary>
        Public Fin As Integer

        ''' <summary>
        ''' Point de début de la droite
        ''' </summary>
        Public PointDeb As IPoint

        ''' <summary>
        ''' Point de fin de la droite
        ''' </summary>
        Public PointFin As IPoint

        ''' <summary>
        ''' Angle de la droite
        ''' </summary>
        Public Angle As Double

        ''' <summary>
        ''' Indique si la droite a été traitée
        ''' </summary>
        Public Traiter As Boolean

        ''' <summary>
        ''' Initialiser la structure de la droite <see cref="Droite"/> struct
        ''' </summary>
        ''' 
        ''' <param name="no">Numéro de la droite</param>
        ''' <param name="deb">Numéro du point de début de la droite</param>
        ''' <param name="fin">Numéro du point de fin de la droite</param>
        ''' <param name="angle">Angle de la droite</param>
        ''' <param name="traiter">Indique si la droite a été traitée</param>
        ''' 
        Public Sub New(no As Integer, deb As Integer, fin As Integer, angle As Double, traiter As Boolean)
            Me.No = no
            Me.Deb = deb
            Me.Fin = fin
            Me.Angle = angle
            Me.Traiter = traiter
        End Sub

        ''' <summary>
        ''' Initialiser la structure de la droite <see cref="Droite"/> struct
        ''' </summary>
        ''' 
        ''' <param name="no">Numéro de la droite</param>
        ''' <param name="deb">Numéro du point de début de la droite</param>
        ''' <param name="fin">Numéro du point de fin de la droite</param>
        ''' <param name="angle">Angle de la droite</param>
        ''' <param name="traiter">Indique si la droite a été traitée</param>
        ''' 
        Public Sub New(no As Integer, deb As IPoint, fin As IPoint, angle As Double, traiter As Boolean)
            Me.No = no
            Me.PointDeb = deb
            Me.PointFin = fin
            Me.Angle = angle
            Me.Traiter = traiter
        End Sub

        ''' <summary>
        ''' Implémenter la méthode IComparable_CompareTo afin de pouvoir trier les droites.
        ''' </summary>
        ''' 
        ''' <param name="obj">Objet Droite</param>
        ''' 
        ''' <returns>valeur de comparaison</returns>
        ''' 
        Private Function IComparable_CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Dim other As Droite = CType(obj, Droite)
            If Me.Angle > other.Angle Then
                Return 1
            ElseIf Me.Angle < other.Angle Then
                Return -1
            ElseIf Me.Fin > other.Fin Then
                Return 1
            ElseIf Me.Fin < other.Fin Then
                Return -1
            Else
                Return 0
            End If
        End Function
    End Structure

    ''' <summary>
    ''' Déclarer la structure d'un point
    ''' </summary>
    ''' 
    Protected Friend Structure SimplePoint
        Implements IComparable

        ''' <summary>
        ''' Coordonnées X et Y du point
        ''' </summary>
        Public X As Double, Y As Double

        ''' <summary>
        ''' Initialiser la structure du point <see cref="SimplePoint"/> struct
        ''' </summary>
        ''' 
        ''' <param name="x">Coordinnée X du point</param>
        ''' <param name="y">Coordonnée Y du point</param>
        ''' 
        Public Sub New(x As Double, y As Double)
            Me.X = x
            Me.Y = y
        End Sub

        ''' <summary>
        ''' Implémenter la méthode IComparable_CompareTo afin de pouvoir trier les points.
        ''' </summary>
        ''' 
        ''' <param name="obj">object simple point</param>
        ''' 
        ''' <returns>value of compare</returns>
        ''' 
        Private Function IComparable_CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Dim other As SimplePoint = CType(obj, SimplePoint)
            If Me.X > other.X Then
                Return 1
            ElseIf Me.X < other.X Then
                Return -1
            Else
                Return 0
            End If
        End Function
    End Structure

    ''' <summary>
    ''' Déclarer la structure d'un triangle
    ''' </summary>
    ''' 
    Protected Friend Structure SimpleTriangle
        ''' <summary>
        ''' Index des trois points
        ''' </summary>
        Public A As Integer, B As Integer, C As Integer

        ''' <summary>
        ''' Point du centre d'un cercle
        ''' </summary>
        Public CircumCentre As SimplePoint

        ''' <summary>
        ''' Rayon du cercle
        ''' </summary>
        Public Radius As Double

        ''' <summary>
        ''' Initialiser la structure du traingle <see cref="SimpleTriangle"/> struct
        ''' </summary>
        ''' <param name="a">index vertex a</param>
        ''' <param name="b">index vertex b</param>
        ''' <param name="c">index vertex c</param>
        ''' <param name="circumcentre">center of triangle</param>
        ''' <param name="radius">radius of triangle</param>
        Public Sub New(a As Integer, b As Integer, c As Integer, circumcentre As SimplePoint, radius As Double)
            Me.A = a
            Me.B = b
            Me.C = c
            Me.CircumCentre = circumcentre
            Me.Radius = radius
        End Sub
    End Structure

    ''' <summary>
    ''' Calculer le centre d'un cercle et son rayon à partir de troix points.
    ''' </summary>
    ''' 
    ''' <param name="p1">Premier point</param>
    ''' <param name="p2">Deuxième point</param>
    ''' <param name="p3">Troisième point</param>
    ''' <param name="circumCentre">center of circle</param>
    ''' <param name="radius">value of radius</param>
    ''' 
    Protected Friend Shared Sub CalculateCircumcircle(p1 As SimplePoint, p2 As SimplePoint, p3 As SimplePoint, ByRef circumCentre As SimplePoint, ByRef radius As Double)
        ' Calculate the length of each side of the triangle
        Dim a As Double = Distance(p2, p3)
        ' side a is opposite point 1
        Dim b As Double = Distance(p1, p3)
        ' side b is opposite point 2 
        Dim c As Double = Distance(p1, p2)
        ' side c is opposite point 3
        ' Calculate the radius of the circumcircle
        Dim area As Double = Math.Abs(CDbl((p1.X * (p2.Y - p3.Y)) + (p2.X * (p3.Y - p1.Y)) + (p3.X * (p1.Y - p2.Y))) / 2)
        radius = a * b * c / (4 * area)

        ' Define area coordinates to calculate the circumcentre
        Dim pp1 As Double = Math.Pow(a, 2) * (Math.Pow(b, 2) + Math.Pow(c, 2) - Math.Pow(a, 2))
        Dim pp2 As Double = Math.Pow(b, 2) * (Math.Pow(c, 2) + Math.Pow(a, 2) - Math.Pow(b, 2))
        Dim pp3 As Double = Math.Pow(c, 2) * (Math.Pow(a, 2) + Math.Pow(b, 2) - Math.Pow(c, 2))

        ' Normalise
        Dim t1 As Double = pp1 / (pp1 + pp2 + pp3)
        Dim t2 As Double = pp2 / (pp1 + pp2 + pp3)
        Dim t3 As Double = pp3 / (pp1 + pp2 + pp3)

        ' Convert to Cartesian
        Dim x As Double = (t1 * p1.X) + (t2 * p2.X) + (t3 * p3.X)
        Dim y As Double = (t1 * p1.Y) + (t2 * p2.Y) + (t3 * p3.Y)

        circumCentre = New SimplePoint(x, y)
    End Sub

    ''' <summary>
    ''' Calculer la distance entre deux points.
    ''' </summary>
    ''' 
    ''' <param name="p1">Premier point</param>
    ''' <param name="p2">Deuxième point</param>
    ''' 
    ''' <returns>Distance entre les deux points</returns>
    ''' 
    Protected Friend Shared Function Distance(p1 As SimplePoint, p2 As SimplePoint) As Double
        Distance = Math.Sqrt(Math.Pow(p2.X - p1.X, 2) + Math.Pow(p2.Y - p1.Y, 2))
    End Function

    ''' <summary>
    ''' Calculer l'angle en degrés entre deux points selon l'axe des X.
    ''' L'angle calculé doit être supérieure à l'angle de base.
    ''' </summary>
    ''' 
    ''' <param name="p1">Premier point</param>
    ''' <param name="p2">Deuxième point</param>
    ''' <param name="angleBase">Angle de base utilisée pour calculer l'angle résultant</param>
    ''' 
    ''' <returns>Angle en degrés entre deux points</returns>
    ''' 
    Protected Friend Shared Function Angle(p1 As IPoint, p2 As IPoint, Optional angleBase As Double = 0) As Double
        Dim radiant As Double = 0

        'Si la droite est nulle, 90 ou 270 degrés
        If p1.X = p2.X Then
            'Si la droite est nulle
            If p1.Y = p2.Y Then
                'Mettre l'angle à 0 degrés
                Angle = 0

                'Si l'angle est à 90 degrés
            ElseIf p1.Y < p2.Y Then
                'Mettre l'angle à 90 degrés
                Angle = 90

                'Si l'angle est à 270 degrés
            Else
                'Mettre l'angle à 270 degrés
                Angle = 270
            End If

            'Si la droite est 0 ou 180 degrés
        ElseIf p1.Y = p2.Y Then
            'Si l'angle est à 0 degrés
            If p1.X < p2.X Then
                'Mettre l'angle à 0 degrés
                Angle = 0

                'Si l'angle est à 180 degrés
            Else
                'Mettre l'angle à 180 degrés
                Angle = 180
            End If

            'Si la droite est dans le quadrant 1
        ElseIf p1.X < p2.X And p1.Y < p2.Y Then
            'Calculer l'angle en radiant
            radiant = Math.Atan((p2.Y - p1.Y) / (p2.X - p1.X))
            'Calculer l'angle en degrés
            Angle = radiant * (180 / Math.PI)

            'Si la droite est dans le quadrant 2
        ElseIf p1.X > p2.X And p1.Y < p2.Y Then
            'Calculer l'angle en radiant
            radiant = Math.Atan((p1.X - p2.X) / (p2.Y - p1.Y))
            'Calculer l'angle en degrés
            Angle = radiant * (180 / Math.PI) + 90

            'Si la droite est dans le quadrant 3
        ElseIf p1.X > p2.X And p1.Y > p2.Y Then
            'Calculer l'angle en radiant
            radiant = Math.Atan((p1.Y - p2.Y) / (p1.X - p2.X))
            'Calculer l'angle en degrés
            Angle = radiant * (180 / Math.PI) + 180

            'Si la droite est dans le quadrant 4
        ElseIf p1.X < p2.X And p1.Y > p2.Y Then
            'Calculer l'angle en radiant
            radiant = Math.Atan((p2.X - p1.X) / (p1.Y - p2.Y))
            'Calculer l'angle en degrés
            Angle = radiant * (180 / Math.PI) + 270
        End If

        'L'angle calculée doit être plus grand que l'angle de base
        If Angle < angleBase Then Angle = Angle + 360
    End Function
#End Region
End Class

''' <summary>
''' Classe qui permet de créer et gérer un Diagramme de Voronoi.
''' </summary>
Public Class DiagrammeVoronoi
    Inherits TriangulationDelaunay

#Region "Routines et fonctions publiques"
    ''' <summary>
    ''' Function qui permet de créer et retourner les lignes du squelette d'un polygone dans une Polyline.
    ''' Le squelette est créé à partir des lignes du diagramme de Voronoi du Polygone
    ''' Le diagramme de Voronoi est créé à partir de la liste des triangles de Delaunay.
    ''' La liste des triangles de Delaunay est créée à partir de tous les sommets non duppliqués et triés de la géométrie spécifiée.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Polygone utilisée pour créer le squelette.</param>
    '''<param name="pPointsConnexion">Interface contenant les points de connexion entre le polygon et les éléments en relation.</param>
    '''<param name="dDistLat"> Distance latérale utilisée pour éliminer des sommets en trop.</param>
    '''<param name="dDistMin"> Distance minimum utilisée pour ajouter des sommets.</param>
    '''<param name="dLongMin"> Longueur minimale utilisée pour éliminer des lignes du squelette trop petites.</param>
    ''' 
    ''' <returns>IPolyline contenant les lignes du diagramme de Voronoi.</returns>
    ''' 
    Public Shared Function CreerSquelettePolygoneVoronoi(ByVal pPolygon As IPolygon, ByVal pPointsConnexion As IMultipoint, _
                                                         Optional ByVal dDistLat As Double = 1.5, _
                                                         Optional ByVal dDistMin As Double = 10, _
                                                         Optional ByVal dLongMin As Double = 50) As IPolyline
        'Déclarer les variables de travail
        Dim pLigne As IPolyline = Nothing           'Interface contenant une ligne du diagramme de Voronoi.
        Dim pPolyligne As IPolyline = Nothing       'Interface contenant les lignes du diagramme de Voronoi.

        'Définir la valeur par défaut
        CreerSquelettePolygoneVoronoi = Nothing

        Try
            'Densifer les sommets du polygone
            pPolygon.Densify(dDistMin, 0)

            'Définir les lignes du diagramme de Voronoi sans découper
            pPolyligne = CreerListeLignes(pPolygon, False)

            'Définir le squelette de base
            CreerSquelettePolygoneVoronoi = CreerSqueletteBaseVoronoi(pPolygon, pPolyligne)

            'Connecter le squelette aux sommets en relation
            'Call ConnecterSquelettePointsConnexion(pPointsConnexion, dDistMin, dLongMin, CreerSquelettePolygoneVoronoi)
            Call ConnecterSquelettePointsConnexion(pPointsConnexion, CreerSquelettePolygoneVoronoi)
            'Call ConnecterSquelettePointsConnexion(pPolygon, pPointsConnexion, CreerSquelettePolygoneVoronoi)

            'Enlever les extrémités de lignes superflux dans le squelette
            Call EnleverExtremiteLigneVoronoi(pPointsConnexion, CreerSquelettePolygoneVoronoi, dDistLat, dLongMin)

            'Connecter le squelette au polygone
            Call ConnecterSquelettePolygone(pPolygon, pPointsConnexion, CreerSquelettePolygoneVoronoi)

            'Enlever les sommets en trop afin d'enlever les extrémités en ligne droite
            'CreerSquelettePolygoneVoronoi.Generalize(dDistLat)

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLigne = Nothing
            pPolyligne = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function qui permet de créer et retourner les lignes du diagramme de Voronoi dans une Polyline pour une géométrie.
    ''' Le diagramme de Voronoi est créé à partir de la liste des triangles de Delaunay.
    ''' La liste des triangles de Delaunay est créée à partir de tous les sommets non duppliqués et triés de la géométrie spécifiée.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour créer le diagramme de Voronoi à partir de ses sommets.</param>
    ''' <param name="bDecouper">Indique si on doit découper les lignes du diagramme de Voronoi selon la zone de traitement.</param>
    ''' 
    ''' <returns>IPolyline contenant les lignes du diagramme de Voronoi, Nothing sinon.</returns>
    ''' 
    Public Shared Function CreerListeLignes(pGeometry As IGeometry, Optional ByVal bDecouper As Boolean = True) As IPolyline
        'Déclarer les variables de travail
        Dim pGeometryBag As IGeometryBag = Nothing          'Interface contenant les lignes du diagramme de Voronoi.
        Dim plignesColl As IGeometryCollection = Nothing    'Interface pour ajouter les lignes du diagramme de Voronoi.
        Dim pPolyline As IPolyline = Nothing                'Interface contenant une ligne
        Dim pZone As IGeometry = Nothing                    'Interface contenant la zone de traitement (Enveloppe ou Polygone de la géométrie).
        Dim pMultipoint As IPointCollection = Nothing       'Interface contenant les sommets d'un triangle.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le polygone de Voronoi (ConvexHull).
        Dim pPolygonVoronoi As IGeometry = Nothing          'Interface contenant le polygone de Voronoi.
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pListeTriangles As List(Of SimpleTriangle) = Nothing 'Liste des triangles de Delaunay pour créer le diagramme de Voronoi.

        'Définir la valeur par défaut
        CreerListeLignes = New Polyline
        CreerListeLignes.SpatialReference = pGeometry.SpatialReference

        Try
            'Créer la liste des sommets
            pListeSommets = CreerListeSommets(pGeometry)

            'Créer la liste des triangles de Delaunay pour créer le diagramme de Voronoi
            pListeTriangles = CreerListeTriangles(pGeometry.Envelope, pListeSommets)

            'Créer un nouveau Bag vide
            pGeometryBag = New GeometryBag
            pGeometryBag.SpatialReference = pGeometry.SpatialReference

            'Initialiser le GeometryBag contenant les lignes du diagramme de Voronoi
            plignesColl = CType(pGeometryBag, IGeometryCollection)

            'Traiter tous les sommets
            For i = 0 To pListeSommets.Count - 1
                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pMultipoint = New Multipoint

                'Traiter tous les triangles
                For Each triangle As SimpleTriangle In pListeTriangles
                    'Si le triangle intersecte ce point
                    If triangle.A = i OrElse triangle.B = i OrElse triangle.C = i Then
                        'Ajouter le point du centre du cercle
                        pMultipoint.AddPoint(New PointClass() With {.X = triangle.CircumCentre.X, .Y = triangle.CircumCentre.Y})
                    End If
                Next

                'Vérifier le nombre de sommets
                If pMultipoint.PointCount > 2 Then
                    'Interface pour créer le polygone de Voronoi à partir du résultat ConvexHull du Multipoint
                    pTopoOp = TryCast(pMultipoint, ITopologicalOperator2)
                    'Créer le polygone de Voronoi à partir du résultat ConvexHull du multipoint
                    pPolygonVoronoi = pTopoOp.ConvexHull()

                    'Interface pour extraire la limite du polygone de Voronoi
                    pTopoOp = TryCast(pPolygonVoronoi, ITopologicalOperator2)
                    'Extraire la limite du polygone de Voronoi
                    pPolyline = CType(pTopoOp.Boundary, IPolyline)

                    'Si la limite n'est pas invalide ou vide
                    If (pPolyline IsNot Nothing) AndAlso (Not pPolyline.IsEmpty) Then
                        'Ajouter la ligne du polygone de Voronoi
                        plignesColl.AddGeometry(pPolyline)
                    End If
                End If
            Next

            'Créer une polyligne vide
            pPolyline = New Polyline
            pPolyline.SpatialReference = pGeometry.SpatialReference

            'Interface pour construire la squelette sans dupplication et découper le polygone de Voronoi selon la zone de traitement
            pTopoOp = TryCast(CreerListeLignes, ITopologicalOperator2)
            'Construire la squelette sans dupplication
            pTopoOp.ConstructUnion(CType(plignesColl, IEnumGeometry))

            'Vérifier si on doit découper les lignes du diagramme de Voronoi
            If bDecouper Then
                'Par défaut, la Zone de traitement est l'enveloppe de la géométrie
                pZone = pGeometry.Envelope

                'Vérifier si la géométrie est un polygone
                If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                    'La zone de traitement est le polygone de la géométrie
                    pZone = pGeometry
                End If

                'Découper les lignes du polygone de Voronoi selon la zone de traitement
                CreerListeLignes = CType(pTopoOp.Intersect(pZone, esriGeometryDimension.esriGeometry1Dimension), IPolyline)
            End If

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeometryBag = Nothing
            pPolyline = Nothing
            plignesColl = Nothing
            pMultipoint = Nothing
            pZone = Nothing
            pTopoOp = Nothing
            pPolygonVoronoi = Nothing
            pListeSommets = Nothing
            pListeTriangles = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function qui permet de créer et retourner les polygones du diagramme de Voronoi dans un GeometryBag pour une géométrie.
    ''' Le diagramme de Voronoi est créé à partir de la liste des triangles de Delaunay.
    ''' La liste des triangles de Delaunay est créée à partir de tous les sommets non duppliqués et triés de la géométrie spécifiée.
    ''' </summary>
    ''' 
    ''' <param name="pGeometry">Géométrie utilisée pour créer le diagramme de Voronoi à partir de ses sommets.</param>
    ''' <param name="bDecouper">Indique si on doit découper les lignes du diagramme de Voronoi selon la zone de traitement.</param>
    ''' 
    ''' <returns>IGeometryBag contenant les polygones du diagramme de Voronoi, Nothing sinon.</returns>
    ''' 
    Public Shared Function CreerBagPolygonesDiagrammeVoronoi(pGeometry As IGeometry, Optional ByVal bDecouper As Boolean = True) As IGeometryBag
        'Déclarer les variables de travail
        Dim pGeometryBag As IGeometryCollection = Nothing   'Interface contenant les polygones du diagramme de Voronoi.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour extraire les sommets d'une géométrie.
        Dim pZone As IGeometry = Nothing                    'Interface contenant la zone de traitement.
        Dim pMultipoint As IPointCollection = Nothing       'Interface contenant les sommets d'un triangle.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le polygone de Voronoi (ConvexHull).
        Dim pPolygonVoronoi As IGeometry = Nothing          'Interface contenant le polygone de Voronoi.
        Dim pResultat As IGeometry = Nothing                'Interface contenant le polygon de Voronoi coupé.
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pListeTriangles As List(Of SimpleTriangle) = Nothing 'Liste des triangles de Delaunay pour créer le diagramme de Voronoi.

        'Définir la valeur par défaut
        CreerBagPolygonesDiagrammeVoronoi = Nothing

        Try
            'Créer la liste des sommets
            pListeSommets = CreerListeSommets(pGeometry)

            'Créer la liste des triangles de Delaunay pour créer le diagramme de Voronoi
            pListeTriangles = CreerListeTriangles(pGeometry.Envelope, pListeSommets)

            'Par défaut, la Zone de traitement est l'enveloppe de la géométrie
            pZone = pGeometry.Envelope
            'Vérifier si la géométrie est un polygone
            If pGeometry.GeometryType = esriGeometryType.esriGeometryPolygon Then
                'La zone de traitement est le polygone de la géométrie
                pZone = pGeometry
            End If

            'Initialiser le GeometryBag contenant les polygones du diagramme de Voronoi
            pGeometryBag = New GeometryBagClass()

            'Traiter tous les sommets
            For i = 0 To pListeSommets.Count - 1
                'Initiliaser un nouveau MultiPoint utilisé pour créer le polygone voronoi
                pMultipoint = New MultipointClass()

                'Traiter tous les triangles
                For Each triangle As SimpleTriangle In pListeTriangles
                    'Si le triangle intersecte ce point
                    If triangle.A = i OrElse triangle.B = i OrElse triangle.C = i Then
                        'Ajouter le point du centre du cercle
                        pMultipoint.AddPoint(New PointClass() With {.X = triangle.CircumCentre.X, .Y = triangle.CircumCentre.Y})
                    End If
                Next

                'Interface pour créer le polygone de Voronoi à partir du résultat ConvexHull du Multipoint
                pTopoOp = TryCast(pMultipoint, ITopologicalOperator2)
                'Créer le polygone de Voronoi à partir du résultat ConvexHull du multipoint
                pPolygonVoronoi = pTopoOp.ConvexHull()

                'Vérifier si on doit découper
                If bDecouper Then
                    'Interface pour découper le polygone selon la zone de traitement
                    pTopoOp = TryCast(pPolygonVoronoi, ITopologicalOperator2)
                    'Découper le polygone selon la zone de traitement
                    pResultat = pTopoOp.Intersect(pZone, esriGeometryDimension.esriGeometry2Dimension)
                    'Si le résultat est valide
                    If (pResultat IsNot Nothing) AndAlso (Not pResultat.IsEmpty) Then
                        'Ajouter le polygone de Voronoi coupé dans le Bag
                        pGeometryBag.AddGeometry(TryCast(pResultat, IGeometry))
                    End If

                    'Si on ne doit pas découper
                Else
                    'Ajouter le polygone de Voronoi coupé dans le Bag
                    pGeometryBag.AddGeometry(TryCast(pPolygonVoronoi, IGeometry))
                End If
            Next

            'Définir le GeometryBag de retour contenant les polygones de Voronoi
            CreerBagPolygonesDiagrammeVoronoi = TryCast(pGeometryBag, IGeometryBag)
            'Définir la référence spatiale du GeometryBag
            CreerBagPolygonesDiagrammeVoronoi.SpatialReference = pGeometry.SpatialReference

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeometryBag = Nothing
            pMultipoint = Nothing
            pPointColl = Nothing
            pZone = Nothing
            pTopoOp = Nothing
            pPolygonVoronoi = Nothing
            pResultat = Nothing
            pListeSommets = Nothing
            pListeTriangles = Nothing
        End Try
    End Function
#End Region

#Region "Routines et fonctions privées"
    ''' <summary>
    ''' Function qui permet de créer et retourner les lignes du squelette de base d'un polygone dans une Polyline.
    ''' La fonction utilise les lignes du diagramme de Voronoi pour créer le squelette.
    ''' </summary>
    ''' 
    ''' <param name="pPolygon">Polygone utilisée pour créer le squelette.</param>
    ''' <param name="pPolyligne">Interface contenant les lignes du diagramme de Voronoi.</param>
    ''' 
    ''' <returns>IPolyline contenant le squelette de base du polyone.</returns>
    ''' 
    Private Shared Function CreerSqueletteBaseVoronoi(ByVal pPolygon As IPolygon, ByVal pPolyligne As IPolyline) As IPolyline
        'Déclarer les variables de travail
        Dim pLigne As IPolyline = Nothing                   'Interface contenant une ligne du diagramme de Voronoi.
        Dim pRelOpNxM As IRelationalOperatorNxM = Nothing   'Interface utilisé pour traiter la relation spatiale.
        Dim pRelResult As IRelationResult = Nothing         'Interface contenant le résultat du traitement de la relation spatiale.
        Dim pGeomBag As IGeometryBag = Nothing              'Interface contenant des géométries.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface pour extraire les composantes d'une géométrie.
        Dim pGeomSelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments à traiter.
        Dim pGeomRelColl As IGeometryCollection = Nothing   'Interface contenant les géométries des éléments en relation.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le polygone de Voronoi (ConvexHull).
        Dim iSel As Integer = -1            'Numéro de séquence de la géométrie traitée.
        Dim iRel As Integer = -1            'Numéro de séquence de la géométrie en relation.

        'Définir la valeur par défaut
        CreerSqueletteBaseVoronoi = New Polyline
        CreerSqueletteBaseVoronoi.SpatialReference = pPolygon.SpatialReference

        Try
            'Interface pour définir les géométries de base
            pGeomBag = New GeometryBag
            pGeomBag.SpatialReference = pPolygon.SpatialReference
            'Ajouter le polygone de base
            pGeomSelColl = CType(pGeomBag, IGeometryCollection)
            pGeomSelColl.AddGeometry(pPolygon)

            'Interface pour traiter les relations spatiales
            pRelOpNxM = CType(pGeomSelColl, IRelationalOperatorNxM)

            'Interface pour définir les géométries en relation
            pGeomBag = New GeometryBag
            pGeomBag.SpatialReference = pPolygon.SpatialReference
            pGeomRelColl = CType(pGeomBag, IGeometryCollection)

            'Interface pour extraire les lignes du diagramme de Voronoi
            pGeomColl = CType(pPolyligne, IGeometryCollection)
            'Traiter toutes les lignes du diagramme de Voronoi
            For i = 0 To pGeomColl.GeometryCount - 1
                'Créer une nouvelle ligne
                pLigne = New Polyline
                pLigne.SpatialReference = pPolygon.SpatialReference
                'Interface pour créer une ligne
                pLigneColl = CType(pLigne, IGeometryCollection)
                'Ajouter les composantes de la ligne
                pLigneColl.AddGeometry(pGeomColl.Geometry(i))
                'Ajouter la ligne dans le Bag des relations
                pGeomRelColl.AddGeometry(pLigne)
            Next

            'Interface pour ajouter les lignes du squelette
            pLigneColl = CType(CreerSqueletteBaseVoronoi, IGeometryCollection)

            'Traiter la relation spatiale 'Contient'
            pRelResult = pRelOpNxM.Contains(CType(pGeomRelColl, IGeometryBag))

            'Traiter toutes les lignes
            For i = 0 To pRelResult.RelationElementCount - 1
                'Extraire la géométrie traitée (left) et celle en relation (right) qui respectent la relation spatiale
                pRelResult.RelationElement(i, iSel, iRel)

                'Ajouter la lignes dans le squellette qui respecte la relation 'Contient'
                pLigneColl.AddGeometryCollection(CType(pGeomRelColl.Geometry(iRel), IGeometryCollection))
            Next

            'Interface pour simplifier le résultat
            pTopoOp = TryCast(CreerSqueletteBaseVoronoi, ITopologicalOperator2)
            pTopoOp.IsKnownSimple_2 = False
            pTopoOp.Simplify()

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pLigne = Nothing
            pRelOpNxM = Nothing
            pRelResult = Nothing
            pGeomBag = Nothing
            pGeomColl = Nothing
            pLigneColl = Nothing
            pGeomSelColl = Nothing
            pGeomRelColl = Nothing
            pTopoOp = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Routine qui permet d'enlever les lignes inutiles du squelette d'un polygone créé à partir du diagramme de Voronoi.
    ''' Les extrémités qui possède seulement 2 sommets ou qui possède une longueur inférieure à la longueur minimale seront enlever du squelette. 
    ''' </summary>
    ''' 
    ''' <param name="pMultiPoint">Sommets pour lesquels le squelette doit être connecté.</param>
    ''' <param name="pSquelette">Lignes contenant le squelette de départ.</param>
    ''' <param name="dDistLat">Contient la distance latérale filtrer les sommets superflux.</param>
    ''' <param name="dLongMin">Contient la longueur minimale des lignes à conserver dans le squelette.</param>
    ''' 
    Private Shared Sub EnleverExtremiteLigneVoronoi(ByRef pMultipoint As IMultipoint, ByRef pSquelette As IPolyline, _
                                                    Optional ByVal dDistLat As Double = 1.5, _
                                                    Optional ByVal dLongMin As Double = 50)
        'Déclarer les variables de travail
        Dim pPolyline As IPolyline = Nothing                'Interface contenant les lignes du nouveau squelette.
        Dim pTopoOp As ITopologicalOperator2 = Nothing      'Interface pour extraire le polygone de Voronoi (ConvexHull).
        Dim pLigneColl As IGeometryCollection = Nothing     'Interface contenant une ligne.
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les lignes d'extrémité.
        Dim pListeNo As List(Of Integer) = Nothing          'Liste des numéros de sommet non dupliqués de la géométrie.
        Dim pListeNb As List(Of Integer) = Nothing          'Liste du nombre de sommets dupliqués de la géométrie.
        Dim pListeLignes As List(Of Integer) = Nothing      'Liste des lignes de la géométrie.
        Dim pListeSommets As List(Of SimplePoint) = Nothing 'Liste des sommets non dupliqués de la géométrie.
        Dim pPath As IPath = Nothing                        'Interface contenant une composante de ligne.
        Dim pPoint As IPoint = Nothing                      'Interface ESRI contenant le sommet non connecté.
        Dim iNoLigne As Integer = -1                        'Contient le numéro de la ligne.

        Try
            'Traiter tout
            Do
                'Interface pour traiter toutes les lignes
                pGeomColl = CType(pSquelette, IGeometryCollection)

                'Enlever les sommets en trop afin d'enlever les extrémités en ligne droite
                pSquelette.Generalize(dDistLat)

                'Ajouter les sommets pour conserver les longs extrémités
                pSquelette.Densify((pSquelette.Length / pGeomColl.GeometryCount), 0)

                'Identifier tous les sommets qui sont une extrémité de ligne du squelette
                Call IdentifierExtremitesLignes(pMultipoint, pSquelette, pListeNo, pListeNb, pListeSommets)

                'Identifier toutes les lignes qui sont une extrémité à enlever
                '---------------------------------------------------------
                'Initialiser la liste des lignes d'extrémité
                pListeLignes = New List(Of Integer)

                'Traiter tous les sommets
                For i = 0 To pListeSommets.Count - 1
                    'Vérifier si le sommet est une extrémité
                    If Not pListeNb.Contains(i) Then
                        'Définir le no de la ligne du sommet traité
                        iNoLigne = pListeNo.Item(i)
                        'Définir la ligne
                        pPath = CType(pGeomColl.Geometry(iNoLigne), IPath)
                        'Ajouter la ligne
                        pPointColl = CType(pPath, IPointCollection)
                        'Vérifier si seulement deux sommets
                        If pPointColl.PointCount = 2 Or pPath.Length < dLongMin Then
                            'Ajouter le no de ligne qui est une extremité
                            If Not pListeLignes.Contains(iNoLigne) Then pListeLignes.Add(iNoLigne)
                        End If
                    End If
                Next

                'Créer le nouveau squelette sans les lignes qui ne sont pas une extrémité à enlever
                '---------------------------------------------------------
                'Créer une nouvelle polyligne vide
                pPolyline = New Polyline
                pPolyline.SpatialReference = pSquelette.SpatialReference

                'Interface pour ajouter les lignes qui ne sont pas les extrémités de lignes
                pLigneColl = CType(pPolyline, IGeometryCollection)

                'Traiter tous les composantes de géométrie
                For i = 0 To pGeomColl.GeometryCount - 1
                    'Vérifier si la ligne n'est pas une extrémité à enlever
                    If Not pListeLignes.Contains(i) Then
                        'Ajouter la ligne
                        pLigneColl.AddGeometry(pGeomColl.Geometry(i))
                    End If
                Next

                'Vérifier si la géométrie n'est pas vide
                If pLigneColl.GeometryCount > 0 Then
                    'Redéfinir la polyligne d'entrée
                    pSquelette = pPolyline
                    'Interface pour simplifier le résultat
                    pTopoOp = TryCast(pSquelette, ITopologicalOperator2)
                    pTopoOp.IsKnownSimple_2 = False
                    pTopoOp.Simplify()
                End If

                'Traiter tant qu'il y a des changements
            Loop While pGeomColl.GeometryCount <> pLigneColl.GeometryCount And pLigneColl.GeometryCount > 0

            'Définir les sommets du squelette non connectés
            '---------------------------------------------------------
            'Vider les sommets non connectés
            pMultipoint.SetEmpty()

            'Interface pour ajouter les sommets non-connectés
            pPointColl = CType(pMultipoint, IPointCollection)

            'Traiter tous les sommets
            For i = 0 To pListeSommets.Count - 1
                'Vérifier si le sommet est une extrémité non connecté
                If Not pListeNb.Contains(i) Then
                    'Créer le point à partir du SimplePoint
                    pPoint = New Point With {.X = pListeSommets.Item(i).X, .Y = pListeSommets.Item(i).Y}
                    pPoint.SpatialReference = pSquelette.SpatialReference

                    'Ajouter le point non connecté
                    pPointColl.AddPoint(pPoint)
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pPolyline = Nothing
            pTopoOp = Nothing
            pLigneColl = Nothing
            pGeomColl = Nothing
            pPointColl = Nothing
            pListeNo = Nothing
            pListeNb = Nothing
            pListeLignes = Nothing
            pListeSommets = Nothing
            pPath = Nothing
            pPoint = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Routine qui permet d'identifier tous les sommets qui sont une extrémité de ligne du squelette.
    ''' </summary>
    ''' 
    ''' <param name="pMultiPoint">Sommets pour lesquels le squelette doit être connecté.</param>
    ''' <param name="pSquelette">Lignes contenant le squelette de départ.</param>
    ''' <param name="pListeNo">Contient la liste des numéros de sommet non dupliqués de la géométrie.</param>
    ''' <param name="pListeNb">Contient la liste du nombre de sommets dupliqués de la géométrie.</param>
    ''' <param name="pListeSommets">Contient la liste des sommets qui sont une extrémité de ligne du squelette.</param>
    ''' 
    Private Shared Sub IdentifierExtremitesLignes(ByRef pMultipoint As IMultipoint, ByRef pSquelette As IPolyline, _
                                                  ByRef pListeNo As List(Of Integer), ByRef pListeNb As List(Of Integer), _
                                                  ByRef pListeSommets As List(Of SimplePoint))
        'Déclarer les variables de travail
        Dim pGeomColl As IGeometryCollection = Nothing      'Interface pour extraire les composantes d'une géométrie.
        Dim pPointColl As IPointCollection = Nothing        'Interface pour ajouter les lignes d'extrémité.
        Dim pSimplePoint As SimplePoint = Nothing           'Sommet utilisé pour créer le diagramme de Voronoi.
        Dim pPath As IPath = Nothing                        'Interface contenant une composante de ligne.
        Dim iNoLigne As Integer = -1                        'Contient le numéro de la ligne.

        Try
            'Initialiser la liste des nos
            pListeNo = New List(Of Integer)

            'Initialiser la liste des nos
            pListeNb = New List(Of Integer)

            'Initialiser la liste des sommets
            pListeSommets = New List(Of SimplePoint)

            'Interface pour traiter toutes les lignes
            pGeomColl = CType(pSquelette, IGeometryCollection)

            'Traiter toutes les lignes
            For i = 0 To pGeomColl.GeometryCount - 1
                'Définir la ligne
                pPath = CType(pGeomColl.Geometry(i), IPath)

                'Créer un nouveau sommet
                pSimplePoint = New SimplePoint(pPath.FromPoint.X, pPath.FromPoint.Y)
                'Vérifier si le sommet est déjà présent 
                If pListeSommets.Contains(pSimplePoint) Then
                    'Définir le no de la ligne qui n est pas une extremité
                    iNoLigne = pListeSommets.IndexOf(pSimplePoint)
                    'Ajouter le no du sommet qui n'est pas une extremité
                    If Not pListeNb.Contains(iNoLigne) Then pListeNb.Add(iNoLigne)
                Else
                    'Ajouter le sommet dans la liste des sommets
                    pListeSommets.Add(pSimplePoint)
                    'Ajouter le no de ligne du sommet
                    pListeNo.Add(i)
                End If

                'Créer un nouveau sommet
                pSimplePoint = New SimplePoint(pPath.ToPoint.X, pPath.ToPoint.Y)
                'Vérifier si le sommet est déjà présent 
                If pListeSommets.Contains(pSimplePoint) Then
                    'Définir le no de la ligne qui n est pas une extremité
                    iNoLigne = pListeSommets.IndexOf(pSimplePoint)
                    'Ajouter le no du sommet qui n'est pas une extremité
                    If Not pListeNb.Contains(iNoLigne) Then pListeNb.Add(iNoLigne)
                Else
                    'Ajouter le sommet dans la liste des sommets
                    pListeSommets.Add(pSimplePoint)
                    'Ajouter le no de ligne du sommet
                    pListeNo.Add(i)
                End If
            Next

            'Ajouter les sommets en relation dans la liste des sommets
            '---------------------------------------------------------
            'Interface pour extraire tous les sommets en relation
            pPointColl = CType(pMultipoint, IPointCollection)
            'Traiter tous les sommets en relation
            For i = 0 To pPointColl.PointCount - 1
                'Créer un nouveau sommet
                pSimplePoint = New SimplePoint(pPointColl.Point(i).X, pPointColl.Point(i).Y)
                'Définir le no de la ligne qui n est pas une extremité
                iNoLigne = pListeSommets.IndexOf(pSimplePoint)
                'Ajouter le no de ligne qui n'est pas une extremité
                pListeNb.Add(iNoLigne)
            Next

        Catch ex As Exception
            Throw
        Finally
            'Vider la mémoire
            pGeomColl = Nothing
            pPointColl = Nothing
            pSimplePoint = Nothing
            pPath = Nothing
        End Try
    End Sub
#End Region
End Class