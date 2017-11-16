<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dckMenuGeneralisation
    Inherits System.Windows.Forms.UserControl

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnInitialiser = New System.Windows.Forms.Button()
        Me.tabEdition = New System.Windows.Forms.TabControl()
        Me.pgeDimensionMinimale = New System.Windows.Forms.TabPage()
        Me.grpGeneralisationExterieure = New System.Windows.Forms.GroupBox()
        Me.lblLongueurExterieure = New System.Windows.Forms.Label()
        Me.cboLongueurExterieure = New System.Windows.Forms.ComboBox()
        Me.lblLargeurExterieure = New System.Windows.Forms.Label()
        Me.cboLargeurExterieure = New System.Windows.Forms.ComboBox()
        Me.grpDistance = New System.Windows.Forms.GroupBox()
        Me.lblDistanceDensifier = New System.Windows.Forms.Label()
        Me.cboDistanceDensifier = New System.Windows.Forms.ComboBox()
        Me.lblDistanceLaterale = New System.Windows.Forms.Label()
        Me.cboDistanceLaterale = New System.Windows.Forms.ComboBox()
        Me.grpLongueur = New System.Windows.Forms.GroupBox()
        Me.lblLongueurLigne = New System.Windows.Forms.Label()
        Me.cboLongueurLigne = New System.Windows.Forms.ComboBox()
        Me.cboLongueurDroite = New System.Windows.Forms.ComboBox()
        Me.lblLongueurDroite = New System.Windows.Forms.Label()
        Me.grpSuperficie = New System.Windows.Forms.GroupBox()
        Me.lblSuperficieInterieure = New System.Windows.Forms.Label()
        Me.cboSuperficieInterieure = New System.Windows.Forms.ComboBox()
        Me.lblSuperficieExterieure = New System.Windows.Forms.Label()
        Me.cboSuperficieExterieure = New System.Windows.Forms.ComboBox()
        Me.grpGeneralisationInterieure = New System.Windows.Forms.GroupBox()
        Me.lblLongueurInterieure = New System.Windows.Forms.Label()
        Me.cboLongueurInterieure = New System.Windows.Forms.ComboBox()
        Me.lblLargeurInterieure = New System.Windows.Forms.Label()
        Me.cboLargeurInterieure = New System.Windows.Forms.ComboBox()
        Me.lblPrecision = New System.Windows.Forms.Label()
        Me.cboPrecision = New System.Windows.Forms.ComboBox()
        Me.lblEchelle = New System.Windows.Forms.Label()
        Me.cboEchelle = New System.Windows.Forms.ComboBox()
        Me.pgeParametres = New System.Windows.Forms.TabPage()
        Me.lblClasseGeneraliser = New System.Windows.Forms.Label()
        Me.cboClasseGeneraliser = New System.Windows.Forms.ComboBox()
        Me.lblClasseSquelette = New System.Windows.Forms.Label()
        Me.cboClasseSquelette = New System.Windows.Forms.ComboBox()
        Me.lblMethodeSquelette = New System.Windows.Forms.Label()
        Me.cboMethodeSquelette = New System.Windows.Forms.ComboBox()
        Me.lblAttributsAdjacence = New System.Windows.Forms.Label()
        Me.cboAttributsAdjacence = New System.Windows.Forms.ComboBox()
        Me.chkCreerFichierErreurs = New System.Windows.Forms.CheckBox()
        Me.pgeMessages = New System.Windows.Forms.TabPage()
        Me.rtbMessages = New System.Windows.Forms.RichTextBox()
        Me.tabEdition.SuspendLayout()
        Me.pgeDimensionMinimale.SuspendLayout()
        Me.grpGeneralisationExterieure.SuspendLayout()
        Me.grpDistance.SuspendLayout()
        Me.grpLongueur.SuspendLayout()
        Me.grpSuperficie.SuspendLayout()
        Me.grpGeneralisationInterieure.SuspendLayout()
        Me.pgeParametres.SuspendLayout()
        Me.pgeMessages.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnInitialiser
        '
        Me.btnInitialiser.Location = New System.Drawing.Point(3, 267)
        Me.btnInitialiser.Name = "btnInitialiser"
        Me.btnInitialiser.Size = New System.Drawing.Size(117, 30)
        Me.btnInitialiser.TabIndex = 0
        Me.btnInitialiser.Text = "Initialiser"
        Me.btnInitialiser.UseVisualStyleBackColor = True
        '
        'tabEdition
        '
        Me.tabEdition.Controls.Add(Me.pgeDimensionMinimale)
        Me.tabEdition.Controls.Add(Me.pgeParametres)
        Me.tabEdition.Controls.Add(Me.pgeMessages)
        Me.tabEdition.ItemSize = New System.Drawing.Size(114, 18)
        Me.tabEdition.Location = New System.Drawing.Point(3, 3)
        Me.tabEdition.Name = "tabEdition"
        Me.tabEdition.SelectedIndex = 0
        Me.tabEdition.Size = New System.Drawing.Size(296, 260)
        Me.tabEdition.TabIndex = 1
        '
        'pgeDimensionMinimale
        '
        Me.pgeDimensionMinimale.Controls.Add(Me.grpGeneralisationExterieure)
        Me.pgeDimensionMinimale.Controls.Add(Me.grpDistance)
        Me.pgeDimensionMinimale.Controls.Add(Me.grpLongueur)
        Me.pgeDimensionMinimale.Controls.Add(Me.grpSuperficie)
        Me.pgeDimensionMinimale.Controls.Add(Me.grpGeneralisationInterieure)
        Me.pgeDimensionMinimale.Controls.Add(Me.lblPrecision)
        Me.pgeDimensionMinimale.Controls.Add(Me.cboPrecision)
        Me.pgeDimensionMinimale.Controls.Add(Me.lblEchelle)
        Me.pgeDimensionMinimale.Controls.Add(Me.cboEchelle)
        Me.pgeDimensionMinimale.Location = New System.Drawing.Point(4, 22)
        Me.pgeDimensionMinimale.Name = "pgeDimensionMinimale"
        Me.pgeDimensionMinimale.Padding = New System.Windows.Forms.Padding(3)
        Me.pgeDimensionMinimale.Size = New System.Drawing.Size(288, 234)
        Me.pgeDimensionMinimale.TabIndex = 1
        Me.pgeDimensionMinimale.Text = "Dimensions minimales"
        Me.pgeDimensionMinimale.UseVisualStyleBackColor = True
        '
        'grpGeneralisationExterieure
        '
        Me.grpGeneralisationExterieure.Controls.Add(Me.lblLongueurExterieure)
        Me.grpGeneralisationExterieure.Controls.Add(Me.cboLongueurExterieure)
        Me.grpGeneralisationExterieure.Controls.Add(Me.lblLargeurExterieure)
        Me.grpGeneralisationExterieure.Controls.Add(Me.cboLargeurExterieure)
        Me.grpGeneralisationExterieure.Location = New System.Drawing.Point(7, 190)
        Me.grpGeneralisationExterieure.Name = "grpGeneralisationExterieure"
        Me.grpGeneralisationExterieure.Size = New System.Drawing.Size(275, 40)
        Me.grpGeneralisationExterieure.TabIndex = 24
        Me.grpGeneralisationExterieure.TabStop = False
        Me.grpGeneralisationExterieure.Text = "Généralisation extérieure"
        '
        'lblLongueurExterieure
        '
        Me.lblLongueurExterieure.AutoSize = True
        Me.lblLongueurExterieure.Location = New System.Drawing.Point(141, 17)
        Me.lblLongueurExterieure.Name = "lblLongueurExterieure"
        Me.lblLongueurExterieure.Size = New System.Drawing.Size(55, 13)
        Me.lblLongueurExterieure.TabIndex = 17
        Me.lblLongueurExterieure.Text = "Longueur:"
        '
        'cboLongueurExterieure
        '
        Me.cboLongueurExterieure.FormattingEnabled = True
        Me.cboLongueurExterieure.Items.AddRange(New Object() {"5.0", "25.0"})
        Me.cboLongueurExterieure.Location = New System.Drawing.Point(198, 14)
        Me.cboLongueurExterieure.Name = "cboLongueurExterieure"
        Me.cboLongueurExterieure.Size = New System.Drawing.Size(70, 21)
        Me.cboLongueurExterieure.TabIndex = 16
        Me.cboLongueurExterieure.Text = "25.0"
        '
        'lblLargeurExterieure
        '
        Me.lblLargeurExterieure.AutoSize = True
        Me.lblLargeurExterieure.Location = New System.Drawing.Point(3, 17)
        Me.lblLargeurExterieure.Name = "lblLargeurExterieure"
        Me.lblLargeurExterieure.Size = New System.Drawing.Size(46, 13)
        Me.lblLargeurExterieure.TabIndex = 15
        Me.lblLargeurExterieure.Text = "Largeur:"
        '
        'cboLargeurExterieure
        '
        Me.cboLargeurExterieure.FormattingEnabled = True
        Me.cboLargeurExterieure.Items.AddRange(New Object() {"25.0", "125.0"})
        Me.cboLargeurExterieure.Location = New System.Drawing.Point(63, 14)
        Me.cboLargeurExterieure.Name = "cboLargeurExterieure"
        Me.cboLargeurExterieure.Size = New System.Drawing.Size(70, 21)
        Me.cboLargeurExterieure.TabIndex = 14
        Me.cboLargeurExterieure.Text = "125.0"
        '
        'grpDistance
        '
        Me.grpDistance.Controls.Add(Me.lblDistanceDensifier)
        Me.grpDistance.Controls.Add(Me.cboDistanceDensifier)
        Me.grpDistance.Controls.Add(Me.lblDistanceLaterale)
        Me.grpDistance.Controls.Add(Me.cboDistanceLaterale)
        Me.grpDistance.Location = New System.Drawing.Point(6, 27)
        Me.grpDistance.Name = "grpDistance"
        Me.grpDistance.Size = New System.Drawing.Size(275, 40)
        Me.grpDistance.TabIndex = 23
        Me.grpDistance.TabStop = False
        Me.grpDistance.Text = "Distance"
        '
        'lblDistanceDensifier
        '
        Me.lblDistanceDensifier.AutoSize = True
        Me.lblDistanceDensifier.Location = New System.Drawing.Point(141, 17)
        Me.lblDistanceDensifier.Name = "lblDistanceDensifier"
        Me.lblDistanceDensifier.Size = New System.Drawing.Size(51, 13)
        Me.lblDistanceDensifier.TabIndex = 23
        Me.lblDistanceDensifier.Text = "Densifier:"
        '
        'cboDistanceDensifier
        '
        Me.cboDistanceDensifier.FormattingEnabled = True
        Me.cboDistanceDensifier.Items.AddRange(New Object() {"15.0", "75.0"})
        Me.cboDistanceDensifier.Location = New System.Drawing.Point(198, 14)
        Me.cboDistanceDensifier.Name = "cboDistanceDensifier"
        Me.cboDistanceDensifier.Size = New System.Drawing.Size(70, 21)
        Me.cboDistanceDensifier.TabIndex = 22
        Me.cboDistanceDensifier.Text = "75.0"
        '
        'lblDistanceLaterale
        '
        Me.lblDistanceLaterale.AutoSize = True
        Me.lblDistanceLaterale.Location = New System.Drawing.Point(4, 17)
        Me.lblDistanceLaterale.Name = "lblDistanceLaterale"
        Me.lblDistanceLaterale.Size = New System.Drawing.Size(48, 13)
        Me.lblDistanceLaterale.TabIndex = 21
        Me.lblDistanceLaterale.Text = "Latérale:"
        '
        'cboDistanceLaterale
        '
        Me.cboDistanceLaterale.FormattingEnabled = True
        Me.cboDistanceLaterale.Items.AddRange(New Object() {"1.5", "7.5"})
        Me.cboDistanceLaterale.Location = New System.Drawing.Point(63, 14)
        Me.cboDistanceLaterale.Name = "cboDistanceLaterale"
        Me.cboDistanceLaterale.Size = New System.Drawing.Size(70, 21)
        Me.cboDistanceLaterale.TabIndex = 20
        Me.cboDistanceLaterale.Text = "7.5"
        '
        'grpLongueur
        '
        Me.grpLongueur.Controls.Add(Me.lblLongueurLigne)
        Me.grpLongueur.Controls.Add(Me.cboLongueurLigne)
        Me.grpLongueur.Controls.Add(Me.cboLongueurDroite)
        Me.grpLongueur.Controls.Add(Me.lblLongueurDroite)
        Me.grpLongueur.Location = New System.Drawing.Point(6, 68)
        Me.grpLongueur.Name = "grpLongueur"
        Me.grpLongueur.Size = New System.Drawing.Size(275, 40)
        Me.grpLongueur.TabIndex = 22
        Me.grpLongueur.TabStop = False
        Me.grpLongueur.Text = "Longueur"
        '
        'lblLongueurLigne
        '
        Me.lblLongueurLigne.AutoSize = True
        Me.lblLongueurLigne.Location = New System.Drawing.Point(141, 17)
        Me.lblLongueurLigne.Name = "lblLongueurLigne"
        Me.lblLongueurLigne.Size = New System.Drawing.Size(36, 13)
        Me.lblLongueurLigne.TabIndex = 9
        Me.lblLongueurLigne.Text = "Ligne:"
        '
        'cboLongueurLigne
        '
        Me.cboLongueurLigne.FormattingEnabled = True
        Me.cboLongueurLigne.Items.AddRange(New Object() {"250.0", "1500.0"})
        Me.cboLongueurLigne.Location = New System.Drawing.Point(198, 14)
        Me.cboLongueurLigne.Name = "cboLongueurLigne"
        Me.cboLongueurLigne.Size = New System.Drawing.Size(70, 21)
        Me.cboLongueurLigne.TabIndex = 8
        Me.cboLongueurLigne.Text = "1500.0"
        '
        'cboLongueurDroite
        '
        Me.cboLongueurDroite.FormattingEnabled = True
        Me.cboLongueurDroite.Items.AddRange(New Object() {"3.0", "15.0"})
        Me.cboLongueurDroite.Location = New System.Drawing.Point(63, 14)
        Me.cboLongueurDroite.Name = "cboLongueurDroite"
        Me.cboLongueurDroite.Size = New System.Drawing.Size(70, 21)
        Me.cboLongueurDroite.TabIndex = 6
        Me.cboLongueurDroite.Text = "15.0"
        '
        'lblLongueurDroite
        '
        Me.lblLongueurDroite.AutoSize = True
        Me.lblLongueurDroite.Location = New System.Drawing.Point(3, 15)
        Me.lblLongueurDroite.Name = "lblLongueurDroite"
        Me.lblLongueurDroite.Size = New System.Drawing.Size(38, 13)
        Me.lblLongueurDroite.TabIndex = 5
        Me.lblLongueurDroite.Text = "Droite:"
        '
        'grpSuperficie
        '
        Me.grpSuperficie.Controls.Add(Me.lblSuperficieInterieure)
        Me.grpSuperficie.Controls.Add(Me.cboSuperficieInterieure)
        Me.grpSuperficie.Controls.Add(Me.lblSuperficieExterieure)
        Me.grpSuperficie.Controls.Add(Me.cboSuperficieExterieure)
        Me.grpSuperficie.Location = New System.Drawing.Point(6, 109)
        Me.grpSuperficie.Name = "grpSuperficie"
        Me.grpSuperficie.Size = New System.Drawing.Size(275, 40)
        Me.grpSuperficie.TabIndex = 21
        Me.grpSuperficie.TabStop = False
        Me.grpSuperficie.Text = "Superficie"
        '
        'lblSuperficieInterieure
        '
        Me.lblSuperficieInterieure.AutoSize = True
        Me.lblSuperficieInterieure.Location = New System.Drawing.Point(141, 17)
        Me.lblSuperficieInterieure.Name = "lblSuperficieInterieure"
        Me.lblSuperficieInterieure.Size = New System.Drawing.Size(54, 13)
        Me.lblSuperficieInterieure.TabIndex = 13
        Me.lblSuperficieInterieure.Text = "Intérieure:"
        '
        'cboSuperficieInterieure
        '
        Me.cboSuperficieInterieure.FormattingEnabled = True
        Me.cboSuperficieInterieure.Items.AddRange(New Object() {"1800.0", "35000.0"})
        Me.cboSuperficieInterieure.Location = New System.Drawing.Point(198, 14)
        Me.cboSuperficieInterieure.Name = "cboSuperficieInterieure"
        Me.cboSuperficieInterieure.Size = New System.Drawing.Size(70, 21)
        Me.cboSuperficieInterieure.TabIndex = 12
        Me.cboSuperficieInterieure.Text = "35000.0"
        '
        'lblSuperficieExterieure
        '
        Me.lblSuperficieExterieure.AutoSize = True
        Me.lblSuperficieExterieure.Location = New System.Drawing.Point(4, 17)
        Me.lblSuperficieExterieure.Name = "lblSuperficieExterieure"
        Me.lblSuperficieExterieure.Size = New System.Drawing.Size(57, 13)
        Me.lblSuperficieExterieure.TabIndex = 11
        Me.lblSuperficieExterieure.Text = "Extérieure:"
        '
        'cboSuperficieExterieure
        '
        Me.cboSuperficieExterieure.FormattingEnabled = True
        Me.cboSuperficieExterieure.Items.AddRange(New Object() {"3600.0", "70000.0"})
        Me.cboSuperficieExterieure.Location = New System.Drawing.Point(63, 14)
        Me.cboSuperficieExterieure.Name = "cboSuperficieExterieure"
        Me.cboSuperficieExterieure.Size = New System.Drawing.Size(70, 21)
        Me.cboSuperficieExterieure.TabIndex = 10
        Me.cboSuperficieExterieure.Text = "70000.0"
        '
        'grpGeneralisationInterieure
        '
        Me.grpGeneralisationInterieure.Controls.Add(Me.lblLongueurInterieure)
        Me.grpGeneralisationInterieure.Controls.Add(Me.cboLongueurInterieure)
        Me.grpGeneralisationInterieure.Controls.Add(Me.lblLargeurInterieure)
        Me.grpGeneralisationInterieure.Controls.Add(Me.cboLargeurInterieure)
        Me.grpGeneralisationInterieure.Location = New System.Drawing.Point(6, 150)
        Me.grpGeneralisationInterieure.Name = "grpGeneralisationInterieure"
        Me.grpGeneralisationInterieure.Size = New System.Drawing.Size(275, 40)
        Me.grpGeneralisationInterieure.TabIndex = 20
        Me.grpGeneralisationInterieure.TabStop = False
        Me.grpGeneralisationInterieure.Text = "Généralisation intérieure"
        '
        'lblLongueurInterieure
        '
        Me.lblLongueurInterieure.AutoSize = True
        Me.lblLongueurInterieure.Location = New System.Drawing.Point(141, 17)
        Me.lblLongueurInterieure.Name = "lblLongueurInterieure"
        Me.lblLongueurInterieure.Size = New System.Drawing.Size(55, 13)
        Me.lblLongueurInterieure.TabIndex = 17
        Me.lblLongueurInterieure.Text = "Longueur:"
        '
        'cboLongueurInterieure
        '
        Me.cboLongueurInterieure.FormattingEnabled = True
        Me.cboLongueurInterieure.Items.AddRange(New Object() {"50.0", "250.0"})
        Me.cboLongueurInterieure.Location = New System.Drawing.Point(198, 14)
        Me.cboLongueurInterieure.Name = "cboLongueurInterieure"
        Me.cboLongueurInterieure.Size = New System.Drawing.Size(70, 21)
        Me.cboLongueurInterieure.TabIndex = 16
        Me.cboLongueurInterieure.Text = "250.0"
        '
        'lblLargeurInterieure
        '
        Me.lblLargeurInterieure.AutoSize = True
        Me.lblLargeurInterieure.Location = New System.Drawing.Point(3, 17)
        Me.lblLargeurInterieure.Name = "lblLargeurInterieure"
        Me.lblLargeurInterieure.Size = New System.Drawing.Size(46, 13)
        Me.lblLargeurInterieure.TabIndex = 15
        Me.lblLargeurInterieure.Text = "Largeur:"
        '
        'cboLargeurInterieure
        '
        Me.cboLargeurInterieure.FormattingEnabled = True
        Me.cboLargeurInterieure.Items.AddRange(New Object() {"25.0", "125.0"})
        Me.cboLargeurInterieure.Location = New System.Drawing.Point(63, 14)
        Me.cboLargeurInterieure.Name = "cboLargeurInterieure"
        Me.cboLargeurInterieure.Size = New System.Drawing.Size(70, 21)
        Me.cboLargeurInterieure.TabIndex = 14
        Me.cboLargeurInterieure.Text = "125.0"
        '
        'lblPrecision
        '
        Me.lblPrecision.AutoSize = True
        Me.lblPrecision.Location = New System.Drawing.Point(147, 8)
        Me.lblPrecision.Name = "lblPrecision"
        Me.lblPrecision.Size = New System.Drawing.Size(53, 13)
        Me.lblPrecision.TabIndex = 17
        Me.lblPrecision.Text = "Précision:"
        '
        'cboPrecision
        '
        Me.cboPrecision.FormattingEnabled = True
        Me.cboPrecision.Items.AddRange(New Object() {"0.001"})
        Me.cboPrecision.Location = New System.Drawing.Point(204, 5)
        Me.cboPrecision.Name = "cboPrecision"
        Me.cboPrecision.Size = New System.Drawing.Size(70, 21)
        Me.cboPrecision.TabIndex = 16
        Me.cboPrecision.Text = "0.001"
        '
        'lblEchelle
        '
        Me.lblEchelle.AutoSize = True
        Me.lblEchelle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEchelle.Location = New System.Drawing.Point(9, 8)
        Me.lblEchelle.Name = "lblEchelle"
        Me.lblEchelle.Size = New System.Drawing.Size(53, 13)
        Me.lblEchelle.TabIndex = 1
        Me.lblEchelle.Text = "Échelle:"
        '
        'cboEchelle
        '
        Me.cboEchelle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEchelle.FormattingEnabled = True
        Me.cboEchelle.Items.AddRange(New Object() {"50000", "250000"})
        Me.cboEchelle.Location = New System.Drawing.Point(69, 5)
        Me.cboEchelle.Name = "cboEchelle"
        Me.cboEchelle.Size = New System.Drawing.Size(70, 21)
        Me.cboEchelle.TabIndex = 0
        '
        'pgeParametres
        '
        Me.pgeParametres.Controls.Add(Me.lblClasseGeneraliser)
        Me.pgeParametres.Controls.Add(Me.cboClasseGeneraliser)
        Me.pgeParametres.Controls.Add(Me.lblClasseSquelette)
        Me.pgeParametres.Controls.Add(Me.cboClasseSquelette)
        Me.pgeParametres.Controls.Add(Me.lblMethodeSquelette)
        Me.pgeParametres.Controls.Add(Me.cboMethodeSquelette)
        Me.pgeParametres.Controls.Add(Me.lblAttributsAdjacence)
        Me.pgeParametres.Controls.Add(Me.cboAttributsAdjacence)
        Me.pgeParametres.Controls.Add(Me.chkCreerFichierErreurs)
        Me.pgeParametres.Location = New System.Drawing.Point(4, 22)
        Me.pgeParametres.Name = "pgeParametres"
        Me.pgeParametres.Size = New System.Drawing.Size(288, 234)
        Me.pgeParametres.TabIndex = 3
        Me.pgeParametres.Text = "Paramètres"
        Me.pgeParametres.UseVisualStyleBackColor = True
        '
        'lblClasseGeneraliser
        '
        Me.lblClasseGeneraliser.AutoSize = True
        Me.lblClasseGeneraliser.Location = New System.Drawing.Point(4, 171)
        Me.lblClasseGeneraliser.Name = "lblClasseGeneraliser"
        Me.lblClasseGeneraliser.Size = New System.Drawing.Size(122, 13)
        Me.lblClasseGeneraliser.TabIndex = 33
        Me.lblClasseGeneraliser.Text = "Classe pour généraliser :"
        '
        'cboClasseGeneraliser
        '
        Me.cboClasseGeneraliser.FormattingEnabled = True
        Me.cboClasseGeneraliser.Location = New System.Drawing.Point(6, 188)
        Me.cboClasseGeneraliser.Name = "cboClasseGeneraliser"
        Me.cboClasseGeneraliser.Size = New System.Drawing.Size(264, 21)
        Me.cboClasseGeneraliser.TabIndex = 32
        '
        'lblClasseSquelette
        '
        Me.lblClasseSquelette.AutoSize = True
        Me.lblClasseSquelette.Location = New System.Drawing.Point(3, 126)
        Me.lblClasseSquelette.Name = "lblClasseSquelette"
        Me.lblClasseSquelette.Size = New System.Drawing.Size(105, 13)
        Me.lblClasseSquelette.TabIndex = 31
        Me.lblClasseSquelette.Text = "Classe du squelette :"
        '
        'cboClasseSquelette
        '
        Me.cboClasseSquelette.FormattingEnabled = True
        Me.cboClasseSquelette.Location = New System.Drawing.Point(6, 143)
        Me.cboClasseSquelette.Name = "cboClasseSquelette"
        Me.cboClasseSquelette.Size = New System.Drawing.Size(264, 21)
        Me.cboClasseSquelette.TabIndex = 30
        '
        'lblMethodeSquelette
        '
        Me.lblMethodeSquelette.AutoSize = True
        Me.lblMethodeSquelette.Location = New System.Drawing.Point(3, 82)
        Me.lblMethodeSquelette.Name = "lblMethodeSquelette"
        Me.lblMethodeSquelette.Size = New System.Drawing.Size(163, 13)
        Me.lblMethodeSquelette.TabIndex = 29
        Me.lblMethodeSquelette.Text = "Méthode pour créer le squelette :"
        '
        'cboMethodeSquelette
        '
        Me.cboMethodeSquelette.FormattingEnabled = True
        Me.cboMethodeSquelette.Items.AddRange(New Object() {"0:Delaunay", "1:Voronoi"})
        Me.cboMethodeSquelette.Location = New System.Drawing.Point(6, 99)
        Me.cboMethodeSquelette.Name = "cboMethodeSquelette"
        Me.cboMethodeSquelette.Size = New System.Drawing.Size(264, 21)
        Me.cboMethodeSquelette.TabIndex = 28
        Me.cboMethodeSquelette.Text = "0:Delaunay"
        '
        'lblAttributsAdjacence
        '
        Me.lblAttributsAdjacence.AutoSize = True
        Me.lblAttributsAdjacence.Location = New System.Drawing.Point(3, 36)
        Me.lblAttributsAdjacence.Name = "lblAttributsAdjacence"
        Me.lblAttributsAdjacence.Size = New System.Drawing.Size(112, 13)
        Me.lblAttributsAdjacence.TabIndex = 27
        Me.lblAttributsAdjacence.Text = "Attributs d'adjacence :"
        '
        'cboAttributsAdjacence
        '
        Me.cboAttributsAdjacence.FormattingEnabled = True
        Me.cboAttributsAdjacence.Items.AddRange(New Object() {"", "CODE_SPEC", "CODE_SPEC/DATASET_NAME"})
        Me.cboAttributsAdjacence.Location = New System.Drawing.Point(6, 53)
        Me.cboAttributsAdjacence.Name = "cboAttributsAdjacence"
        Me.cboAttributsAdjacence.Size = New System.Drawing.Size(264, 21)
        Me.cboAttributsAdjacence.TabIndex = 26
        '
        'chkCreerFichierErreurs
        '
        Me.chkCreerFichierErreurs.AutoSize = True
        Me.chkCreerFichierErreurs.Location = New System.Drawing.Point(6, 12)
        Me.chkCreerFichierErreurs.Name = "chkCreerFichierErreurs"
        Me.chkCreerFichierErreurs.Size = New System.Drawing.Size(146, 17)
        Me.chkCreerFichierErreurs.TabIndex = 25
        Me.chkCreerFichierErreurs.Text = "Créer les fichiers d'erreurs"
        Me.chkCreerFichierErreurs.UseVisualStyleBackColor = True
        '
        'pgeMessages
        '
        Me.pgeMessages.Controls.Add(Me.rtbMessages)
        Me.pgeMessages.Location = New System.Drawing.Point(4, 22)
        Me.pgeMessages.Name = "pgeMessages"
        Me.pgeMessages.Padding = New System.Windows.Forms.Padding(3)
        Me.pgeMessages.Size = New System.Drawing.Size(288, 234)
        Me.pgeMessages.TabIndex = 2
        Me.pgeMessages.Text = "Messages"
        Me.pgeMessages.UseVisualStyleBackColor = True
        '
        'rtbMessages
        '
        Me.rtbMessages.Location = New System.Drawing.Point(3, 3)
        Me.rtbMessages.Name = "rtbMessages"
        Me.rtbMessages.Size = New System.Drawing.Size(284, 228)
        Me.rtbMessages.TabIndex = 0
        Me.rtbMessages.Text = ""
        '
        'dckMenuGeneralisation
        '
        Me.Controls.Add(Me.tabEdition)
        Me.Controls.Add(Me.btnInitialiser)
        Me.Name = "dckMenuGeneralisation"
        Me.Size = New System.Drawing.Size(300, 300)
        Me.tabEdition.ResumeLayout(False)
        Me.pgeDimensionMinimale.ResumeLayout(False)
        Me.pgeDimensionMinimale.PerformLayout()
        Me.grpGeneralisationExterieure.ResumeLayout(False)
        Me.grpGeneralisationExterieure.PerformLayout()
        Me.grpDistance.ResumeLayout(False)
        Me.grpDistance.PerformLayout()
        Me.grpLongueur.ResumeLayout(False)
        Me.grpLongueur.PerformLayout()
        Me.grpSuperficie.ResumeLayout(False)
        Me.grpSuperficie.PerformLayout()
        Me.grpGeneralisationInterieure.ResumeLayout(False)
        Me.grpGeneralisationInterieure.PerformLayout()
        Me.pgeParametres.ResumeLayout(False)
        Me.pgeParametres.PerformLayout()
        Me.pgeMessages.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnInitialiser As System.Windows.Forms.Button
    Friend WithEvents tabEdition As System.Windows.Forms.TabControl
    Friend WithEvents pgeDimensionMinimale As System.Windows.Forms.TabPage
    Friend WithEvents grpSuperficie As System.Windows.Forms.GroupBox
    Friend WithEvents grpGeneralisationInterieure As System.Windows.Forms.GroupBox
    Friend WithEvents lblLongueurInterieure As System.Windows.Forms.Label
    Friend WithEvents cboLongueurInterieure As System.Windows.Forms.ComboBox
    Friend WithEvents lblLargeurInterieure As System.Windows.Forms.Label
    Friend WithEvents cboLargeurInterieure As System.Windows.Forms.ComboBox
    Friend WithEvents lblPrecision As System.Windows.Forms.Label
    Friend WithEvents cboPrecision As System.Windows.Forms.ComboBox
    Friend WithEvents lblEchelle As System.Windows.Forms.Label
    Friend WithEvents cboEchelle As System.Windows.Forms.ComboBox
    Friend WithEvents lblSuperficieInterieure As System.Windows.Forms.Label
    Friend WithEvents cboSuperficieInterieure As System.Windows.Forms.ComboBox
    Friend WithEvents lblSuperficieExterieure As System.Windows.Forms.Label
    Friend WithEvents cboSuperficieExterieure As System.Windows.Forms.ComboBox
    Friend WithEvents grpDistance As System.Windows.Forms.GroupBox
    Friend WithEvents lblDistanceDensifier As System.Windows.Forms.Label
    Friend WithEvents cboDistanceDensifier As System.Windows.Forms.ComboBox
    Friend WithEvents lblDistanceLaterale As System.Windows.Forms.Label
    Friend WithEvents cboDistanceLaterale As System.Windows.Forms.ComboBox
    Friend WithEvents grpLongueur As System.Windows.Forms.GroupBox
    Friend WithEvents lblLongueurLigne As System.Windows.Forms.Label
    Friend WithEvents cboLongueurLigne As System.Windows.Forms.ComboBox
    Friend WithEvents cboLongueurDroite As System.Windows.Forms.ComboBox
    Friend WithEvents lblLongueurDroite As System.Windows.Forms.Label
    Friend WithEvents pgeParametres As System.Windows.Forms.TabPage
    Friend WithEvents lblAttributsAdjacence As System.Windows.Forms.Label
    Friend WithEvents cboAttributsAdjacence As System.Windows.Forms.ComboBox
    Friend WithEvents chkCreerFichierErreurs As System.Windows.Forms.CheckBox
    Friend WithEvents pgeMessages As System.Windows.Forms.TabPage
    Friend WithEvents rtbMessages As System.Windows.Forms.RichTextBox
    Friend WithEvents lblMethodeSquelette As System.Windows.Forms.Label
    Friend WithEvents cboMethodeSquelette As System.Windows.Forms.ComboBox
    Friend WithEvents lblClasseSquelette As System.Windows.Forms.Label
    Friend WithEvents cboClasseSquelette As System.Windows.Forms.ComboBox
    Friend WithEvents lblClasseGeneraliser As System.Windows.Forms.Label
    Friend WithEvents cboClasseGeneraliser As System.Windows.Forms.ComboBox
    Friend WithEvents grpGeneralisationExterieure As System.Windows.Forms.GroupBox
    Friend WithEvents lblLongueurExterieure As System.Windows.Forms.Label
    Friend WithEvents cboLongueurExterieure As System.Windows.Forms.ComboBox
    Friend WithEvents lblLargeurExterieure As System.Windows.Forms.Label
    Friend WithEvents cboLargeurExterieure As System.Windows.Forms.ComboBox

End Class
