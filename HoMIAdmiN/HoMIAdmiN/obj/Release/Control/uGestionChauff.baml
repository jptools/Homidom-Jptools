   M S B A M L   `   `   `  ���� D  @HoMIAdmiN, Version=1.3.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication1WpfApplication1  $    WpfApplication1.uGestionChauff   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     /localclr-namespace:WpfApplication1     $	��419��$��5q�$��659.792��$��576.836��.��� ���#636060  � ColumnDefinitions  �� $��0* ��� .+��� -    ��Name$ Border1��
�������$��2q�$��2��$��0,0,0,0q�._�]� $��
0,0,-11,69q�.+�]� $	��Top=�.+��� -    ��Name$ 	ImgDevice��$
��Left=�$��53��$
��Left=�$��7,1,0,0q�$:��4/HoMIAdmiN;component/Images/NouveauComposant_128.png��$
��Fill=�$	��Top=�$��53�� � LayoutTransform ��  �� Angle$	 -10Q� �� CenterX$ 0Q� �� CenterY$ 0Q��� -    ��Name$ Label1��$��24(�$��Italic$�
��� ���$	��Top=�$��Center=�$��Gestion du Chauffage��$	��349�� J� RenderTransformOrigin$ 0.508,0.429&�$
��Left=��� 	 �� Orientation$	 
Horizontal=�$	��644��$��	0,0,0,-39q�.+�&   System.Windows.Controls.DataGrid  -   
  Name$
 	DataGrid1��$	��200��$	��550��
���   �$��2q�   SelectionMode$ Single=�$	��520��$��0,0,0,-19.562q�   	RowHeight$ 20Q�   CanUserResizeRows .     CanUserReorderColumns .     CanUserResizeColumns .     CanUserSortColumns .     AutoGenerateColumns . $��Center=�   CanUserAddRows .  �� $	��247��$��5,253,0,-28q�$	��Top=�$	��112��.+��� -    ��Name$ RemplirFond��$��Remplir le fond�� d� 	IsChecked$
 TrueY����	#FFFFFEFE � ToolTip$  Activer la couleur du fond��$ 0.491,11.867&��� $��Page Active��$
��Left=�$��64��$	��112��
�������
���   �$��0,25,0,0q�.���� -    ��Name$ ModeChauffage��$
��Left=�$��7,10,0,8q� �� TextWrapping$
 Wrap=�$��TextBox��$��86��
����������	#FF646363$ Affiche le mode Actif�� �� 
IsReadOnly .  �� MaxLines$ 1�� �� IsUndoEnabled .  $
��Bold"� �� TextAlignment$ Center=��� $��Valeurs Admises��$
��Left=�$	��133��$	��112��
�������
���   �$��0,10,0,0q�.���� -   $ ValeursAdmis��$
��Left=�$��7,10,0,4q�$ WrapWithOverflow=�$��TextBox��$��86��
����������	#FF646363$ Valeurs Admises�� . $ 5�� . $ Center=� �� AcceptsReturn .  �� 	UndoLimit$ 97�� �� 
AcceptsTab . �� $��2��$��546,24,0,459q�.+��� -    ��Name$ BtnOk��
�������$��25��$��Right=�$��0,0,25,0q�   GlassButton#���   $	��Top=�$��89��$��Ok��$&  Fermer et sauvegarder le fichier���� -	   $ 	BtnFermer��
�������$��25��$��Right=�$��0,0,25,0q�#���   $	��Top=�$��89��$��Fermer��$ 	0.573,2.2&�$ Fermer sans sauvegarde���� $��2��$��546,132,0,273q�.+��� -
   $ BtnCharger1��
�������$��25��$��Right=�$��0,0,25,0q�#���   $	��Top=�$��89��$��Mode Charger��$  Afficher le mode : Charger���� -   $ 
BtnNormal1��
�������$��25��$��Right=�$��0,0,25,0q�#���   $	��Top=�$��89��$��Mode Normal��$ Afficher le mode : Normal���� -   $ 
BtnConger1��
�������$��25��$��Right=�$��0,0,25,0q�#���   $	��Top=�$��89��$��Mode Conger��$ Afficher le mode : Conger���� -   $ BtnAbsence1��
�������$��25��$��Right=�$��0,0,25,0q�#���   $	��Top=�$��89��$��Mode Absence��$  Afficher le mode : Absence���� -   $ 
BtnReduit1��
�������$��25��$��Right=�$��0,0,25,0q�#���   $	��Top=�$��89��$��Mode Reduit��$  Afficher le mode : Réduit���� -   $ BtnCalendrier1��
�������$��25��$��Right=�$��0,0,25,0q�#���   $	��Top=�$��89��$��
Calendrier��$ 
0.539,3.64&�$ Afficher le Calendrier��