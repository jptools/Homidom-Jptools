   M S B A M L   `   `   `  ����D  @HoMIAdmiN, Version=1.3.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication1WpfApplication1  5   �   $    WpfApplication1.uGestionChauff   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      @d2http://schemas.microsoft.com/expression/blend/2008     6<   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     6y   /localclr-namespace:WpfApplication1     6�   $	��419��5      $��5q�6   $��659.792��6    $��576.836��60   .��5      � ���#636060  � ColumnDefinitions  5   
   �� 5      $��0* �6   �� 5	      .+�5   
   �� -   5         ��Name$ Border1��5      
�������$��2q�65   $��2��6I   $��0,0,0,0q�6]   ._�5      ]� $��
0,0,-11,69q�6   .+�5      ]� $	��Top=�6   .+�5      �� -   5         ��Name$ 	ImgDevice��5   k   $
��Left=�6   $��53��63   $
��Left=�6?   $��7,1,0,0q�6Z   $:��4/HoMIAdmiN;component/Images/NouveauComposant_128.png��6~   $
��Fill=�6�   $	��Top=�6�   $��53��6�    � LayoutTransform 5      �� 5       �� Angle$	 -10Q�6.    �� CenterX$ 0Q�6:    �� CenterY$ 0Q�6F   5      �� 5      -   5         ��Name$ Label1��5   U   $��24(�6   $��Italic$�6*   
��� ���$	��Top=�6e   $��Center=�6}   $��Gestion du Chauffage��6�   $	��349��6�    J� RenderTransformOrigin$ 0.508,0.429&�6�   $
��Left=�6�   5      �� 5      	 �� Orientation$	 
Horizontal=�6   $	��644��66   $��	0,0,0,-39q�6B   .+�5      &   System.Windows.Controls.DataGrid  -   5        
  Name$
 	DataGrid1��5      $	��200��62   $	��550��6A   
���   �$��2q�6d      SelectionMode$ Single=�6p   $	��520��6�   $��0,0,0,-19.562q�6�      	RowHeight$ 20Q�6�      CanUserResizeRows .     CanUserReorderColumns .     CanUserResizeColumns .     CanUserSortColumns .     AutoGenerateColumns . $��Center=�6D     CanUserAddRows .  5      �� 5      $	��247��6!   $��5,253,0,-28q�6.   $	��Top=�6C   $	��112��6[   .+�5      �� -   5         ��Name$ RemplirFond��5   #   $��Remplir le fond��68    d� 	IsChecked$
 TrueY�6R   ���	#FFFFFEFE � ToolTip$  Activer la couleur du fond��6z   $ 0.491,11.867&�6�   �� 5      $��Page Active��6#   $
��Left=�68   $��64��6S   $	��112��6_   
�������
���   �$��0,25,0,0q�6�   .��5      �� -   5         ��Name$ ModeChauffage��5   &   $
��Left=�6=   $��7,10,0,8q�6X    �� TextWrapping$
 Wrap=�6j   $��TextBox��6~   $��86��6�   
����������	#FF646363$ Affiche le mode Actif��6�    �� 
IsReadOnly .  �� MaxLines$ 1��6�    �� IsUndoEnabled .  $
��Bold"�6   �� TextAlignment$ Center=�6)  5      �� 5       $��Valeurs Admises��6#   $
��Left=�6<   $	��133��6W   $	��112��6d   
�������
���   �$��0,10,0,0q�6�   .��5!      �� -   5        $ ValeursAdmis��5!   &   $
��Left=�6<   $��7,10,0,4q�6W   $ WrapWithOverflow=�6i   $��TextBox��6�   $��86��6�   
����������	#FF646363$ Valeurs Admises��6�    . $ 5��6�    . $ Center=�6   �� AcceptsReturn .  �� 	UndoLimit$ 97��6G   �� 
AcceptsTab . 6V  5"      5#      5$      5%      5&      �� 5'   
   $��2��6   $��546,24,0,459q�6)   .+�5(      �� -   5         ��Name$ BtnOk��5(   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P      GlassButton#���   6q   $	��Top=�6�   $��89��6�   $��Ok��6�   $&  Fermer et sauvegarder le fichier��6�   �� 5)      -	   5        $ 	BtnFermer��5)   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P   #���   6u   $	��Top=�6�   $��89��6�   $��Fermer��6�   $ 	0.573,2.2&�6�   $ Fermer sans sauvegarde��6�   5*      �� 5+   
   $��2��6   $��546,132,0,273q�6)   .+�5,      �� -
   5        $ BtnCharger1��5,   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P   #���   6w   $	��Top=�6�   $��89��6�   $��Mode Charger��6�   $  Afficher le mode : Charger��6�   �� 5-      -   5        $ 
BtnNormal1��5-   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P   #���   6v   $	��Top=�6�   $��89��6�   $��Mode Normal��6�   $ Afficher le mode : Normal��6�   �� 5.      -   5        $ 
BtnConger1��5.   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P   #���   6v   $	��Top=�6�   $��89��6�   $��Mode Conger��6�   $ Afficher le mode : Conger��6�   �� 5/      -   5        $ BtnAbsence1��5/   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P   #���   6w   $	��Top=�6�   $��89��6�   $��Mode Absence��6�   $  Afficher le mode : Absence��6�   �� 50      -   5        $ 
BtnReduit1��50   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P   #���   6v   $	��Top=�6�   $��89��6�   $��Mode Reduit��6�   $  Afficher le mode : Réduit��6�   �� 51      -   5        $ BtnCalendrier1��51   b   
�������$��25��6(   $��Right=�64   $��0,0,25,0q�6P   #���   6z   $	��Top=�6�   $��89��6�   $��
Calendrier��6�   $ 
0.539,3.64&�6�   $ Afficher le Calendrier��6�   52      53      54      