   M S B A M L   `   `   `  ���� B  >HoMIWpF, Version=1.3.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:HoMIWpFHoMIWpF  C&clr-namespace:HoMIWpF.UAniScrollViewerHoMIWpF.UAniScrollViewer  3clr-namespace:HoMIWpF.DesignerHoMIWpF.Designer  @ <WpfKb, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null>+clr-namespace:WpfKb.Controls;assembly=WpfKbWpfKb.Controls     HoMIWpF.Window1   myclr-namespace:HoMIWpF  R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     8local&clr-namespace:HoMIWpF.UAniScrollViewer     ,sclr-namespace:HoMIWpF.Designer     ;my1+clr-namespace:WpfKb.Controls;assembly=WpfKb       � Title$  Window1��$	��878��$
��1011�� � WindowState$ 	Maximized=� � WindowStyle$
 None=� � 	Resources    sb_Rect �� 
TargetName �� TargetProperty T� Source g� Duration g� AccelerationRatio	 g� RepeatBehavior
 g� AutoReverse T� PathGeometry  sb_Rect2 g� DecelerationRatio  sb_Rect3  	sb_Rect31  sb_Rect4  	sb_Rect41 �� MergedDictionaries �� Source��  �� $! Resources/DesignerItem.xaml?��� $ Resources/SizeChrome.xaml?��� $' !Resources/ResizeRotateChrome.xaml?�%*  &	        &	 }     &	 �     &	 w    &	 �    &	 q    �� .��T� $ ttAnimationTarget_Rect��$ Y�$ Y=�$ 0:0:0.4D�$	 0.5Q�$	 1x��
 .   B� $��M0,0 L0,200C��� .��T� $ ttAnimationTarget_Rect��$ Y�$ Y=�$ 0:0:0.4D�$	 1x��$	 0.6Q�
 .   B� $��M0,200 L0,0C��� .��T� $ ttAnimationTarget_Rect3��$ X�$ X=�$ 0:0:1D�$	 1x��$	 0.6Q�
 .   B� $��M0,0 L2000,0C��� .��T� $ ttAnimationTarget_Rect4��$ X�$ X=�$ 0:0:1D�$	 1x��$	 0.6Q�
 .   B� $��M0,0 L2000,0C��� .��T� $ ttAnimationTarget_Rect3��$ X�$ X=�$ 0:0:0.9D�$	 1x��$	 0.1Q�
 .   B� $��
M0,0 L-1,0C��� .��T� $ ttAnimationTarget_Rect4��$ X�$ X=�$ 0:0:0.9D�$	 1x��$	 0.1Q�
 .   B� $��
M0,0 L-1,0C�.��� .+�]� -   $��
DockPanel1��
���   �.+��� $��0��$
��Left=��� $��30��$	��Top=� �� Orientation$ 
Horizontal=�$��Right=��� $	��130��$��Bottom=��� -   $��ImgBackground��$
��Fill=�$��0,-30,0,-130q�z�V� .~�T� -   ttAnimationTarget_Rect4�� T� X$ 0Q� T� Y$ 0Q�]� .+��� -   $��StkTop��$ 
Horizontal=�$	��Top=�$��0,0,1,0q����	#3A000000.+��� ���	#00000000.+��� -   $��Menu1��$��40��$��40��$��3q�$��Images/Rond_64x64.png�� � ContextMenu �� -   $��	CtxMnuBtn������ .@�m� -   $��ChkMove��$��Déplacer Widget�� m� IsCheckable .  m� 	IsChecked .  $y�	Collapsed=�$��12(�
�����������  �� EndPoint$ 0.5,1&� �� 
StartPoint$ 0.5,0&�.���  � Color$ Black�� � Offset$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� m� Icon �� $��40��$��40��$(��"/HoMIWpF;component/Images/move.png��m� -   $��ChkEdit��$��Editer Widget�� .  .  $y�	Collapsed=�$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$/��)/HoMIWpF;component/Images/Modifier_32.png��m� -	   $��MnuTemplate��$��Template��$y�	Collapsed=�$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$/��)/HoMIWpF;component/Images/Modifier_32.png��.@�m� -
   $��SaveTemplate��$��Sauvegarder le template��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$+��%/HoMIWpF;component/Images/save_32.png��m� -   $
��Chk3��$��Ajouter��$y�	Collapsed=�$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��.@�m� -   $��NewWidgetEmpty��$��Nouveau Widget Vide��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetWeb��$��Nouveau Widget Web��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetRss��$��Nouveau Widget Rss��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetMeteo��$��Nouveau Widget Meteo��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetKeyPad��$��Nouveau Widget KeyPad��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetLabel��$��Nouveau Widget Label��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetGauge��$��Nouveau Widget Gauge��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetChart��$��Nouveau Widget Chart��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetThermostat��$��Nouveau Widget Thermostat��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetCamera��$��Nouveau Widget Camera��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetImage��$��Nouveau Widget Image��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetVolet��$��Nouveau Widget Volet��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetMoteur��$��Nouveau Widget Moteur��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��NewWidgetPrise��$��Nouveau Widget Prise��$��14(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$.��(/HoMIWpF;component/Images/ajouter_32.png��m� -   $��ViewLog��$��Consulter les logs��$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$*��$/HoMIWpF;component/Images/Log_32.png��.@�m� -   $��
ViewLogSrv��$��Serveur��$��12(�
 � Tag$ 0��
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$*��$/HoMIWpF;component/Images/Server.png��m� -   $��ViewCalendar��$��
Calendrier��$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$,��&/HoMIWpF;component/Images/Calendar.png��m� -   $��MnuHisto��$��
Historique��$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$-��'/HoMIWpF;component/Images/Courbe_32.png��m� -   $��MnuMacro��$��Macros��$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$,��&/HoMIWpF;component/Images/Macro_64.png��m� -   $��	MnuConfig��$��
Configurer��$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$1��+/HoMIWpF;component/Images/Configurer_64.png��m� -    $��MnuLastError��$��Dernières erreurs��$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$)��#/HoMIWpF;component/Images/Error.png��m� -!   $��MnuSave��$��Sauvegarder��$��12(�$y�	Collapsed=�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$+��%/HoMIWpF;component/Images/save_32.png��m� -"   $��
MnuRefresh��$��Mettre à jour la config��$��12(�$y�Visible=�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$+��%/HoMIWpF;component/Images/Refresh.png��m� -#   $��
MnuQuitter��$��Quitter��$��12(�
����������� $ 0.5,1&�$ 0.5,0&�.��� $ Black��$ 0Q�� $ 	#FF4B4B4B��$ 0.663Q� �� $��40��$��40��$1��+/HoMIWpF;component/Images/btnQuitter_32.png���� -$   $��LblTime��$��1q�$��25��$��10(�
�������$��RightToLeft=�$��Center=�.��#####�� $	��100��$��0,-1,0,1q��� -%   $��	ImageLeve��$��1q�$
��Fill=�$ Soleil��$��Right=�$��20��$��20��$&�� /HoMIWpF;component/Images/01.png���� -&   $��LblLeve��$��1q�$��Right=�$ Soleil��
�������$��10(�$��Center=�.��?�� -'   $��ImageCouche��$
��Left=�$��1q�$
��Fill=�$ Soleil��$��20��$��20��$&�� /HoMIWpF;component/Images/36.png���� -(   $��	LblCouche��
�������$
��Left=�$��1q�$ Soleil��$��10(�$��Center=�.��?�� -)   $��ImgMess��$
��Left=�$��30,1,1,1q�$
��Fill=�$��20��$��20��$+��%/HoMIWpF;component/Images/Warning.png��$y�	Collapsed=��� -*   $��LblZone��
�������$
��Left=�$��	140,1,1,1q�$
 Zone��$��16(�$��Center=�$��ExtraCondensed&�$��
ExtraBlack"��� -+   $��	StkBottom��$��Bottom=�.+�.   (HoMIWpF.UAniScrollViewer.AniScrollViewer  -,     Name$ ScrollViewer1��$��2q�$��Hidden=�$��Hidden=�$��Bottom=�
����������	#99000000z�V� .~�T� --   ttAnimationTarget_Rect��$ 0Q�$ 0Q��� -.   $��Canvas1��$��1q�}�.  J� SnapsToDevicePixels . z�V� .~�T� -/   ttAnimationTarget_Rect3��$ 0Q�$ 0Q��� -0   $��StkKeyboard��$ 
Horizontal=�$��40��$��right=�$	��Top=�$��Right=�$��2,2,75,2q�.+�   HoMIWpF.keyboard  -1     Name$ 	Keyboard1��$��40����Z� �� -2    ��Name$ BtnQuit��$��right=�$��40��$��2,2,5,2q�$��40��$	��Top=����	#FFFFFFFF  GlassButton#���  $��Right=�.���� $1��+/HoMIWpF;component/Images/btnQuitter_32.png��