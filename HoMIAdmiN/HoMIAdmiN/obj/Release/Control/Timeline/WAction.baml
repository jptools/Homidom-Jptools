   M S B A M L   `   `   `  ���� D  @HoMIAdmiN, Version=1.3.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication1WpfApplication1  (    "WpfApplication1.WActionParametrage   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     ,myclr-namespace:WpfApplication1     -   $��WActionDevice��  � Title$  Paramétrer Action�� � ShowInTaskbar .  � WindowStyle$ 
ToolWindow=� � WindowStartupLocation$ CenterOwner=� � Topmost .   � SizeToContent$ WidthAndHeight=� � 
ResizeMode$ NoResize=�.��� ���#636060.+��� $��2q�.+���  �� Orientation$ 
Horizontal=�.+��� -   $��StkProperty��$��5q�.+��� -   $
��Lbl1��$��25��$
��###:��$
��Left=�$	��Top=�
��� ���$��14(��� -   $	��Cb1��$��25��$
��Left=�$	��Top=�$	��200��
���   �
����������	#FF474646$��2q��� -   $
��Lbl2��$��25��$
��###:��
��� ���$
��Left=�$	��Top=�$��14(�$��Tahoma)��� -   $	��Cb2��$��25��$
��Left=�$	��Top=�$	��200��
���   �
����������	#FF474646$��2q��� -   $
��Txt2��$��25��$
��Left=�$	��Top=�$	��199��
���   �
�������
�������$y�	Collapsed=��� -   $��LblValue��$��25��$��
Parametre:��$��Tahoma)�$��14(�
��� ���$
��Left=�$	��Top=��� -	   $��TxtValue��$��25��$
��Left=�$	��Top=�$	��199��
�������
���   �
������� �� AcceptsReturn . �� $��5q�.+��� $��Delai:��$��14(�
��� ����� $ 
Horizontal=�.+��� .+��� -
   $��Label1��$��Heure��$��Center=�$��1q�$	��Top=�$��50��
�������$��Center=�$��10(��� -   $��BtnPHr��$��24��$��Center=�$��1q�$	��Top=�$��35��   GlassButton#���   
�������.���� $4��./HoMIAdmiN;component/Images/Fleche_haut_32.png���� -   $��TxtHr��$��20(�$��Center=�$��Center=�$��1q�$��00��$	��Top=�$��35��
���   �
��������� -   $��BtnMHr��$��24��$��Center=�$��1q�$	��Top=�$��35��#���   
�������.���� $B��</HoMIAdmiN;component/Images/bullet-fleche-vers-le-bas-32.png���� .+��� -   $��Label2��$��Minute��
�������$��Center=�$��1q�$	��Top=�$��50��$��Center=�$��10(��� -   $��BtnPMn��$��Center=�$��1q�$��35��$��24��$	��Top=�#���   
�������.���� $4��./HoMIAdmiN;component/Images/Fleche_haut_32.png���� -   $��TxtMn��$��Center=�$��1q�$	��Top=�$��35��$��00��$��20(�$��Center=�
���   �
��������� -   $��BtnMMn��$��24��$��Center=�$��1q�$	��Top=�$��35��#���   
�������.���� $B��</HoMIAdmiN;component/Images/bullet-fleche-vers-le-bas-32.png���� .+��� -   $��Label3��$��Seconde��
�������$��Center=�$��1q�$	��Top=�$��50��$��Center=�$��10(��� -   $��BtnPSc��$��24��$��Center=�$��1q�$	��Top=�$��35��#���   
�������.���� $4��./HoMIAdmiN;component/Images/Fleche_haut_32.png���� -   $��TxtSc��$��20(�$��Center=�$��1q�$��00��$	��Top=�$��35��$��Center=�
���   �
��������� -   $��BtnMSec��$��24��$��Center=�$��1q�$	��Top=�$��35��#���   
�������.���� $B��</HoMIAdmiN;component/Images/bullet-fleche-vers-le-bas-32.png���� $	��Top=�.+��� -   $��BtnOk��$��Ok��$��28��$
��Left=�$��15q�$	��Top=�$��72��#���   
��������� -   	 ��Name$	 BtnArchiver��$��Archiver��$��28��$
��Left=�$��
15,0,15,15q�$	��Top=�$��72��#���   
�������$y�	Collapsed=��� -   $��TabControl1��$
��Left=�$��5q�$��Center=�$	��611��.@��� -   $��TabItem1��$��	Condition��.��� 
���   �$	��420��.+��� -   $��StackPanel1��$
��Left=�$	��Top=�$ Vertical=�.+��� $ 
Horizontal=�.+��� -   $��BtnCondiTime��$1��+/HoMIAdmiN;component/Images/horloge_128.png��$��24��$��24��$��3q��� -   $��BtnCondiDevice��$3��-/HoMIAdmiN;component/Images/Composant_128.png��$��24��$��24��$��3q��� -   $��BtnCondiVariable��$9��3/HoMIAdmiN;component/Images/Debug-Variable-icon.png��$��24��$��24��$��3q��� -   $��ScrollViewer1��$��5,0,0,0q�$	��Top=�$
��Auto=�$
��Auto=���. $	��590��
 J� 	AllowDrop
 . $	��360��.���� -   $��StkCondition��$ Vertical=�
�������$	��Top=�$
��Left=��� $
��Then��.��   WpfApplication1.uScenario  -      Name$ 
UScenario1��$	��382��$	��593��$	��6,0q��� $
��Else��.��  -!   $ 
UScenario2��$	��417��$	��593��