   M S B A M L   `   `   `  ���� D  @HoMIAdmiN, Version=1.3.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication1WpfApplication1  #    WpfApplication1.uModuleSimple   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     ,myclr-namespace:WpfApplication1     $	��700��$	��550��
  � Tag$  MODULESIMPLE��$��20q� � DataContext �� .��� ���#636060.+��� -   $��Border1��
�������$��2q�._��� .+��� -   $��Label1��$	��Top=�$��24(�$��Italic$�
��� ���$��Center=�.��Modules Simples�� -   $��	ImgDevice��$��60��$
��Left=�$��	17,10,0,0q�$
��Fill=�$	��Top=�$��60��$0��*/HoMIAdmiN;component/Images/Module_128.png�� � LayoutTransform ��  �� CenterX$ 0Q� �� CenterY$ 0Q� �� Angle$	 -10Q��� -   $��
BtnAnnuler��$��Annuler��
�������$��25��$��Right=�$��	0,41,12,0q�   GlassButton#���   $	��Top=�$��88.75���� -   $��
BtnAjouter��$��Ajouter��
�������$��25��$��Right=�$��	0,10,12,0q�#���   $	��Top=�$��88.75���� $��12,90,12,12q�.+���  �� Orientation$ 
Horizontal=�.+��� -   $��Label2��$��75��$��Type :��$
��Left=�$	��Top=�
��������� -   $��CbType��$��25��$
��Left=�$	��Top=�$	��450��
�������
���   ����	#FF474646$��2q� J� 	AllowDrop .  � ToolTip$% Type de module simple à créer���� $ 
Horizontal=�$��0,10,0,0q�.+��� -   $��Label3��$��75��$��Nom :��
�������$
��Left=�$	��Top=��� -	   $��TxtNom��$��21��$
��Left=�$	��Top=�$	��450��
�������
���   �
�������$: 4Nom à donner aux macros/triggers qui seront créés���� -
   $��	GroupBox1��$��0,10,0,0q����� 
�������.��Configuration.��� .+��� $��10,10,10,10q�.+��� -   $��StkCompEmet��$ 
Horizontal=�$��0,0,0,5q�$y�	collapsed=�.+��� -   $��Label4��$	��150��$��Composant émetteur :��
�������$
��Left=�$	��Top=��� -   $��
CbEmetteur��$��25��$
��Left=�$	��Top=�$	��450��
�������
���   �
�������$��2q� . $U OComposant qui lancera le module : Interrupteur / télécommande / Détecteur...���� -   $��StkCompRecep��$ 
Horizontal=�$��0,0,0,5q�$y�	collapsed=�.+��� -   $��Label5��$	��150��$��Composant récepteur :��
�������$
��Left=�$	��Top=��� -   $��CbRecepteur��$��25��$
��Left=�$	��Top=�$	��450��
�������
���   �
�������$��2q� . $A ;Composant qui recevra l'ordre : Lampe / Appareil / Volet...���� -   $��StkMulti��$ 
Horizontal=�$��0,0,0,5q�$y�	Collapsed=�.+��� -   $��StkCompMultiEmet��$ vertical=�$��0,0,0,5q�$y�	Collapsed=�$	��320��.+��� -   $��Label6��$	��150��$��Composant émétteur :��
�������$
��Left=�$	��Top=��� -   $��LbEmetteurMulti��$	��200��$
��Left=�$	��280��
�������
���   �
�������$��2q� . $E ?Composants qui enverront l'ordre : Interrupteurs, detecteurs...���� -   $��StkCompMultiRecep��$ vertical=�$��0,0,0,5q�$y�	Collapsed=�$	��320��.+��� -   $��Label7��$	��150��$��Composant récepteur :��
�������$
��Left=�$	��Top=��� -   $��LbRecepteurMulti��$	��200��$
��Left=�$	��280��
�������
���   �
�������$��2q� . $A ;Composant qui recevra l'ordre : Lampe / Appareil / Volet...���� -   $��StkOrdre��$ vertical=�$��0,0,0,5q�$y�visible=�.+��� -   $��StkComposant��$ 
Horizontal=�$��0,0,0,5q�$y�Visible=�.+��� -   $��Label12��$	��150��$��Composant :��
�������$
��Left=�$	��Top=��� -   $��CbComposant��$��25��$
��Left=�$	��Top=�$	��450��
�������
���   ����	#FF474646$��2q� . $A ;Composant qui recevra l'ordre : Lampe / Appareil / Volet...���� -   $��	StkAction��$ 
Horizontal=�$��0,0,0,5q�$y�visible=�.+��� -   $��Label10��$	��150��$��Action :��
�������$
��Left=�$	��Top=��� -   $��CbAction��$��25��$
��Left=�$	��Top=�$	��450��
�������
���   ����	#FF474646$��2q� . $% Type de module simple à créer���� -   $��StkParametre��$ 
Horizontal=�$��0,0,0,5q�$y�visible=�.+��� -    $��Labelparametre��$	��150��$��Paramètre :��
�������$
��Left=�$	��Top=��� -!   $��TxtParametre��$��21��$
��Left=�$	��Top=�$	��450��
�������
���   �
�������$: 4Nom à donner aux macros/triggers qui seront créés���� -"   $��StkCron��$ 
Horizontal=�$��0,0,0,5q�$y�visible=�.+��� -#   $��Label9��$	��150��$��Date-Heure (Cron) :��
�������$
��Left=�$	��Top=��� $ 
Horizontal=�.+��� .+��� $ 
Horizontal=�.+��� -$   $
��LblH��$��Hr��$��Center=�$��1q�$	��Top=�$��35��
�������$��Center=��� -%   $
��LBLM��$��Mn��
�������$��Center=�$��1q�$	��Top=�$��35��$��Center=��� -&   $
��LBLS��$��Ss��
�������$��Center=�$��1q�$	��Top=�$��35��$��Center=��� $ 
Horizontal=�.+��� -'   $��BtnPHr��$
��Left=�$��1q�$��35��$��24��$	��Top=�#���   
�������.���� $4��./HoMIAdmiN;component/Images/fleche_haut_32.png���� -(   $��BtnPMn��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $4��./HoMIAdmiN;component/Images/fleche_haut_32.png���� -)   $��BtnPSc��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $4��./HoMIAdmiN;component/Images/fleche_haut_32.png���� $ 
Horizontal=�.+��� -*   $��TxtHr��$��20(�$
��Left=�$��Center=�$��1q�$��00��$	��Top=�$��35��
���   �
��������� -+   $��TxtMn��$
��Left=�$��1q�$	��Top=�$��35��$��00��$��20(�$��Center=�
���   �
��������� -,   $��TxtSc��$��20(�$
��Left=�$��1q�$��00��$	��Top=�$��35��$��Center=�
���   �
��������� $ 
Horizontal=�.+��� --   $��BtnMHr��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $3��-/HoMIAdmiN;component/Images/fleche_bas_32.png���� -.   $��BtnMMn��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $3��-/HoMIAdmiN;component/Images/fleche_bas_32.png���� -/   $��BtnMSec��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $3��-/HoMIAdmiN;component/Images/fleche_bas_32.png���� $��20���� .+��� $ 
Horizontal=�.+��� -0   $��LblJour��$��Jr��$��Center=�$��1q�$	��Top=�$��35��
�������$��Center=��� -1   $��LblMois��$��Ms��
�������$��Center=�$��1q�$	��Top=�$��35��$��Center=��� $ 
Horizontal=�.+��� -2   $��BtnPJr��$
��Left=�$��1q�$��35��$��24��$	��Top=�#���   
�������.���� $4��./HoMIAdmiN;component/Images/fleche_haut_32.png���� -3   $��BtnPMs��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $4��./HoMIAdmiN;component/Images/fleche_haut_32.png���� $ Horizontal =�.+��� -4   $��TxtJr��$��20(�$
��Left=�$��Center=�$��1q�$�� ��$	��Top=�$��35��
���   �
��������� -5   $��TxtMs��$
��Left=�$��1q�$	��Top=�$��35��$�� ��$��20(�$��Center=�
���   �
��������� $ 
Horizontal=�.+��� -6   $��BtnMJr��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $3��-/HoMIAdmiN;component/Images/fleche_bas_32.png���� -7   $��BtnMMs��$��24��$
��Left=�$��1q�$	��Top=�$��35��#���   
�������.���� $3��-/HoMIAdmiN;component/Images/fleche_bas_32.png���� $��20���� $��2q�.+��� -8   $��	CheckBox1��$��Lundi��$��15��$
��Left=�$	��Top=�
�������
���   �
��������� -9   $��	CheckBox2��$��Mardi��$��15��$
��Left=�$	��Top=�
�������
���   �
��������� -:   $��	CheckBox3��$��Mercredi��$��15��$
��Left=�$	��Top=�
�������
���   �
��������� -;   $��	CheckBox4��$��Jeudi��$��15��$
��Left=�$	��Top=�
�������
���   �
��������� -<   $��	CheckBox5��$��Vendredi��$��15��$
��Left=�$	��Top=�
�������
���   �
��������� -=   $��	CheckBox6��$��Samedi��$��15��$
��Left=�$	��Top=�
�������
���   �
��������� ->   $��	CheckBox7��$��Dimanche��$��15��$
��Left=�$	��Top=�
�������
���   �
��������� -?   $��BtnHelp��$��40��$��Right=�#���   $��0,72.25,20,0q�$	��Top=�$��40��
�������.���� $.��(/HoMIAdmiN;component/Images/Help_128.png��