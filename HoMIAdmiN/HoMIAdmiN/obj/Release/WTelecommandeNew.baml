   M S B A M L   `   `   `  ���� D  @HoMIAdmiN, Version=1.3.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication1WpfApplication1  a ]PresentationFramework.Aero, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35gJclr-namespace:Microsoft.Windows.Themes;assembly=PresentationFramework.AeroMicrosoft.Windows.Themes &     WpfApplication1.WTelecommandeNew   $myclr-namespace:WpfApplication1  \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     oMicrosoft_Windows_ThemesJclr-namespace:Microsoft.Windows.Themes;assembly=PresentationFramework.Aero     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006       � Title$  Télécommandes�� J� 	AllowDrop .  � SizeToContent$ WidthAndHeight=� � WindowStartupLocation$ CenterScreen=�$	��Top=� � 
ResizeMode$ NoResize=� � 	Resources    ImageBrushDataSource  ButtonControlTemplate1 �� 
TargetType �� RadiusY �� RadiusX	 J� Opacity
 �� EndPoint � MappingMode �� 
StartPoint �� RelativeTransform �� CenterY �� CenterX �� CenterY �� CenterX �� Angle �� CenterY �� CenterX � Color � Offset 
� RadiusY 
� RadiusX 
� Center 
� GradientOrigin  ButtonFocusVisual �� Property �� Value J� SnapsToDevicePixels �� StrokeDashArray  ButtonNormalBackground  ButtonNormalBorder �� Color  ButtonStyle1  �� 
TargetType+  %Microsoft.Windows.Themes.ButtonChrome!  Name"   BorderBrush#   
Background$ J� IsMouseOver%   RenderMouseOver&   RenderPressed' �� IsDefaulted(   RenderDefaulted) �� Triggers* P� Property+ P� Value, �� 
TargetName  	OnLoaded1- �� TargetProperty. �� 
TargetName9  3System.Windows.Media.Animation.EasingDoubleKeyFrame/ O� KeyTime0 O� Value1 R� KeyTime  System.Windows.Visibility2   Hidden3 R� Value4   Visible8  2System.Windows.Media.Animation.EasingColorKeyFrame5 �� KeyTime6 �� Value 0 ,HorizontalSliderThumbHoverBackgroundInverted $  HorizontalSliderThumbHoverBorder 2	 .HorizontalSliderThumbPressedBackgroundInverted &
 "HorizontalSliderThumbPressedBorder  SliderThumbDisabledGeometry & "SliderPointedThumbDisabledGeometry ) %SliderPointedThumbOuterBorderGeometry 1 -HorizontalSliderThumbNormalBackgroundInverted * &SliderPointedThumbMiddleBorderGeometry   HorizontalSliderUpThumbStyle7 T� X8 T� Y9 F�Name ( $HorizontalSliderThumbHoverBackground * &HorizontalSliderThumbPressedBackground ) %HorizontalSliderThumbNormalBackground " HorizontalSliderDownThumbStyle % !HorizontalSliderTrackNormalBorder ) %HorizontalSliderTrackNormalBackground  SliderRepeatButtonStyle " SliderThumbOuterBorderGeometry # SliderThumbMiddleBorderGeometry  HorizontalSliderThumbStyle & "VerticalSliderThumbHoverBackground ( $VerticalSliderThumbPressedBackground ' #VerticalSliderThumbNormalBackground   VerticalSliderLeftThumbStyle . *VerticalSliderThumbHoverBackgroundInverted 0  ,VerticalSliderThumbPressedBackgroundInverted /! +VerticalSliderThumbNormalBackgroundInverted !" VerticalSliderRightThumbStyle ## VerticalSliderTrackNormalBorder $ VerticalSliderThumbStyle % SliderStyle1: � RowDefinitions; m�Name< m� Fill= m� 	Placement> ��Name? �� CornerRadius@ ��NameA ]�NameB ]� DecreaseRepeatButtonC �� DecreaseLargeD ]� IncreaseRepeatButtonE �� IncreaseLargeF ]� ThumbG n�NameH �� TriggersI � ColumnDefinitions��%D5  &	        &	      &	 A    &	 �    &	     &	 0    7  7  7  &	 `    &	 1    &	 �    &		 �    &	
 #    &	 6    &	 i    &	 �    &	 ,    &	 �    &	 �    7  7  7  7  7  7  7  7  7 	 7 
 7  7  &	 l    &	 �    &	 K    &	 �    7  7  7  7  7  7  7  7  7  7 
 7  7  &	 1    &	 �    &	 �    &	 (    &	 �    &	 �    7  7  7  7  7  7  7  7  7  7 
 7  &	     &	 �    &	 �    &	 j    7  7  7  7  7  7  7  7  7  7 
 7  &	 �    &	  G     &	! �     &	" '!    7  7 ! 7  7  7  7  7  7  7   7 
 7  &	# �$    &	$ �$    7  7  7  7  7  7  7  7  7  7 
 7  &	% `(    7  7  7  7  7  7  7  7 # 7  7  7  7 $ 7  7 " �� �� " ��.R�� .+��� $ 5��$ 5��
���   �$��0��
���   ��� $ 5��$ 5��
���   �$��0��$	 0.505Q����� $
 0.5,1&�$ RelativeToBoundingBox=�$ 0.5,0&� V� .~��� $	 0.5Q�$	 0.5Q��� $	 0.5Q�$	 0.5Q��� $	 -90Q�$	 0.5Q�$	 0.5Q�T� .��� $ Black��$
 0.05Q�� $	 Red��� $ Black��$
 0.95Q�� $	 Red��$ 1Q�A� $��0,8,0,-8q�
���   �$��0��$	 1Q�$
��Auto��$��Stretch=���
� $ 0.979Q�$ 0.657Q�$ 0.525,0.521&�$ 0.525,0.521&�.��� $ 	#00000000��$ 1Q�� $	 Red���� .���  � �� �� .R��� $��2q� . #��� ��$��1��$	 1 2R��� $	
 0,1&�$	 0,0&�.��� $ #F3F3F3��$ 0Q�� $ #EBEBEB��$	 0.5Q�� $ #DDDDDD��$	 0.5Q�� $ #CDCDCD��$ 1Q��� $ 	#FF707070���� "  ��.���  �@�FocusVisualStyle8   ��  � ��8  ��  � ��8  ��  � ��$ 1q���  � ��# � ����  � ��$ Center=���  � ��$ Center=���  � ��$ 1q���  � �� �� " ��.R�  $! Chrome��#" z��## z��#% z$ #& z��#( z'  . ._��� #��z��#��z����. # z #��z��) P� * �@J�IsKeyboardFocused+ . . ��� $, Chrome�� �@ RenderDefaulted . P� * �@d�	IsChecked$
+ trueY�. ��� $, Chrome�� �@ RenderPressed . P� * � {�+ .  . ���  � �� �#ADADAD�� .��U� $- (UIElement.Opacity)�$. Window��.\�  $/ 0��$0 0Q�  $/ 0:0:0.5��$0 1Q�T� $- (UIElement.Visibility)�$. 
rectangle1��..�i� $1 0��#3 Z2 i� $1 0:0:0.8��#3 Z4 U� $- (UIElement.Opacity)�$. 
rectangle1��.\�  $/ 0��$0 1Q�  $/ 0:0:0.8��$0 0Q�  $/ 0:0:1��$0 1Q�  $/ 0:0:1.4��$
0 0.16Q�T� $- (UIElement.Visibility)�$. 	rectangle��..�i� $1 0��#3 Z2 i� $1 0:0:1.2��#3 Z4 U� $- (UIElement.Opacity)�$. 	rectangle��.\�  $/ 0��$0 1Q�  $/ 0:0:1.2��$0 0Q�  $/ 0:0:1.4��$0 1Q��� $*- $(Shape.Fill).(SolidColorBrush.Color)�$. 
rectangle1��.g�  $5 0��$6 	#33FFFFFF��  $5 0:0:0.8��$6 	#33FF0000��  $5 0:0:1��$6 	#51FFFFFF��  $5 0:0:1.4��$6 White���� $N- H(Panel.Background).(GradientBrush.GradientStops)[1].(GradientStop.Color)�$. Window��.g�  $5 0��$6 Black��  $5 0:0:0.8��$6 	#FF4E4E4E���� $H- B(Shape.Fill).(GradientBrush.GradientStops)[3].(GradientStop.Color)�$. 	rectangle��.g�  $5 0��$6 White��  $5 0:0:1.2��$6 White��  $5 0:0:1.4��$6 White���� $H- B(Shape.Fill).(GradientBrush.GradientStops)[0].(GradientStop.Color)�$. 	rectangle��.g�  $5 0��$6 Black��  $5 0:0:1.2��$6 White��  $5 0:0:1.4��$6 Black���� $H- B(Shape.Fill).(GradientBrush.GradientStops)[2].(GradientStop.Color)�$. 	rectangle��.g�  $5 0��$6 Black��  $5 0:0:1.2��$6 White��  $5 0:0:1.4��$6 Black���� $	
 0,0&�$	 0,1&�.��� $ white��$ 0Q�� $ #B8E2F9��$	 0.5Q�� $ #B0DFF8��$ 1Q��� $ #3C7FB1���� $	
 0,0&�$	 0,1&�.��� $ #B8E4FA��$
 0.25Q�� $ #5EB4E3��$	 0.5Q�� $ #4C8BB4��$ 1Q��� $ #2C628B���-)M -5,-10.5 L 5,-10.5 L 5,10.5 L -5,10.5 Z��FBM 4.5,-8.5 L -4.5,-8.5 L -4.5,4.5 L -0.5,8.5 L 0.5,8.5 L 4.5,4.5 Z��qmM 4.5,-7.5 A 1 1 0 0 0 3.5,-8.5 L -3.5,-8.5 A 1 1 0 0 0 -4.5,-7.5 L -4.5,4.5 L -0.5,8.5 L 0.5,8.5 L 4.5,4.5 Z��� $	
 0,0&�$	 0,1&�.��� $ White��$ 0Q�� $ #F0EFEF��$	 0.4Q�� $ #D6D5D5��$ .8Q��73M 3.5,-7.5 L -3.5,-7.5 L -3.5,4.5 L 0,8 L 3.5,4.5 Z��� "  n�.���  � |� .  ��  �@�OverridesDefaultStyle . ��  � ��$ 18����  � ��$ 10����  � ��
 �������  � �� �� " n�.R���  . z�V� ~�X� �� $	 180Q�T� $7 5Q�$8 9Q�.+�F� $9 
Background��8��  8�� F� $9 InnerBorder��8�� 
�������$��1��F� $9 OuterBorder��8�� ���	#FF929292$��1��) P� * �@J�IsMouseOver+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * � ��
+ ��  �. ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * �@n�
IsDragging+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8 	  � ��P� * � {�+ .  . ��� $, 
Background�� � �� �	#FFF4F4F4�� $, InnerBorder�� Z�  � ���� $, OuterBorder��8 
  � ���� $, OuterBorder�� � �� �	#FFAEB1AF�� $, OuterBorder��8   � ���� $	
 0,1&�$	 0,0&�.��� $ white��$ 0Q�� $ #B8E2F9��$	 0.5Q�� $ #B0DFF8��$ 1Q��� $	
 0,1&�$	 0,0&�.��� $ #B8E4FA��$
 0.25Q�� $ #5EB4E3��$	 0.5Q�� $ #4C8BB4��$ 1Q��� $	
 0,1&�$	 0,0&�.��� $ White��$ 0Q�� $ #F0EFEF��$	 0.4Q�� $ #D6D5D5��$ .8Q��� "  n�.���  � |� .  ��  �@�OverridesDefaultStyle . ��  � ��$ 18����  � ��$ 10����  � ��
 �������  � �� �� " n�.R���  . z�T� $7 5Q�$8 9Q�.+�F� $9 
Background��8��  8�� F� $9 InnerBorder��8�� 
�������$��1��F� $9 OuterBorder��8�� ���	#FF929292$��1��) P� * �@J�IsMouseOver+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * � ��
+ ��  �. ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * �@n�
IsDragging+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8 	  � ��P� * � {�+ .  . ��� $, 
Background�� � �� �	#FFF4F4F4�� $, InnerBorder�� Z�  � ���� $, OuterBorder��8 
  � ���� $, OuterBorder�� � �� �	#FFAEB1AF�� $, OuterBorder��8   � ���� $	
 0,1&�$	 0,0&�.��� $ 	#FFAEB1AF��$	 0.1Q�� $ White��$ .9Q��� $ 	#FFE7EAEA���� "  ��.���  �@�OverridesDefaultStyle . ��  � �� .  ��  � |� .  ��  � �� �� " ��.R��� 
������ �yuM -5,-9.5 A 1 1 0 0 1 -4,-10.5 L 4,-10.5 A 1 1 0 0 1 5,-9.5 L 5,9.5 A 1 1 0 0 1 4,10.5 L -4,10.5 A 1 1 0 0 1 -5,9.5 Z��)%M -4,-9.5 L 4,-9.5 L 4,9.5 L -4,9.5 Z��� "  n�.���  � |� .  ��  �@�OverridesDefaultStyle . ��  � ��$ 22����  � ��$ 11����  � ��
 �������  � �� �� " n�.R���  . z�T� $	7 5.5Q�$8 11Q�.+�F� $9 
Background��8��  8�� F� $9 InnerBorder��8�� 
�������F� $9 OuterBorder��8�� ���	#FF929292) P� * �@J�IsMouseOver+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * � ��
+ ��  �. ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * �@n�
IsDragging+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8 	  � ��P� * � {�+ .  . ��� $, 
Background�� � �� �	#FFF4F4F4�� $, InnerBorder�� Z�  � ���� $, OuterBorder��8 
  � ���� $, OuterBorder�� � �� �	#FFAEB1AF�� $	
 1,0&�$	 0,0&�.��� $ white��$ 0Q�� $ #B8E2F9��$	 0.5Q�� $ #B0DFF8��$ 1Q��� $	
 1,0&�$	 0,0&�.��� $ #B8E4FA��$
 0.25Q�� $ #5EB4E3��$	 0.5Q�� $ #4C8BB4��$ 1Q��� $	
 1,0&�$	 0,0&�.��� $ White��$ 0Q�� $ #F0EFEF��$	 0.4Q�� $ #D6D5D5��$ .8Q��� "  n�.���  � |� .  ��  �@�OverridesDefaultStyle . ��  � ��$ 10����  � ��$ 18����  � ��
 �������  � �� �� " n�.R���  . z�V� ~�X� �� $ 90Q�T� $7 9Q�$8 5Q�.+�F� $9 
Background��8��  8�� F� $9 InnerBorder��8�� 
�������$��1��F� $9 OuterBorder��8�� ���	#FF929292$��1��) P� * �@J�IsMouseOver+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * � ��
+ ��  �. ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * �@n�
IsDragging+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8 	  � ��P� * � {�+ .  . ��� $, 
Background�� � �� �	#FFF4F4F4�� $, InnerBorder�� Z�  � ���� $, OuterBorder��8 
  � ���� $, OuterBorder�� � �� �	#FFAEB1AF�� $	
 0,0&�$	 1,0&�.��� $ white��$ 0Q�� $ #B8E2F9��$	 0.5Q�� $ #B0DFF8��$ 1Q��� $	
 0,0&�$	 1,0&�.��� $ #B8E4FA��$
 0.25Q�� $ #5EB4E3��$	 0.5Q�� $ #4C8BB4��$ 1Q��� $	
 0,0&�$	 1,0&�.��� $ White��$ 0Q�� $ #F0EFEF��$	 0.4Q�� $ #D6D5D5��$ .8Q��� "  n�.���  � |� .  ��  �@�OverridesDefaultStyle . ��  � ��$ 10����  � ��$ 18����  � ��
 �������  � �� �� " n�.R���  . z�V� ~�X� �� $	 -90Q�T� $7 9Q�$8 5Q�.+�F� $9 
Background��8��  8�� F� $9 InnerBorder��8�� 
�������$��1��F� $9 OuterBorder��8�� ���	#FF929292$��1��) P� * �@J�IsMouseOver+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * � ��
+ ��  �. ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * �@n�
IsDragging+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8 	  � ��P� * � {�+ .  . ��� $, 
Background�� � �� �	#FFF4F4F4�� $, InnerBorder�� Z�  � ���� $, OuterBorder��8 
  � ���� $, OuterBorder�� � �� �	#FFAEB1AF�� $	
 1,0&�$	 0,0&�.��� $ 	#FFAEB1AF��$	 0.1Q�� $ White��$ .9Q��� "  n�.���  � |� .  ��  �@�OverridesDefaultStyle . ��  � ��$ 11����  � ��$ 22����  � ��
 �������  � �� �� " n�.R���  . z�V� ~�X� �� $ 90Q�T� $7 11Q�$	8 5.5Q�.+�F� $9 
Background��8��  8�� F� $9 InnerBorder��8�� 
�������$��1��F� $9 OuterBorder��8�� ���	#FF929292$��1��) P� * �@J�IsMouseOver+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * � ��
+ ��  �. ��� $, 
Background��8   � ���� $, OuterBorder��8   � ��P� * �@n�
IsDragging+ . . ��� $, 
Background��8   � ���� $, OuterBorder��8 	  � ��P� * � {�+ .  . ��� $, 
Background�� � �� �	#FFF4F4F4�� $, InnerBorder�� Z�  � ���� $, OuterBorder��8 
  � ���� $, OuterBorder�� � �� �	#FFAEB1AF�� "  ��.���  �@��IsPressAndHoldEnabled .  ��  � ��
 ���� ��  � �� �	#FFC4C4C4��  � �� �� " ��.R��� #��z��#��z��#��z�� . ._�� : �� $
��Auto ��� $
��Auto �#��z���� $
��Auto �.+�m� $; TopTick��#< z��$��4��$	= Top=�$��0��$y�	Collapsed=�m� $; 
BottomTick��#< z��$��4��$= Bottom=�$��2��$y�	Collapsed=��� $> TrackBackground��8��  $��1q�8�� $? 1��$	��4.0��$	��5,0q�$��1��$��center=�._��� $��-6,-1q�.+��� $@ PART_SelectionRange��
���  ��$	��4.0��#��� ��$	��1.0��$y�Hidden=�]� $A 
PART_Track��$��1��B �� #��ZC 8�� D �� #��ZE 8�� F n� $G Thumb��8�� ) P� * �@��TickPlacement$+ TopLeft=�. ��� $, TopTick�� � y�$ Visible=��� $, Thumb��8   � ���� $, TrackBackground�� � ��$ 5,2,5,0q�P� * �@��TickPlacement$+ BottomRight=�. ��� $, 
BottomTick�� � y�$ Visible=��� $, Thumb��8   � ���� $, TrackBackground�� � ��$ 5,0,5,2q�P� * �@��TickPlacement$
+ Both=�. ��� $, TopTick�� � y�$ Visible=��� $, 
BottomTick�� � y�$ Visible=�P� * �@��IsSelectionRangeEnabled+ . . ��� $, PART_SelectionRange�� � y�$ Visible=�P� * �@J�IsKeyboardFocused+ . . ��� $, Thumb�� � ��
 ��  �H P� * �@��Orientation$+ Vertical=�. ���  � �� �� " ��.R��� #��z��#��z��#��z�� . ._�� I �� $
��Auto ��� #��z��$
��Auto ��� $
��Auto �.+�m� $; TopTick��$��0��#< z��$
= Left=�$y�	Collapsed=�$��4��m� $; 
BottomTick��$��2��#< z��$= Right=�$y�	Collapsed=�$��4���� $> TrackBackground��8�� $��1q�8�� $��1��$? 1��$��center=�$	��0,5q�$	��4.0��._��� $��-1,-6q�.+��� $@ PART_SelectionRange��#��� ��#��� ��$	��1.0��$y�Hidden=�$	��4.0��]� $A 
PART_Track��$��1��B �� #��ZC 8��	 D �� #��ZE 8��
 F n� $G Thumb��8�� ) P� * �@��TickPlacement$+ TopLeft=�. ��� $, TopTick�� � y�$ Visible=��� $, Thumb��8   � ���� $, TrackBackground�� � ��$ 2,5,0,5q�P� * �@��TickPlacement$+ BottomRight=�. ��� $, 
BottomTick�� � y�$ Visible=��� $, Thumb��8   � ���� $, TrackBackground�� � ��$ 0,5,2,5q�P� * �@��TickPlacement$
+ Both=�. ��� $, TopTick�� � y�$ Visible=��� $, 
BottomTick�� � y�$ Visible=�P� * �@��IsSelectionRangeEnabled+ . . ��� $, PART_SelectionRange�� � y�$ Visible=�P� * �@J�IsKeyboardFocused+ . . ��� $, Thumb�� � ��
 ��  ����� $
 0.5,1&�$ RelativeToBoundingBox=�$ 0.5,0&�.��� $ 	#FF121212��$
 0.03Q�� $ 	#FF4E4E4E��$ 1Q�J � TriggersJ :� K :� RoutedEvent$K FrameworkElement.Loaded��.��� -   L �Name$L Layout_root��$��2q�M � DataContextM �� N �� Source#N [  .+��� O �� Orientation$O 
Horizontal=�.+��� -   $��StkCde��$
��Left=�$	��875��.+��� $��Universal Driver��
�������$��Italic$�$��SemiBold"�9� -   $��ExpdProperty��$��Propriétés��
�������P 9� 
IsExpandedP . .���� .+��� $O 
Horizontal=�.+��� $��	Template:��$��75��
��������� -   $��
cbTemplate��$	��150�����	#FF474646�� -   $��BtnNewTemplate��$��25��$��2q�$��25��
������� & GlassButton#��� & .���� $0��*/HoMIAdmiN;component/Images/ajouter_32.png���� -   $��BtnDeleteTemplate��$��25��$��2q�$��25��
�������#��� & $y�	Collapsed=�.���� $2��,/HoMIAdmiN;component/Images/Supprimer_32.png���� -   $��BtnImportTemplate��$��25��$��2q�$��25��
�������#��� & $y�	Collapsed=�.���� $/��)/HoMIAdmiN;component/Images/Import_32.png���� -   $��BtnExportTemplate��$��25��$��2q�$��25��
�������#��� & $y�	Collapsed=�.���� $/��)/HoMIAdmiN;component/Images/Export_32.png���� -	   $��
StkTplBase��$O 
Horizontal=�$y�	Collapsed=�.+��� -
   $��BtnSaveTemplate��$��2q�$��25��$��25��$��Ok��$y�Hidden=�
�������#��� & �� -   $��StkNomTemplate��$O 
Horizontal=�.+��� $
��Nom:��
�������$��75���� -   $��
TxtTplName��$	��150��
�������
���   �
�������Q �� 
IsReadOnlyQ . �� -   $��StkNewTemplate��$O 
Horizontal=�.+��� $��Fabriquant:��
�������$��75���� -   $��	TxtTplFab��$	��150��
�������
���   �
�������Q . �� $O 
Horizontal=�.+��� $��Modele:��$��75��
��������� -   $��	TxtTplMod��$	��150��
�������
���   �
�������Q . �� $O 
Horizontal=�.+��� $��Type:��$��75��
�������	� -   $��RdHttp��$
��Http��
�������$��2q�R d� 	IsChecked$
R TrueY�	� -   $
��RdIR��$��IR��
�������$��2q�	� -   $��RdRS232��$��RS232��
�������$��2q��� $O 
Horizontal=�.+��� -   $��ChkMultimedia��$��
Multimedia��
�������$R  TrueY��� $O 
Horizontal=�.+��� $��Trame d'init:��$��75��
�������S � ToolTip$S Trame d'initialisation���� -   $��TxtTrameInit��$	��150��
�������
���   �
�������Q . �� $O 
Horizontal=�.+��� $!��Caractère de fin de trame:��
��������� -   $��TxtCharEndReceive��$��40��
�������
���   �
�������Q . 9� -   $��ExpdCmd��$��	Commandes��
�������.���� -   $��StkCmd��$y�	Collapsed=�.+��� -   $��ListCmd��$
��Left=�$��1q�$	��100��$	��145��
�������
���   �
��������� $O 
Horizontal=�.+��� -   $��Label16��$��50��$
��Left=�
�������$��Nom :���� -   $��
TxtCmdName��$��21��$	��150��$S Nom de la commande��
�������
���   �
��������� $O 
Horizontal=�.+��� -   $��Label17��$��50��
�������$��Right=�$��Repeat :���� -   $��TxtCmdRepeat��$��21��$��Right=�$	��Top=�$��33.244��$'S !Nombre de fois à envoyer le code��
�������
���   �
�������$��0���� -   $��Label27��$��50��
�������$��Right=�$��Format :��$y�	Collapsed=��� -   $��CbFormat��$��21��$��Right=�$	��Top=�$��63.244��$S Format (Uuirt/Pronto)��
�������
���   �
�������T �� Text$T 0��U �� SelectedIndex$U 0��$y�	Collapsed=�.@��� $��Uuirt���� $��Pronto���� .+��� $O 
Horizontal=�.+�	� -   $��RdBIsScriptFalse��$��2,0,0,0q�$��Trame��
�������$R  TrueY�	� -    $��RdBIsScriptTrue��$��2,0,0,0q�$��Script��
��������� $O 
Horizontal=�.+��� -!   $��Label18��$��50��
�������$��Right=�$��Data :���� -"   $��
TxtCmdData��$	��200��$��41��$S Code à envoyer��
�������
���   �
�������$V �� HorizontalScrollBarVisibility$
V Auto=�"W �� VerticalScrollBarVisibility$
W Auto=��� $O 
Horizontal=�.+��� -#   $��	BtnNewCmd��$��26.664��$��2q�$��67��
�������#��� & .��	Nouveau�� -$   $��
BtnSaveCmd��$��26.664��$��2q�$��67��
�������#��� & .��Enregistrer�� -%   $��	BtnDelCmd��$��26.664��$��2q�$��67��$y�	Collapsed=�
�������#��� & .��	Supprimer�� $O 
Horizontal=�.+��� -&   $��	BtnTstCmd��$��26.664��$��2q�$��67��$y�Hidden=�
�������#��� & .��Tester�� -'   $��BtnLearn��$��26.664��$��2q�$��67��
�������#��� & .��	Apprendre9� -(   $��ExpdVar��$��	Variables��
�������.���� -)   $��StkVar��$y�	Collapsed=�.+��� $O 
Horizontal=�.+��� -*   $��ListVar��$	��120��$��90���� .+��� -+   X ��Name$X 	BtnNewVar��$��+��Y J� RenderTransformOrigin$Y 0.5,0.5&�$��60��$
��30.5��$��3q�#��� & 
�������$
��Left=��� -,   $X BtnDeleteVar��$��-��$Y 0.5,0.5&�$��60��$
��30.5��$��3q�#��� & 
�������$
��Left=��� $O 
Horizontal=�$
��Left=�.+��� --   $X BtnModifVar��$��OK��$Y 0.5,0.5&�$��60��$
��30.5��$��3q�#��� & 
��������� -.   $X BtnAnnulVar��$��Annuler��$Y 0.5,0.5&�$��60��$
��30.5��$��3q�#��� & 
��������� -/   $��StkParamVar��$y�	Collapsed=�.+��� $O 
Horizontal=�.+��� $
��Nom:��
��������� -0   $��
TxtVarName��$	��120���� $O 
Horizontal=�.+��� $��Type:��
��������� -1   $��	CbVarType��$U 0��$��90��.@��� $��String���� $��Double���� $��Tableau���� $O 
Horizontal=�.+��� $��Valeur:��
��������� -2   $��	TxtVarVal��$	��120���� .+��� -3   $X buttonOk��$��Ok��$Y 0.5,0.5&�$��60��$
��30.5��$��3q�#��� & 
��������� -4   $X button��$��Fermer��$Y 0.5,0.5&�$��60��$
��30.5��$��3q�#��� & 
�������