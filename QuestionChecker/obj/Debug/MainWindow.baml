   M S B A M L   `   `   `  ����M  IConsoleControl.WPF, Version=1.0.1.0, Culture=neutral, PublicKeyToken=nullS<clr-namespace:ConsoleControl.WPF;assembly=ConsoleControl.WPFConsoleControl.WPF  5   	   J FQuestionChecker, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:QuestionCheckerQuestionChecker 5   	       QuestionChecker.MainWindow   5   	   DWPF<clr-namespace:ConsoleControl.WPF;assembly=ConsoleControl.WPF  R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Z VFontAwesome.WPF, Version=4.7.0.37774, Culture=neutral, PublicKeyToken=0758b07a11a4f466-fa$http://schemas.fontawesome.io/icons/  5   	   Dmc;http://schemas.openxmlformats.org/markup-compatibility/2006  5   	   )localclr-namespace:QuestionChecker  5   	     � Title$  AnswerChecker��5	   	   $	��550��6   $	��900��6,    � WindowStartupLocation$ CenterScreen=�68   
�������.��5
      � ��5   
   �� 5       �� EndPoint$ 0.5,1&�6"    �� 
StartPoint$ 0.5,0&�63   .��5      �  � Color$ Black��6    � Offset$ 0Q�6-   � 5      $ White��6   $ 1Q�6-   5       � RowDefinitions 5   
   �� 5      $��90 �6   �� 5      .+�5   
   �� $��Center=�6   $��0��62   .+�5      �� 
���   �$��0,0,0,2q�6)   $��90��6C   $��Center=�6O   $��Center=�6l   $	��890��6�   ._�5      �  � ColumnDefinitions 5      �� 5      $	��470 �6+   �� 5      .+�5      � $��0��6   $��Center=�6+    5      �� 5      �� 5      .+�5!      ��  �� Orientation$ 
Horizontal=�6%   $��0��6>   .+�5"      �� -   5        $��txtFilePath��5"   2   $��40��6&   $	��313��6E   $��	10,5,10,0q�6Q   $��10(�6d   	 J� Opacity$	 1Q�6r   
 �� 
IsReadOnly
 . $��Center=�6�   $��Center=�6�    �� TextWrapping$
 Wrap=�6�   �� 5#      -   5        $��btnOpen��5#   <   $��Browse Excel��6%   $	��120��6K   $��15(�6W   $��30��6e   $
��Bold"�6q   6�   5%      �� 5&      $ 
Horizontal=�6%   $��1��6>   .+�5'      �� -   5        $��tBSheetName��5'   (   $��
10,15,10,0q�6;   $��15(�6O   $	 1Q�6]   $��Choose Sheet Name:��6i   $	��Top=�6�   
�������6�   �� 5(      -   5         ��Name$ cboSheetName��5(   '   $	��250��6=   $��33��6I   {�.  $��Center=�6g   6�   5)      5,      � 5-      $��1��6   $��Center=�6+    5.      �� 5/      �� 50      .+�52      �� $ 
Horizontal=�6%   $��0��6>   .+�53      �� -   5        $��tBLang��53   (   $��
10,15,10,0q�66   $��15(�6J   $	 1Q�6X   $��	Language:��6d   $	��Top=�6u   
�������6�   �� 54      -   5        $��txtLang��54   2   $��33��6&   $��60��6A   $��10(�6M   $	 1Q�6[   $�� ��6g   $��Center=�6o   $��Center=�6�   �� 55      -   5        $��tBUserid��55   (   $��
20,15,10,0q�68   $��15(�6L   $	 1Q�6Z   $��UserId:��6f   $	��Top=�6u   
�������6�   �� 56      -   5        $��	txtUserId��56   2   $��33��6&   $	��160��6C   $��10(�6P   $	 1Q�6^   $�� ��6j   $��Center=�6r   $��Center=�6�   57      �� 58      $ 
Horizontal=�6%   $��1��6>   $��Center=�6K   $��Center=�6h   .+�59      �� -	   5        $��btnPost��59   e   $��	Post Data��6%   $��15(�69   $��30��6G   $
��Bold"�6S   $��80��6t   $��0 0 20 0q�6   ���	#FFEFE5E5{�.  6�   �� 5:      -
   5        $��	btnExport��5:   b   $��Export��6%   $��15(�66   $��30��6D   $
��Bold"�6P   $��80��6s   $��20 0 0 0q�6~   ���	#FFEFE5E5{�.  6�   �� 5;      -   5         ��Name$ 
chkautoGen��5;   '   $��Auto Generate Result��6;   
�������$��15(�6r   $��Center=�6�   $��20 0 0 0q�6�   5<      5=      5>      5?      5@      � 5A   
   $��1��6   $��Center=�6   .+�5B      �� -   5        $��txtIns��5B      $��	Load Data��6"   $��15(�66   $
��Bold"�6D   $��15q�6V   $��Center=�6b   
�������6   ]� 5C      -   5         ]�Name$ panel_z��5C       K� ZIndex$
 1000��6)   
������ $y�	Collapsed=�6U   .+�5D      "  FontAwesome.WPF.ImageAwesome  -   5          Name$ 
spinner_ic��5D   "      Icon$ Spinner=�66      Spin .    SpinDuration$ 10Q�6Q   $	��100��6c   5E      � 5F      $	��310��6   $ 1��6    .+�5G      &   System.Windows.Controls.DataGrid  -   5        $��dtGrid��5G   e      AutoGenerateColumns .    
IsReadOnly . $��Center=�6H   $��Center=�6s   ��. $
��Auto=�5I      $
��Auto=�5J      5K      5L      �� 5N      -   5        $��	txtResult��5N      $��Result��6%   $��15(�66   $
��Bold"�6D   $��15q�6V   $��Center=�6b   $y�	Collapsed=�6�     5O      -   5          Name$ dgExcel��5O      $y�	Collapsed=�6(    .  . 6Z     5P      -   5        $ dgErrorRows��5P      $y�	Collapsed=�6,    .  . 6^   �� 5Q      -   5        $��txtCount��5Q      $��Row Count: ��6$   $��15(�6:   $
��Bold"�6H   $��10,355,20,-325q�6Z   $��Center=�6r   $y�Visible=�6�   5R      5U      5V      