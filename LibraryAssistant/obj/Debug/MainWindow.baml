   M S B A M L   `   `   `  ����K  GLibraryAssistant, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null3clr-namespace:LibraryAssistantLibraryAssistant  5   	   !    LibraryAssistant.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   0localclr-namespace:LibraryAssistant     5   	     � Title$5  /Подключение к базе данных��5   	   $	��325��5	   	   $	��600��6    � 
ResizeMode$ NoResize=�6"    � WindowStartupLocation$ CenterScreen=�5
   	   .��5      � .+�5   
   �� -   5         ��Name$ DatabaseConnectionTitle��5      $
��Left=�65   $	��Top=�6P   $��	10,10,0,0q�6h   $0��*Данные для подключения��6{   $��14(�6�   $
��Bold"�6�   �� 5   
   -   5         ��Name$! DatabaseConnectionHostLabel��5      $��	Хост:��65   $
��Left=�6E   $	��Top=�6`   $��	10,34,0,0q�6x   �� 5   
   -   5         ��Name$! DatabaseConnectionHostInput��5      $
��Left=�67   $	��Top=�6R   $��
152,38,0,0q�6j   $	��400��6~   �� 5   
   -   5        $! DatabaseConnectionPortLabel��5      $��	Порт:��65   $
��Left=�6E   $	��Top=�6`   $��	10,65,0,0q�6x   �� 5   
   -   5        $! DatabaseConnectionPortInput��5      $
��Left=�67   $��
152,69,0,0q�6R    �� TextWrapping$
 Wrap=�6f   $	��Top=�6z   $	��400��6�   6�   �� 5   
   -   5        $! DatabaseConnectionUserLabel��5      $��Пользователь:��65   $
��Left=�6M   $��9,96,0,0q�6h   $	��Top=�6z   $��26��6�   �� 5   
   -   5        $! DatabaseConnectionUserInput��5      $
��Left=�67   $��152,100,0,0q�6R   $	��Top=�6g   $	��400��6   �� 5   
   -   5        $% DatabaseConnectionPasswordLabel��5      $��Пароль:��69   $
��Left=�6K   $��
10,127,0,0q�6f   $	��Top=�6z   �� 5   
   -	   5        $% DatabaseConnectionPasswordInput��5      $
��Left=�6;   $��152,131,0,0q�6V   $	��Top=�6k   $	��400��6�   �� 5   
   -
   5        $) #DatabaseConnectionDatabaseNameLabel��5      $-��'Название базы данных:��6=   $
��Left=�6]   $��
10,158,0,0q�6x   $	��Top=�6�   �� 5   
   -   5        $) #DatabaseConnectionDatabaseNameInput��5      $
��Left=�6?   $��152,162,0,0q�6Z   $	��Top=�6o   $	��400��6�   �� 5   
   -   5         ��Name$% DatabaseConnectionConnectButton��5      $&�� Подключиться к БД��6:   $
��Left=�6V   $��
13,197,0,0q�6q   $	��Top=�6�   6�   �� 5   
   $
��Left=�6   $��
10,222,0,0q�6+   $��Ошибки:��6?   $	��Top=�6Q   �� 5   
   -   5        $$ DatabaseConnectionErrorMessage��5      $
��Left=�6:   $��152,226,0,0q�6U   $��Ошибок нет��6j   $	��Top=�6|   $	��400��6�   {�.  6�   5      5      