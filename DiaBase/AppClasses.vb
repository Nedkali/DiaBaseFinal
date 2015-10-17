﻿'Define Database Fields
'nothing else will go here

Public Class ItemDatabase
    Public ItemName As String
    Public ItemRealm As String
    Public Itemlevel As Integer
    Public ItemBase As String
    Public ItemQuality As String
    Public RequiredCharacter As String
    Public EtherealItem As Boolean
    Public Sockets As Integer
    Public RuneWord As Boolean
    Public ThrowDamageMin As Integer
    Public ThrowDamageMax As Integer
    Public OneHandDamageMin As Integer
    Public OneHandDamageMax As Integer
    Public TwoHandDamageMin As Integer
    Public TwoHandDamageMax As Integer
    Public Defense As Integer
    Public ChanceToBlock As Integer
    Public QuantityMin As Integer
    Public QuantityMax As Integer
    Public DurabilityMin As Integer
    Public DurabilityMax As Integer
    Public RequiredStrength As Integer
    Public RequiredDexterity As Integer
    Public RequiredLevel As Integer
    Public AttackClass As String
    Public AttackSpeed As String
    Public Stat1 As String
    Public Stat2 As String
    Public Stat3 As String
    Public Stat4 As String
    Public Stat5 As String
    Public Stat6 As String
    Public Stat7 As String
    Public Stat8 As String
    Public Stat9 As String
    Public Stat10 As String
    Public Stat11 As String
    Public Stat12 As String
    Public Stat13 As String
    Public Stat14 As String
    Public Stat15 As String
    Public MuleName As String
    Public MuleAccount As String
    Public MulePass As String
    Public PickitAccount As String
    Public HardCore As Boolean
    Public Ladder As Boolean
    Public Expansion As Boolean
    Public UserField As String
    Public ItemImage As Integer
    Public ImportDate As String
    Public ImportTime As String

End Class

Public Class AppSetting
    Public InstallPath As String = Application.StartupPath  'Applications Installation path
    Public EtalPath As String = "C:\D2NT"                   'Etals Installation Path.          
    Public SoundMute As Boolean = False                     'Mute Sound Setting Prefix.              True = Muted   False = On
    Public DefaultDatabase As String = Application.StartupPath & "\Databases\Default.txt"           'FileName (without extension) of the Database To load at startup
    Public AutoLoggingDelay As Integer = 30                 'Delay (in minuites) between automatic attempts to import item logs
    Public HideDupes As Boolean = False                     'Display duplicated items when diplaying search matches bool    TRUE/FALSE
    Public RemoveMuleDupes As Boolean = True                'Prefix to remove previously logged items from later imports - used to stop duplicated items relogging
    Public SaveOnExit As Boolean = False                    'Operator to auto save the current database as app closes       TRUE/FALSE
    Public BackupOnExit As Boolean = False                  'Operator to auto backup (After SaveOnExit) as app closes       TRUE/FALSE
    Public BackupBeforeImports As Boolean = False           'Backup before imprting item logs prefix        TRUE/FALSE
    Public BackupBeforeEdits As Boolean = False             'Backup before applying edits to item fields    TRUE/FALSE
    Public HideMulePass As Boolean = True                   'Prefix to hide the items account password      TRUE/FALSE
    Public CurrentDatabase As String = ""                   'variable to compare current versus Default database
    Public ResetDate As String = "26/4/2015"                'variable used for ressetting ladder to nonladder
    Public DefaultPassword As String = ""                   'variable used to replace Unknown passwords
    Public EastRealmCheckbox As Boolean = False             'Realm Search Checkboxes settings Var 
    Public WestRealmCheckbox As Boolean = False             'Realm Search Checkboxes settings Var 
    Public EuropeRealmCheckbox As Boolean = False           'Realm Search Checkboxes settings Var 
    Public AsiaRealmCheckbox As Boolean = False             'Realm Search Checkboxes settings Var 
End Class

Public Class UserListDatabase
    'Database Fields For Items Stored In The User List (Has to be independant of the main item database)
    'Any fields that are defined in the ItemDatabase object class must inturn also be defined here in this class
    Public ItemName As String
    Public ItemRealm As String
    Public Itemlevel As Integer
    Public ItemBase As String
    Public ItemQuality As String
    Public RequiredCharacter As String
    Public EtherealItem As Boolean
    Public Sockets As Integer
    Public RuneWord As Boolean
    Public ThrowDamageMin As Integer
    Public ThrowDamageMax As Integer
    Public OneHandDamageMin As Integer
    Public OneHandDamageMax As Integer
    Public TwoHandDamageMin As Integer
    Public TwoHandDamageMax As Integer
    Public Defense As Integer
    Public ChanceToBlock As Integer
    Public QuantityMin As Integer
    Public QuantityMax As Integer
    Public DurabilityMin As Integer
    Public DurabilityMax As Integer
    Public RequiredStrength As Integer
    Public RequiredDexterity As Integer
    Public RequiredLevel As Integer
    Public AttackClass As String
    Public AttackSpeed As String
    Public Stat1 As String
    Public Stat2 As String
    Public Stat3 As String
    Public Stat4 As String
    Public Stat5 As String
    Public Stat6 As String
    Public Stat7 As String
    Public Stat8 As String
    Public Stat9 As String
    Public Stat10 As String
    Public Stat11 As String
    Public Stat12 As String
    Public Stat13 As String
    Public Stat14 As String
    Public Stat15 As String
    Public MuleName As String
    Public MuleAccount As String
    Public MulePass As String
    Public PickitAccount As String
    Public HardCore As Boolean
    Public Ladder As Boolean
    Public Expansion As Boolean
    Public UserField As String
    Public ItemImage As Integer
    Public ImportDate As String
    Public ImportTime As String

    'note from ned - why not just add this to the other class? we can just leave it out during file read and write I dont think it will matter
    Public DatabaseFilename As String 'This field is unique to this class - Used to link items back to thier containing database in the user list
End Class