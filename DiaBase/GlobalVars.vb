Imports System.Drawing.Text
Module GlobalVars
    'Attempt to keep globals to a minimum

    Public VersionAndRevision As String = "DIABASE 1.0"                         'Public Displayed Version And Revision Var PLEASE ADJUST BEFORE UPDATING REPO
    Public pfc As New PrivateFontCollection()                                   'Defines Custom Font Collection pfc As Global (Diablo Game Font)

    'DEFINE OBJECT ARRAY LISTS
    Public ItemObjects As List(Of ItemDatabase) = New List(Of ItemDatabase)             'Sets up new list of class items as "ItemObjects" Class Object used for main database
    Public UserObjects As List(Of UserListDatabase) = New List(Of UserListDatabase)     'Sets up new list of class items as "UserObjects" Class Object used for storing user list items

    Public SearchReferenceList As List(Of String) = New List(Of String)                 'Stores searched items main list index values - used to focus on correct item in the main list
    Public RefineSearchReferenceList As List(Of String) = New List(Of String)           'Stores temporary list of searched items index values - used when refining search lists
    Public UserListReferenceList As List(Of String) = New List(Of String)               'Stores list of user item index values - used to focus on correct item when the containing database is open
    Public StringMatches As List(Of String) = New List(Of String)                       'Stores a list of item index values that match string searches  - used by search routine
    Public IntegerMatches As List(Of String) = New List(Of String)                      'Stores a list of item index values that match integer searches - used by search routine
    Public OpenDatabaseDropDown As List(Of String) = New List(Of String)                'Stores a list of all created database filenames - Used when opening / closing / manipulating databases 
    Public DupeReferenceList As List(Of String) = New List(Of String)                   'Stored item index values of duped items - used to hide duplicated matching items in the search list

    'DEFINE APP SETTINGS VARIABLES
    Public AppSettings As New AppSetting
    Public HideSearchDupes As Boolean = False       'Used to Hide duplicate items during search - Is now Saved To Settings File
    Public SessionStartString As String = Nothing   'Used To Record And Hold The Start Time String Of The Current Session

    'DEFINE TIMER VARIABLES
    Public TimerSeconds As Integer = 0                      'Autologging Timer Seconds Counter Value
    Public Timercount As Integer = 0
    Public TimerRestart As Boolean = False                  'Bool to trigger timer restart after function calls
    Public ButtonFlashCount As Integer = 0                  'Upper value in seconds for button text flashing when time is paused
    Public AutoLoggerRunning As Boolean = False             'Bool to show the autologgers current runtime state - used to confirm imports are idle before running most routines and app functions       TRUE/FALSE
    Public AutoLoggerReady As Boolean = False               'Bool to reflect if the autologger is paused or counting down True = Counting Down, False = Paused, Default is False To Delay Countdonwe Until Setup Is Completed 

    'Define Vars used in autologger routine
    Public MuleLogPath As String = ""
    'Public DataBasePath As String = ""
    Public MuleDataPath As String = ""
    Public ArchiveFolder As String = ""
    Public LogFilesList As List(Of String) = New List(Of String)
    Public Pretotal As Integer = 0
    Public PassFiles As List(Of String) = New List(Of String)

    'Test only 
    Public UnDo As List(Of ItemDatabase) = New List(Of ItemDatabase) 'stores items deleted
    Public UnDoPos As List(Of String) = New List(Of String) 'stores location where deleted from
    Public UnDoCount As List(Of String) = New List(Of String) 'stores number deleted on each occasion

    'Search Pulldown Reference Lists Used to hold fields and entry values that exist in the database at runtime
    'this way lists are smaller, itterate quicker, and only have entries that will lead to valid searches
    'these must refersh every session as well as every database edit. Will setup routine in DatabaseManagmentFunctions module

    Public ItemNamePulldownList As List(Of String) = New List(Of String)
    Public UniqueAttribsPulldownList As List(Of String) = New List(Of String)
    Public UserReferencePulldownList As List(Of String) = New List(Of String)


    Public SuccessfulSearch As Boolean = False


    'itemedit var
    Public iEdit As Integer = 0
    'Skin Reference Array - Used to quick reference item skins as needed
    Public ImageArray As Array = {"hax", "axe", "2ax", "mpi", "wax", "lax", "bax", "btx", "gax", "gix", "wnd", "ywn", "bwn", "gwn", "clb", "scp", "gsc", "wsp", "spc", "mac", "mst", "fla",
                                    "whm", "mau", "gma", "ssd", "scm", "sbr", "flc", "crs", "bsd", "lsd", "wsd", "2hs", "clm", "gis", "bsw", "flb", "gsd", "dgr", "dir", "kri", "bld", "tkf",
                                    "tax", "bkf", "bal", "jav", "pil", "ssp", "glv", "tsp", "spr", "tri", "brn", "spt", "pik", "bar", "vou", "scy", "pax", "hal", "wsc", "sst", "lst", "cst",
                                    "bst", "wst", "sbw", "hbw", "lbw", "cbw", "sbb", "lbb", "swb", "lwb", "lxb", "mxb", "hxb", "rxb", "gps", "ops", "gpm", "opm", "gpl", "opl", "nia", "g33",
                                    "leg", "hdm", "hfh", "hst", "msf", "9ha", "9ax", "92a", "9mp", "9wa", "9la", "9ba", "9bt", "9ga", "9gi", "9wn", "9yw", "9bw", "9gw", "9cl", "9sc", "9qs",
                                    "9ws", "9sp", "9ma", "9mt", "9fl", "9wh", "9m9", "9gm", "9ss", "9sm", "9sb", "9fc", "9cr", "9bs", "9ls", "9wd", "92h", "9cm", "9gs", "9b9", "9fb", "9gd",
                                    "9dg", "9di", "9kr", "9bl", "9tk", "9ta", "9bk", "9b8", "9ja", "9pi", "9s9", "9gl", "9ts", "9sr", "9tr", "9br", "9st", "9p9", "9b7", "9vo", "9s8", "9pa",
                                    "9h9", "9wc", "8ss", "8ls", "8cs", "8bs", "8ws", "8sb", "8hb", "8lb", "8cb", "8s8", "8l8", "8sw", "8lw", "8lx", "8mx", "8hx", "8rx", "qf1", "qf2", "ktr",
                                    "wrb", "axf", "ces", "clw", "btl", "skr", "9ar", "9wb", "9xf", "9cs", "9lw", "9tw", "9qr", "7ar", "7wb", "7xf", "7cs", "7lw", "7tw", "7qr", "7ha", "7ax",
                                    "72a", "7mp", "7wa", "7la", "7ba", "7bt", "7ga", "7gi", "7wn", "7yw", "7bw", "7gw", "7cl", "7sc", "7qs", "7ws", "7sp", "7ma", "7mt", "7fl", "7wh", "7m7",
                                    "7gm", "7ss", "7sm", "7sb", "7fc", "7cr", "7bs", "7ls", "7wd", "72h", "7cm", "7gs", "7b7", "7fb", "7gd", "7dg", "7di", "7kr", "7bl", "7tk", "7ta", "7bk",
                                    "7b8", "7ja", "7pi", "7s7", "7gl", "7ts", "7sr", "7tr", "7br", "7st", "7p7", "7o7", "7vo", "7s8", "7pa", "7h7", "7wc", "6ss", "6ls", "6cs", "6bs", "6ws",
                                    "6sb", "6hb", "6lb", "6cb", "6s7", "6l7", "6sw", "6lw", "6lx", "6mx", "6hx", "6rx", "ob1", "ob2", "ob3", "ob4", "ob5", "am1", "am2", "am3", "am4", "am5",
                                    "ob6", "ob7", "ob8", "ob9", "oba", "am6", "am7", "am8", "am9", "ama", "obb", "obc", "obd", "obe", "obf", "amb", "amc", "amd", "ame", "amf", "cap", "skp",
                                    "hlm", "fhl", "ghm", "crn", "msk", "qui", "lea", "hla", "stu", "rng", "scl", "chn", "brs", "spl", "plt", "fld", "gth", "ful", "aar", "ltp", "buc", "sml",
                                    "lrg", "kit", "tow", "gts", "lgl", "vgl", "mgl", "tgl", "hgl", "lbt", "vbt", "mbt", "tbt", "hbt", "lbl", "vbl", "mbl", "tbl", "hbl", "bhm", "bsh", "spk",
                                    "xap", "xkp", "xlm", "xhl", "xhm", "xrn", "xsk", "xui", "xea", "xla", "xtu", "xng", "xcl", "xhn", "xrs", "xpl", "xlt", "xld", "xth", "xul", "xar", "xtp",
                                    "xuc", "xml", "xrg", "xit", "xow", "xts", "xlg", "xvg", "xmg", "xtg", "xhg", "xlb", "xvb", "xmb", "xtb", "xhb", "zlb", "zvb", "zmb", "ztb", "zhb", "xh9",
                                    "xsh", "xpk", "dr1", "dr2", "dr3", "dr4", "dr5", "ba1", "ba2", "ba3", "ba4", "ba5", "pa1", "pa2", "pa3", "pa4", "pa5", "ne1", "ne2", "ne3", "ne4", "ne5",
                                    "ci0", "ci1", "ci2", "ci3", "uap", "ukp", "ulm", "uhl", "uhm", "urn", "usk", "uui", "uea", "ula", "utu", "ung", "ucl", "uhn", "urs", "upl", "ult", "uld",
                                    "uth", "uul", "uar", "utp", "uuc", "uml", "urg", "uit", "uow", "uts", "ulg", "uvg", "umg", "utg", "uhg", "ulb", "uvb", "umb", "utb", "uhb", "ulc", "uvc",
                                    "umc", "utc", "uhc", "uh9", "ush", "upk", "dr6", "dr7", "dr8", "dr9", "dra", "ba6", "ba7", "ba8", "ba9", "baa", "pa6", "pa7", "pa8", "pa9", "paa", "ne6",
                                    "ne7", "ne8", "ne9", "nea", "drb", "drc", "drd", "dre", "drf", "bab", "bac", "bad", "bae", "baf", "pab", "pac", "pad", "pae", "paf", "neb", "neg", "ned",
                                    "nee", "nef", "nia", "nia", "nia", "nia", "nia", "vps", "yps", "rvs", "rvl", "wms", "tbk", "ibk", "amu", "vip", "rin", "nia", "bks", "bkd", "aqv", "tch",
                                    "cqv", "tsc", "isc", "hrt", "nia", "nia", "nia", "nia", "nia", "nia", "nia", "nia", "nia", "nia", "nia", "key", "nia", "xyz", "j34", "g34", "bbb", "box",
                                    "tr1", "mss", "ass", "qey", "qhr", "qbr", "ear", "gcv", "gfv", "gsv", "gzv", "gpv", "gcy", "gfy", "gsy", "gly", "gpy", "gcb", "gfb", "gsb", "glb", "gpb",
                                    "gcg", "gfg", "gsg", "glg", "gpg", "gcr", "gfr", "gsr", "glr", "gpr", "gcw", "gfw", "gsw", "glw", "gpw", "hp1", "hp2", "hp3", "hp4", "hp5", "mp1", "mp2",
                                    "mp3", "mp4", "mp5", "skc", "skf", "sku", "skl", "skz", "nia", "cm1", "cm2", "cm3", "nia", "nia", "nia", "nia", "r01", "r02", "r03", "r04", "r05", "r06",
                                    "r07", "r08", "r09", "r10", "r11", "r12", "r13", "r14", "r15", "r16", "r17", "r18", "r19", "r20", "r21", "r22", "r23", "r24", "r25", "r26", "r27", "r28",
                                    "r29", "r30", "r31", "r32", "r33", "jew", "ice", "nia", "tr2", "pk1", "pk2", "pk3", "dhn", "bey", "mbr", "toa", "tes", "ceh", "bet", "fed", "std"}


End Module
