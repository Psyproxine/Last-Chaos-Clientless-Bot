Attribute VB_Name = "ModString"
Option Explicit
Public LimitSERV As Boolean
Public LimitEXE As Boolean
Public LimitLEVEL As Integer
Public User As String
Public Pass As String
Public Loc As String
Public Server As Integer
Public World As Integer
Public Char As Integer
Public Constate As Integer
Public RecvData As String
Public Active As Boolean
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public People() As PeopleInfo
Public stepwarp As Integer
Public sit_hp As Boolean
Public sit_hp_above As Integer
Public sit_hp_below As Integer
Public sit_mp As Boolean
Public sit_mp_above As Integer
Public sit_mp_below As Integer
Public MapScale As Double
Public delay_use_skill As Integer
Public delay_send_Move As Integer
Public delay_change_equip As Integer
Public delay_feed_pet As Integer
Public delay_select_npc As Integer
Public delay_use_mp As Integer
Public delay_use_hp As Integer
Public delay_refresh_inv As Integer
Public Logining As Boolean
Public Type Coord
    x As Long
    y As Long
End Type

Public Type PeopleInfo
    id As String * 4
    Pos As Coord
    PosX As Long
    PosY As Long
    NextPos As Coord
    time As Long
    nameid As Long
    speed As Integer
    Sex As String
    name As String
    account As String
    Hair As String
    Class As Integer
    Healed As Integer
    Online As Boolean
    Leader As Boolean
    Map As String
    OldMap As String
    hp As Long
    MaxHP As Long
    Guild As String
    EXP As String
    Position As String
    party As String
    LV As String
    Coord As String
    Shop As String
End Type

Type Monsters
    id As String
    name As String
    Attack As Integer
    Pos As Coord
    Distant As Integer
    skill As String
    Type As Long
    hp As Long
    MaxHP As Long
    Coord As String
    Atk_me As Boolean
    IsAttack As Boolean
    NoAttack As Boolean
    IsPet As Boolean
End Type

Type NPCs
    id As String
    name As String
    Pos As Coord
    Type As Long
    hp As Long
    MaxHP As Long
    Coord As String
    Agriculture As Boolean
End Type

Type Items
    id As String
    block As Long
    Durability As Integer
    LV As Long
    name As String
    Pos As Coord
    Amount As Long
    Type As Long
    Detail As String
    Refine As Long
    Coord As String
    equip As Integer
    State As Integer
    time As Date
    keep As Integer
    row As Integer
    col As Integer
End Type

Type Points
    id As Long
    Map As Long
    name As String
    Pos As Coord
    Coord As String
End Type

Type Skills
    id As String
    name As String
    LV As Integer
    Detail As String
End Type

Type tmpSkills
    Type As Long
    name As String
    Detail As String
End Type
 
Type locks
    auto As Boolean
    x As Integer
    y As Integer
    w As Integer
    h As Integer
    m As Integer
End Type
 
Type paths
    lock As locks
End Type

Type invites
    auto As Boolean
    name As String
    State As Integer
End Type

Type resps
    auto As Boolean
    name As String
End Type

Type follows
    auto As Boolean
    name As String
    Max As Integer
    Min As Integer
End Type

Type hps
    auto As Boolean
    name As String
    skill As Integer
    hp As Integer
    mp As Integer
End Type

Type protects
    auto As Boolean
    name As String
End Type

Type partys
    Invite As invites
    resp As resps
    follow As follows
    hp1 As hps
    hp2 As hps
    hp3 As hps
    hp4 As hps
    hp5 As hps
    mp1 As hps
    mp2 As hps
    mp3 As hps
    mp4 As hps
    mp5 As hps
    protect As protects
    atk As follows
End Type

Type basics
    timeout As Integer
    relogin As Integer
End Type

Type sit_hps
    auto As Boolean
    Min As Integer
    Max As Integer
End Type

Type heals
    sit_hp As sit_hps
    sit_mp As sit_hps
    hp1 As hps
    recon As hps
End Type
 
Type atks
    auto As Boolean
    skill As hps
End Type

Type Debugs
    show As Boolean
End Type

Type styles
    theme As Integer
End Type

Public Type opts
    atk As atks
    path As paths
    party As partys
    heal As heals
    basic As basics
    debug As Debugs
    style As styles
    farm As atks
    orc As atks
    pet As Boolean
    hp As hps
    mp As hps
End Type

Public Type Pets
    id As String
    name As String
    Type As String
    hp As Long
    MaxHP As Long
    EN As Long
    MaxEN As Long
    EXP As Long
    MaxEXP As Long
    FL As Long
    MaxFL As Long
    LV As Integer
    Pos As Coord
    Coord As String
    SP As Long
End Type

Public opt As opts
Public Monster() As Monsters
Public tmpMon() As Monsters
Public Item() As Items
Public tmpItem() As Items
Public Inv() As Items
Public sPoint() As Points
Public NPC() As NPCs
Public skill() As Skills
Public tmpSkill() As tmpSkills
Public C_Atk As Monsters
Public N_Atk As NPCs
Public delay_Mon_Atk As Integer
Public Sitting As Boolean
Public SelectedServ As Integer
Public SelectedChar As Integer
Public delay_invite_party As Integer
Public SSec As Integer, SMin As Integer, SHour As Integer, sExp As Long
Public Map As Integer, delay_go_map As Integer
Public pet As Pets
Public Apet() As Pets
