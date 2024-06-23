VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fr 
      Caption         =   "Parameters"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   5055
      Begin VB.TextBox PrmL 
         Height          =   525
         Left            =   840
         ScrollBars      =   1  'Horizontal
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox PrmG 
         Height          =   525
         Left            =   840
         ScrollBars      =   1  'Horizontal
         TabIndex        =   2
         Top             =   960
         Width           =   3975
      End
      Begin VB.CheckBox FirstPage 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PrmL :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PrmG :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FirstPage :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DB :"
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label DB 
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "C :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   195
      End
      Begin VB.Label C 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Program:"
      BeginProperty Font 
         Name            =   "David"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Spool:"
      BeginProperty Font 
         Name            =   "David"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "David"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "David"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Dt 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Tm 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Pgr 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     "
      BeginProperty Font 
         Name            =   "David"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Spool 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Auto 
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SwBorder As Integer
Dim UserC As String
Dim Y As String
Dim Err_Sanderowit As String

Dim Conn As Object ' New ADODB.Connection
Dim Rs As Object ' New ADODB.Recordset
Dim Rs1 As Object ' New ADODB.Recordset
Dim dllPrm As Object '  As New NtvRptPrm.GetPrm

Dim SwSugDoch As String
Dim SanderowitzC As String

Dim Sdate As String
Dim Edate As String
Dim Stime As String
Dim Etime As String
Dim SParm As String

Dim TvTmp_Pratim As String
Dim pathNmVbInp As String
Dim pathVbInp As String

Dim sqlw As String
Dim sqll As String
Dim Tchumim As String
Dim TvTmpPrt_Prm As String
Dim TvTmpPrt As String

Dim aa As String
Dim DoSTR As Object
Dim Fs As Object
Dim CompanyCounter
Dim ArrSugAvera As String
Dim pathT, SERVER_NAME, HTTP_HOST, wDochStatus As String
Dim ErrLine As String
Private Sub Form_Initialize()
  Set DoSTR = CreateObject("Do_TchumimSTR.DoSTR")
  Set dllPrm = CreateObject("NtvRptPrm.GetPrm")
  Dt.Caption = Date
  Tm.Caption = Time
  Pgr.Caption = "סנדרוביץ"
End Sub
Public Sub Form_Load()

On Error GoTo Err_End
Sdate = CStr(Date)
Stime = CStr(Time())
ErrLine = "1"
Set Conn = CreateObject("ADODB.Connection")
Set Rs = CreateObject("ADODB.Recordset")
Set Rs1 = CreateObject("ADODB.Recordset")
Set Fs = CreateObject("FileSys.Main")
Set DoSTR = CreateObject("Do_TchumimSTR.DoSTR")
Set USpool = CreateObject("getSpool_Global.Main")


Spool_C = Trim(CStr(Command$))
App.Title = App.Title & "_" & Spool_C
SParm = USpool.GetData("Sanderowitz_Return", CStr(Spool_C))
If SParm = "" Then End

Main.Spool.Caption = USpool.getNmFromStr(SParm, "Spool")
Main.PrmL.Text = USpool.getNmFromStr(SParm, "PrmL")
Main.PrmG.Text = USpool.getNmFromStr(SParm, "PrmG")
Main.DB.Caption = USpool.getNmFromStr(SParm, "DB")
Main.C.Caption = USpool.getNmFromStr(SParm, "C")
Main.Auto.Caption = USpool.getNmFromStr(SParm, "Auto")

Set connstr = CreateObject("Build_ConnString.Main")
conn_S = connstr.bConnString_Lk(CStr(Trim(Main.DB.Caption)))
Conn.ConnectionTimeout = 300
Conn.Open conn_S
Set connstr = Nothing

Err_Sanderowit = ""

ErrLine = "2"
Rs.CursorLocation = 3
Conn.CommandTimeout = 500
Call SetSQL
Rs.ActiveConnection = Nothing

Err_End:
    swErr = 0
    If Err.Number <> 0 Or Err_Sanderowit <> "" Then
        swErr = 1
        ErrText = " Description=> " + Err.Description + Chr(13) + Chr(10) & _
                         " Err_Sanderowit=>" & Err_Sanderowit + Chr(13) + Chr(10) & " ErrLine=" & ErrLine
        Call USpool.ErrorReport(CStr(USpool.getNmFromStr(SParm, "C")), CStr(ErrText), "Sanderowitz_Return")
    Else
       If CStr(Trim(Auto.Caption)) = "1" Then Conn.Execute ("Update Spool set SwDone=1,SwHtml=1 WHERE C=" & Spool.Caption)
    End If
    
Edate = CStr(Date)
Etime = CStr(Time())

Call Update_Spool(Sdate, Edate, Stime, Etime)

Conn.Close

Set Rs = Nothing
Set Rs1 = Nothing
Set Conn = Nothing

Main.C.Caption = USpool.getNmFromStr(SParm, "C")
If swErr = 0 Then USpool.EndReport (Main.C.Caption)
End
End Sub
Private Function SetSQL() As String

    CompanyName = dllPrm.GetPrm(PrmG.Text, "CompanyName")
    CompanyKod = dllPrm.GetPrm(PrmG.Text, "Company")
    Adress = dllPrm.GetPrm(PrmG.Text, "CompanyAdress")
    City = dllPrm.GetPrm(PrmG.Text, "CompanyCity")
    UsrNm = dllPrm.GetPrm(PrmG.Text, "UsrName")
    DDate = dllPrm.GetPrm(PrmG.Text, "DDate")
    DTime = dllPrm.GetPrm(PrmG.Text, "DTime")
    OurLogo = dllPrm.GetPrm(PrmG.Text, "OurLogo")
    CompanyCounter = dllPrm.GetPrm(PrmG.Text, "CompanyCounter")
    CompanyHeader = dllPrm.GetPrm(PrmG.Text, "PrmCompany")
    Y = dllPrm.GetPrm(PrmG.Text, "Year")
    SwBorder = CInt(dllPrm.GetPrm(PrmG.Text, "SwBorder"))
    SnifCounter = dllPrm.GetPrm(PrmG.Text, "SnifCounter")
    UserC = dllPrm.GetPrm(PrmG.Text, "UsrCounter")
    
    SwSugDoch = dllPrm.GetPrm(PrmL.Text, "SwSugDoch")
    SanderowitzC = dllPrm.GetPrm(PrmL.Text, "SanderowitzC")
    pathNmVbInp = dllPrm.GetPrm(PrmL.Text, "pathNmVbInp")
    pathVbInp = dllPrm.GetPrm(PrmL.Text, "pathVbInp")
    '-------------------------------------------------------------------------------------------------------
    'Drop Tables
    
    TvTmpPrt = "Max2000_RunTime.[dbo].[Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption) & "]"
    
    FindTv = "Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption)
    sql = " SELECT * FROM Max2000_RunTime.sys.objects WHERE name = '" + FindTv + "'"

    Rs.Open sql, Conn
    If Not Rs.EOF Then
        sql = "  DROP TABLE " + TvTmpPrt
        Conn.Execute (sql), , adExecuteNoRecords
     End If
    Rs.Close
                  
    TvTmpPrt_Prm = "Max2000_RunTime.[dbo].[Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption) & "_Prm]"
  
    FindTv = "Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption) & "_Prm"
    sql = " SELECT * FROM Max2000_RunTime.sys.objects WHERE name = '" + FindTv + "'"

    Rs.Open sql, Conn
    If Not Rs.EOF Then
        sql = "  DROP TABLE " + TvTmpPrt_Prm
        Conn.Execute (sql), , adExecuteNoRecords
    End If
    Rs.Close
'============================================================================================================
    Set TblHanitaPratim = CreateObject("Hanita_Pratim.Main")
    TvTmp_Pratim = TblHanitaPratim.Main(CStr(DB.Caption), CStr(UserC), "RptProjects\Sanderowitz_Return")
    Set TblHanitaPratim = Nothing
   
    TvTmpPrt = "Max2000_RunTime.[dbo].[Sanderowitz_" & Trim(DB.Caption) & "_" & Trim(Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")) & "]"
    sql = "CREATE TABLE " + TvTmpPrt + " (" & _
                "[S] [int] IDENTITY(1,1) NOT NULL," & _
                "[CC] [int] NULL  ," & _
                ") ON [PRIMARY] "
     Conn.Execute (sql)
      
     sql = " insert into " & CStr(TvTmpPrt) & "(CC)" & _
                " select Sanderowitz_Lines.Doch " & _
                " from Gv_DochCar " & _
                " left join Sanderowitz_Lines on Sanderowitz_Lines.Doch=Gv_DochCar.C  " & _
                " left join ( " & _
                              " select DochC as DochC ," & _
                              " sum(isnull(Hiuv,0)) as Hiuv, sum(isnull(Zicui,0)) as Zicui,sum(isnull(Hiuv,0)-isnull(Zicui,0)) as Itra " & _
                              " from  Gv_View_DochCar_Tash Gv_View " & _
                              " group by DochC " & _
                            " ) View_Tas on View_Tas.DochC=Gv_DochCar.C " & _
                " where Sanderowitz_Lines.SanderowitzC=" & CStr(SanderowitzC) & " and isnull(View_Tas.Itra,0) > 0 and isnull(Gv_DochCar.OfiDoch,0)=3 "
    Conn.Execute (sql)
    SwNotCourt = "1"
    If CStr(SwSugDoch) = "0" Then SwNotCourt = "2"
     'כתיבת הקובץ
    Set Write_File = CreateObject("Sanderowitz_File.Main")
    ResFile = Write_File.Main(CStr(SwSugDoch), CStr(pathVbInp), CStr(pathNmVbInp), CStr(TvTmpPrt), "RptProjects\Sanderowitz_Return", UserC, DB.Caption, CStr(TvTmp_Pratim), CStr(SwNotCourt))
    Set Write_File = Nothing
    If ResFile <> "1" Then Err_Sanderowit = " Error in Write File"
         
    sql = " select  GDP.Tz as Tz, " & _
             " rtrim(ltrim(isnull(GDP.Nm,'')))+' '+rtrim(ltrim(isnull(GDP.NmF,''))) as GDPNm ," & _
             " isnull(rtrim(ltrim(convert(char,GDP.Street))),'') +' '+ isnull(rtrim(ltrim(convert(char,GDP.StreetNo))),'')+' '" & _
             " +  isnull(rtrim(ltrim(City.Nm)),'') collate SQL_Latin1_General_CP1255_CS_AS  as Adress ," & _
             " isnull(GDP.Mikod,'') as GDPMikod ," & _
             " rtrim(ltrim(convert(char,isnull(Gv_DochCar.Kod,'' ))))  + '-' +  rtrim(ltrim(convert(char,isnull(Gv_DochCar.Sidra,'')))) + '-' +  rtrim(ltrim(convert(char,isnull(Gv_DochCar.Bikoret,''))))as Kod ," & _
             " isnull(Gv_DochCar.Lochit,'') as Lochit ," & _
             " replace(convert(char,Gv_DochCar.D,101),' ','')  as DateAvera" & _
             " from Gv_DochCar " & _
             " inner join " & TvTmpPrt & " Tmp on Tmp.CC=Gv_DochCar.C " & _
             " left join " & TvTmp_Pratim & " GDP on GDP.DochCar=Gv_DochCar.C " & _
             " left join Max2000_Lib..City City  on City.C=GDP.City  "

    a = CreateTbl()

    sql = "Insert into " & TvTmpPrt & " (N1,N2,N3,N4,N5,N6,N7)" & sql
    Conn.CommandTimeout = 100
    Conn.Execute sql
End Function
Private Function CreateTbl() As Boolean
Dim SqlTbl As String
    
    CreateTbl = False
    d = DROPTbl()
  
    SqlTbl = "CREATE TABLE " + TvTmpPrt + " (" & _
            "[S] [int] IDENTITY(1,1) NOT NULL," & _
            "[K1] [numeric](18, 0)  NULL ," & _
            "[K1sort] [tinyint]  NULL ," & _
            "[SwSikum] [tinyint]  NULL ," & _
            "[K1Text] [char] (200) COLLATE Hebrew_CI_AS  NULL ," & _
            "[N1] [numeric](18, 0)  NULL  ," & _
            "[N1sort] [tinyint]  NULL ," & _
            "[N2]  [char] (100) COLLATE Hebrew_CI_AS  NULL ," & _
            "[N3]  [char] (100) COLLATE Hebrew_CI_AS  NULL ," & _
            "[N4] [numeric](18, 0)  NULL  ," & _
            "[N5] [char] (100) COLLATE Hebrew_CI_AS  NULL    ," & _
            "[N6] [char] (100) COLLATE Hebrew_CI_AS  NULL  ," & _
            "[N7] [datetime] NULL ," & _
            ") ON [PRIMARY] "
    Conn.Execute (SqlTbl), , adExecuteNoRecords
    
    C = cheeck_Tbl()
 
    NmString = "מס'' זהות,שם,כתובת,מיקוד,מס'' דוח,מס'' רישוי,ת.עבירה"
    Nm = Split(NmString, ",")
   
    SugString = "N,T,T,N,T,N,D"
    Sugi = Split(SugString, ",")
    
    FormatString = "RTL,RTL,RTL,RTL,RTL,RTL,RTL"
    FFormat = Split(FormatString, ",")
    
    FLenString = "10,10,10,5,20,8,9"
    Flen = Split(FLenString, ",")
    
    SikumString = "0,5,0,0,0,0,6"
    Sikum = Split(SikumString, ",")
    
    SikumGrpString = "0,0,0,0,0,0,0"
    SikumGrp = Split(SikumGrpString, ",")
     
    Company = CompanyCounter
    i = 0
    Do While i <= UBound(Nm)
        If i <> 0 Then Tchumim = ""
        If i <> 0 Then Company = 0
        sql = " insert into " & TvTmpPrt_Prm & "(Nm,Sug,FLen,Sort,Sikum,Format,SikumGrp,Company,Tchumim)" & _
              " values ('" & Nm(i) & "','" & Sugi(i) & "'," & Flen(i) & ",'N5'," & Sikum(i) & ",'" & FFormat(i) & "'," & SikumGrp(i) & "," & Company & ",'" & Replace(Trim(Tchumim), "'", "''") & "')"
        Conn.Execute sql
        i = i + 1
    Loop
    CreateTbl = True
End Function
Function DROPTbl()
   TvTmpPrt = "Max2000_RunTime.[dbo].[Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption) & "]"
    
    FindTv = "Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption)
    sql = " SELECT * FROM Max2000_RunTime.sys.objects WHERE name = '" + FindTv + "'"

    Rs.Open sql, Conn
    If Not Rs.EOF Then
        sql = "  DROP TABLE " + TvTmpPrt
        Conn.Execute (sql), , adExecuteNoRecords
    End If
    Rs.Close
End Function
Function cheeck_Tbl()
    TvTmpPrt_Prm = "Max2000_RunTime.[dbo].[Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption) & "_Prm]"
    
    FindTv = "Spool_" & Trim(DB.Caption) & "_" & Trim(Spool.Caption) & "_Prm"
    sql = " SELECT * FROM Max2000_RunTime.sys.objects WHERE name = '" + FindTv + "'"
 
    Rs.Open sql, Conn
    If Not Rs.EOF Then
      sql = "  DROP TABLE " + TvTmpPrt_Prm
      Conn.Execute (sql), , adExecuteNoRecords
    End If
    Rs.Close
   
    sql = " CREATE TABLE " & TvTmpPrt_Prm & "( " & _
          " [C] [int] IDENTITY(1,1) NOT NULL," & _
          " [Nm] [nchar](50) COLLATE Hebrew_CI_AS NULL, " & _
          " [Sug] [nchar](10) COLLATE SQL_Latin1_General_CP1255_CS_AS NULL, " & _
          " [FLen] [smallint] NULL, " & _
          " [Sort] [nchar](50) COLLATE SQL_Latin1_General_CP1255_CS_AS NULL, " & _
          " [Sikum] [nchar](100) COLLATE SQL_Latin1_General_CP1255_CS_AS NULL , " & _
          " [Format] [nchar](10) NULL, " & _
          " [SwSikumGrp] [smallint] NULL, " & _
          " [Company] [smallint] NULL, " & _
          " [Tchumim] [nchar](200) COLLATE SQL_Latin1_General_CP1255_CS_AS NULL, " & _
          " [SikumGrp] [smallint] NULL " & _
          " ) ON [PRIMARY]"
    
    Conn.Execute sql
End Function
Function DateSQlString(aDate) As String
    If InStr(1, aDate, ":") > 0 Then
      DateSQlString = "'" & aDate & "'"
       Exit Function
    End If
    If IsNull(aDate) Or aDate = "" Then
      DateSQlString = "null"
    Else
      DateSQlString = "'" & Format(Month(aDate), "00") & "/" & Format(Day(aDate), "00") & "/" & Year(aDate) & "'"
    End If
End Function
Function DateSQlStringEnd(aDate) As String
    If InStr(1, aDate, ":") > 0 Then
      DateSQlStringEnd = "'" & aDate & "'"
       Exit Function
    End If
    If IsNull(aDate) Or aDate = "" Then
      DateSQlStringEnd = "null"
    Else
      DateSQlStringEnd = "'" & Format(Month(aDate), "00") & "/" & Format(Day(aDate), "00") & "/" & Year(aDate) & " 23:59'"
    End If
End Function
Function DateSQl(aDate) As String
    If InStr(1, aDate, ":") > 0 Then
       DateSQl = "'" & aDate & "'"
       Exit Function
    End If
    If IsNull(aDate) Or aDate = "" Then
       DateSQl = "null"
    Else
       DateSQl = "'" & Format(Month(aDate), "00") & "/" & Format(Day(aDate), "00") & "/" & Year(aDate)
    End If
End Function
Public Sub Update_Spool(ByVal wSdate As String, ByVal wEdate As String, ByVal wStime As String, ByVal wEtime As String)
Dim sql As String
sql = " Update Spool " & _
        " set Sdate=" & DateSQl(wSdate) & " " & CStr(wStime) & "'," & _
        " Edate=" & DateSQl(wEdate) & " " & CStr(wEtime) & "'" & _
        " where C=" & CStr(Main.Spool.Caption)
        Conn.Execute (sql)
End Sub
