unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, bsSkinCtrls, ExtCtrls, ComCtrls, TFlatButtonUnit, registry,
  bsSkinData, BusinessSkinForm, TFlatListBoxUnit, StdCtrls, TFlatEditUnit,
  dxflchrt, bsMessages, bsSkinBoxCtrls, TFlatComboBoxUnit, TFlatMemoUnit,
  Buttons, flowchart, GridsEh, DBGridEh, ADODB, ShellAPI, DB, Comobj,
  TeeProcs, TeEngine, Chart, CommCtrl, newproject, TFlatProgressBarUnit,
  matrix, FileCtrl, Series, Mask, loadproject, StrUtils, Word2000,
  bsSkinGrids, ecodesign;

  {    BusinessSkinForm, Grids, bsSkinCtrls, bsSkinGrids, cxControls,
  cxSSheet, GridsEh, DBGridEh, DB,  bsDBGrids, StdCtrls, DBGrids,
  bsSkinBoxCtrls, TFlatButtonUnit, Buttons, ExtCtrls, matrix,  }

type
  TForm1 = class(TForm)
    MainMenu1: TMainMenu;
    wenjian: TMenuItem;
    LCA1: TMenuItem;
    LCA2: TMenuItem;
    N39: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N28: TMenuItem;
    N29: TMenuItem;
    N10: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    Panel2: TPanel;
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    Panel3: TPanel;
    LCA3: TMenuItem;
    LCA4: TMenuItem;
    N4: TMenuItem;
    LCI1: TMenuItem;
    LCIA1: TMenuItem;
    N12: TMenuItem;
    bsSkinMessage1: TbsSkinMessage;
    bsSkinData1: TbsSkinData;
    ADOConnection1: TADOConnection;
    ADOConnection2: TADOConnection;
    Label5: TLabel;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    SaveDialog1: TSaveDialog;
    ADOQuery2: TADOQuery;
    Label11: TLabel;
    bsSkinButton1: TbsSkinButton;
    bsSkinButton2: TbsSkinButton;
    bsSkinButton3: TbsSkinButton;
    bsSkinButton4: TbsSkinButton;
    bsSkinButton5: TbsSkinButton;
    bsSkinButton11: TbsSkinButton;
    bsSkinButton12: TbsSkinButton;
    TreeView1: TTreeView;
    Image1: TImage;
    ListBox1: TListBox;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet10: TTabSheet;
    TabSheet12: TTabSheet;
    Panel4: TPanel;
    Label3: TLabel;
    FlatEdit2: TFlatEdit;
    FlatComboBox1: TFlatComboBox;
    FlatEdit3: TFlatEdit;
    FlatEdit4: TFlatEdit;
    FlatEdit5: TFlatEdit;
    FlatEdit6: TFlatEdit;
    FlatEdit7: TFlatEdit;
    FlatMemo2: TFlatMemo;
    FlatMemo6: TFlatMemo;
    FlatEdit12: TFlatEdit;
    FlatEdit16: TFlatEdit;
    bsSkinComboBox1: TbsSkinComboBox;
    FlatEdit8: TFlatEdit;
    FlatEdit9: TFlatEdit;
    FlatMemo10: TFlatMemo;
    FlatMemo11: TFlatMemo;
    FlatEdit10: TFlatEdit;
    FlatEdit11: TFlatEdit;
    FlatEdit13: TFlatEdit;
    FlatEdit14: TFlatEdit;
    FlatEdit15: TFlatEdit;
    FlatEdit17: TFlatEdit;
    FlatEdit24: TFlatEdit;
    FlatMemo12: TFlatMemo;
    FlatMemo13: TFlatMemo;
    FlatMemo14: TFlatMemo;
    bsSkinButton15: TbsSkinButton;
    bsSkinButton16: TbsSkinButton;
    FlatMemo1: TFlatMemo;
    dxFlowChart1: TdxFlowChart;
    bsSkinButton6: TbsSkinButton;
    TabSheet13: TTabSheet;
    Label1: TLabel;
    bsSkinButton7: TbsSkinButton;
    bsSkinButton9: TbsSkinButton;
    bsSkinPanel1: TbsSkinPanel;
    Label12: TLabel;
    Edit2: TEdit;
    bsSkinButton17: TbsSkinButton;
    bsSkinButton34: TbsSkinButton;
    bsSkinButton37: TbsSkinButton;
    bsSkinButton38: TbsSkinButton;
    bsSkinButton39: TbsSkinButton;
    bsSkinPanel4: TbsSkinPanel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    bsSkinListBox1: TbsSkinListBox;
    DBGridEh3: TDBGridEh;
    bsSkinButton40: TbsSkinButton;
    bsSkinButton41: TbsSkinButton;
    bsSkinListBox2: TbsSkinListBox;
    bsSkinButton35: TbsSkinButton;
    OpenDialog1: TOpenDialog;
    DBGridEh9: TDBGridEh;
    bsSkinButton43: TbsSkinButton;
    bsSkinButton44: TbsSkinButton;
    bsSkinButton45: TbsSkinButton;
    bsSkinButton29: TbsSkinButton;
    bsSkinButton28: TbsSkinButton;
    bsSkinButton30: TbsSkinButton;
    Label2: TLabel;
    bsSkinCheckListBox9: TbsSkinCheckListBox;
    Label9: TLabel;
    RadioGroup1: TRadioGroup;
    Label13: TLabel;
    FlatEdit1: TFlatEdit;
    TabSheet3: TTabSheet;
    bsSkinButton48: TbsSkinButton;
    bsSkinButton46: TbsSkinButton;
    ListBox2: TListBox;
    TabSheet4: TTabSheet;
    TabSheet14: TTabSheet;
    FlatEdit20: TFlatEdit;
    FlatEdit23: TFlatEdit;
    bsSkinPanel6: TbsSkinPanel;
    Label19: TLabel;
    DBGridEh2: TDBGridEh;
    bsSkinButton51: TbsSkinButton;
    bsSkinButton52: TbsSkinButton;
    bsSkinButton54: TbsSkinButton;
    bsSkinPanel7: TbsSkinPanel;
    Label17: TLabel;
    DBGridEh4: TDBGridEh;
    bsSkinButton47: TbsSkinButton;
    bsSkinButton49: TbsSkinButton;
    bsSkinButton50: TbsSkinButton;
    ListBox3: TListBox;
    bsSkinButton58: TbsSkinButton;
    DBGridEh5: TDBGridEh;
    bsSkinButton59: TbsSkinButton;
    FlatEdit25: TFlatEdit;
    Label18: TLabel;
    FlatEdit26: TFlatEdit;
    DataSource2: TDataSource;
    DBGridEh6: TDBGridEh;
    FlatProgressBar1: TFlatProgressBar;
    bsSkinButton36: TbsSkinButton;
    Label24: TLabel;
    Label25: TLabel;
    Label27: TLabel;
    bsSkinPanel3: TbsSkinPanel;
    Edit1: TEdit;
    bsSkinButton20: TbsSkinButton;
    ListBox4: TListBox;
    bsSkinButton21: TbsSkinButton;
    bsSkinButton22: TbsSkinButton;
    bsSkinButton26: TbsSkinButton;
    bsSkinButton27: TbsSkinButton;
    bsSkinPanel2: TbsSkinPanel;
    Label8: TLabel;
    ComboBox2: TComboBox;
    CheckBox1: TCheckBox;
    DBGridEh1: TDBGridEh;
    Label4: TLabel;
    TabSheet8: TTabSheet;
    RadioButton5: TRadioButton;
    RadioButton6: TRadioButton;
    RadioButton7: TRadioButton;
    RadioButton8: TRadioButton;
    DBGridEh7: TDBGridEh;
    ADOQuery3: TADOQuery;
    DataSource3: TDataSource;
    Label26: TLabel;
    Label28: TLabel;
    ADOQuery4: TADOQuery;
    DataSource4: TDataSource;
    bsSkinButton13: TbsSkinButton;
    ListBox5: TListBox;
    bsSkinButton18: TbsSkinButton;
    bsSkinButton19: TbsSkinButton;
    Edit3: TEdit;
    Label6: TLabel;
    bsSkinButton10: TbsSkinButton;
    bsSkinButton23: TbsSkinButton;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    bsSkinButton24: TbsSkinButton;
    DataSource5: TDataSource;
    ADOQuery5: TADOQuery;
    Label29: TLabel;
    unitMat: TMatrix;
    PB2: TMatrix;
    Button1: TButton;
    g1_e: TMatrix;
    I_PB2: TMatrix;
    g2_e: TMatrix;
    RB2: TMatrix;
    gBCL: TMatrix;
    g3_e: TMatrix;
    g_onsite_e: TMatrix;
    OB: TMatrix;
    I_PB2_I: TMatrix;
    OB_WQ: TMatrix;
    OG: TMatrix;
    g6_e: TMatrix;
    g_outsite_e: TMatrix;
    g_route_e: TMatrix;
    g1: TMatrix;
    PA1: TMatrix;
    I_PA1: TMatrix;
    I_PA1_I: TMatrix;
    g2: TMatrix;
    g3: TMatrix;
    PA2: TMatrix;
    PA2_WQ: TMatrix;
    PE2: TMatrix;
    g4: TMatrix;
    RA: TMatrix;
    RA_WQ: TMatrix;
    RG: TMatrix;
    g_onsite: TMatrix;
    T: TMatrix;
    g5_1: TMatrix;
    juli: TMatrix;
    T_base: TMatrix;
    T_km: TMatrix;
    OF_1: TMatrix;
    PF2: TMatrix;
    g5: TMatrix;
    g5_e: TMatrix;
    g5_2: TMatrix;
    g6: TMatrix;
    OA: TMatrix;
    OA_WQ: TMatrix;
    g6_1: TMatrix;
    g6_2: TMatrix;
    PG2: TMatrix;
    g7_1: TMatrix;
    g7_2: TMatrix;
    RA_out: TMatrix;
    RA_out_WQ: TMatrix;
    RG_out: TMatrix;
    PH2: TMatrix;
    g7: TMatrix;
    g_outsite: TMatrix;
    g_route: TMatrix;
    char: TMatrix;
    g4_e: TMatrix;
    RB: TMatrix;
    RB_WQ: TMatrix;
    RG2: TMatrix;
    juli_e: TMatrix;
    T_km_e: TMatrix;
    OF_2: TMatrix;
    T_e: TMatrix;
    RB_out: TMatrix;
    RB_out_WQ: TMatrix;
    RG2_out: TMatrix;
    g7_e: TMatrix;
    bsSkinCheckListBox1: TbsSkinCheckListBox;
    DBGridEh8: TDBGridEh;
    bsSkinCheckListBox2: TbsSkinCheckListBox;
    ADOQuery6: TADOQuery;
    DataSource6: TDataSource;
    lci: TMatrix;
    LCIA: TMatrix;
    RadioButton9: TRadioButton;
    RadioButton10: TRadioButton;
    Button2: TButton;
    CheckBox4: TCheckBox;
    ComboBox4: TComboBox;
    CheckBox5: TCheckBox;
    ComboBox5: TComboBox;
    TabSheet9: TTabSheet;
    Panel5: TPanel;
    Label30: TLabel;
    PopupMenu1: TPopupMenu;
    N22: TMenuItem;
    RadioButton16: TRadioButton;
    LCC: TMatrix;
    Panel7: TPanel;
    bsSkinCheckListBox55: TbsSkinCheckListBox;
    ComboBox7: TComboBox;
    bsSkinCheckListBox8: TbsSkinCheckListBox;
    Label36: TLabel;
    dxFlowChart2: TdxFlowChart;
    Label37: TLabel;
    bsSkinCheckListBox5: TbsSkinCheckListBox;
    bsSkinButton31: TbsSkinButton;
    bsSkinButton32: TbsSkinButton;
    bsSkinButton33: TbsSkinButton;
    bsSkinButton42: TbsSkinButton;
    bsSkinButton60: TbsSkinButton;
    bsSkinButton61: TbsSkinButton;
    bsSkinButton62: TbsSkinButton;
    bsSkinButton63: TbsSkinButton;
    bsSkinButton64: TbsSkinButton;
    bsSkinButton65: TbsSkinButton;
    bsSkinButton66: TbsSkinButton;
    bsSkinButton67: TbsSkinButton;
    bsSkinButton68: TbsSkinButton;
    bsSkinButton69: TbsSkinButton;
    ADOQuery7: TADOQuery;
    PopupMenu2: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    ADOConnection3: TADOConnection;
    DBGridEh10: TDBGridEh;
    DataSource7: TDataSource;
    Chart1: TChart;
    CheckBox6: TCheckBox;
    Series1: TBarSeries;
    Series2: TPieSeries;
    OF2_WQ: TMatrix;
    OF_1_WQ: TMatrix;
    ADOConnection4: TADOConnection;
    ComboBox6: TComboBox;
    Shape1: TShape;
    Shape2: TShape;
    PE2_1: TMatrix;
    CheckBox7: TCheckBox;
    CheckBox8: TCheckBox;
    bsSkinEdit1: TbsSkinEdit;
    bsSkinButton25: TbsSkinButton;
    DataSource8: TDataSource;
    ADOQuery8: TADOQuery;
    bsSkinButton70: TbsSkinButton;
    PopupMenu3: TPopupMenu;
    Excel1: TMenuItem;
    N23: TMenuItem;
    ADOConExcel1: TADOConnection;
    Memo11: TMemo;
    Memo1: TMemo;
    RB2_WQ: TMatrix;
    bsSkinButton8: TbsSkinButton;
    bsSkinButton71: TbsSkinButton;
    ComboBox8: TComboBox;
    g2_unit: TMatrix;
    g1_unit: TMatrix;
    LCA5: TMenuItem;
    N5: TMenuItem;
    LCA6: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    DBGridEh11: TDBGridEh;
    Label31: TLabel;
    bsSkinEdit2: TbsSkinEdit;
    Label32: TLabel;
    Label33: TLabel;
    Label38: TLabel;
    bsSkinEdit3: TbsSkinEdit;
    ListBox6: TListBox;
    Label34: TLabel;
    ComboBox1: TComboBox;
    bsSkinButton14: TbsSkinButton;
    bsSkinButton72: TbsSkinButton;
    ListBox7: TListBox;
    Shape3: TShape;
    Shape6: TShape;
    bsSkinButton73: TbsSkinButton;
    ComboBox3: TComboBox;
    Label10: TLabel;
    bsSkinButton78: TbsSkinButton;
    ADOQuery9: TADOQuery;
    DataSource9: TDataSource;
    Scrape: TMatrix;
    Label39: TLabel;
    Label41: TLabel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    ComboBox9: TComboBox;
    Label7: TLabel;
    Shape4: TShape;
    TabSheet7: TTabSheet;
    Button3: TButton;
    bsSkinButton74: TbsSkinButton;
    bsSkinButton76: TbsSkinButton;
    DBGridEh12: TDBGridEh;
    Label43: TLabel;
    DBGridEh13: TDBGridEh;
    Chart2: TChart;
    ADOQuery10: TADOQuery;
    ADOQuery11: TADOQuery;
    DataSource10: TDataSource;
    DataSource11: TDataSource;
    ADOConnection5: TADOConnection;
    Label35: TLabel;
    Label40: TLabel;
    bsSkinButton77: TbsSkinButton;
    bsSkinButton56: TbsSkinButton;
    Series3: TBarSeries;
    N8: TMenuItem;
    bsSkinButton53: TbsSkinButton;
    N9: TMenuItem;
    N11: TMenuItem;
    N24: TMenuItem;
    N31: TMenuItem;
    ADOConnection6: TADOConnection;
    RadioButton3: TRadioButton;
    BPEI: TMatrix;
    TabSheet11: TTabSheet;
    RadioButton11: TRadioButton;
    RadioButton12: TRadioButton;
    bsSkinButton55: TbsSkinButton;
    FlatEdit19: TFlatEdit;
    FlatEdit21: TFlatEdit;
    FlatEdit22: TFlatEdit;
    DBGridEh14: TDBGridEh;
    ADOQuery12: TADOQuery;
    DataSource12: TDataSource;
    Label42: TLabel;
    Label44: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure bsSkinButton1Click(Sender: TObject);
    procedure bsSkinButton16Click(Sender: TObject);
    function AdobeReaderInstalled: boolean;
    function ListDirs(Path: string; List: TStringList): Integer;
    procedure bsSkinButton15Click(Sender: TObject);
    procedure bsSkinButton4Click(Sender: TObject);
    procedure ListBox4Click(Sender: TObject);
    procedure up_Format;
    procedure DBGridEh9GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure DBGridEh9KeyPress(Sender: TObject; var Key: Char);
    procedure ADOQuery1AfterOpen(DataSet: TDataSet); //up数据格式化
    procedure SetFdText(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure bsSkinButton12Click(Sender: TObject);
    procedure bsSkinButton5Click(Sender: TObject);
    procedure bsSkinButton11Click(Sender: TObject);
    procedure transport_base_Format; //运输基础数据 数据格式化
    procedure LCIfactor_Format; //LCIfactor数据格式化
    procedure char_Format; //特征化系数 数据格式化
    procedure shuxing_Format;
    procedure bsSkinButton20Click(Sender: TObject); //物料属性 数据格式化
    procedure SaveToWPSet(bg: TDBGrideh);
    procedure DBgridToExcel(Grid: TDBGrideh);
    procedure bsSkinButton27Click(Sender: TObject);
    procedure bsSkinButton21Click(Sender: TObject);
    procedure bsSkinButton22Click(Sender: TObject);
    procedure bsSkinButton26Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure bsSkinButton3Click(Sender: TObject);
    function chechedcount(bsSkinCheckListBox1: TbsSkinCheckListBox): integer;
  //  procedure bsSkinCheckListBox6ClickCheck(Sender: TObject);
  //  procedure CheckBox2Click(Sender: TObject);
 //   procedure CheckBox3Click(Sender: TObject);
 //   procedure bsSkinCheckListBox7ClickCheck(Sender: TObject);
    procedure bsSkinButton29Click(Sender: TObject);
    procedure ListBox1Renew;
    procedure TreeView1Renew;
    procedure TreeView1Click(Sender: TObject);
    procedure bsSkinButton6Click(Sender: TObject);
    procedure dxFlowChart1DblClick(Sender: TObject);
    procedure bsSkinButton9Click(Sender: TObject);
    procedure bsSkinButton7Click(Sender: TObject);
    procedure ProcessData_Format;
    procedure ListBox2Click(Sender: TObject);
  //  procedure shuxing_Format;
    procedure history_Format; //历史数据格式化
    procedure TransportData_Format;
    procedure bsSkinButton17Click(Sender: TObject);
    procedure bsSkinButton35Click(Sender: TObject);
    procedure bsSkinButton43Click(Sender: TObject);
    procedure bsSkinButton45Click(Sender: TObject);
    procedure bsSkinButton48Click(Sender: TObject);
    procedure DBGridEh1GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure DBGridEh1KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2Click(Sender: TObject);
    procedure bsSkinButton44Click(Sender: TObject);
    procedure bsSkinButton38Click(Sender: TObject);
    procedure bsSkinButton51Click(Sender: TObject);
    procedure bsSkinButton52Click(Sender: TObject);
    procedure bsSkinButton54Click(Sender: TObject);
    procedure bsSkinButton50Click(Sender: TObject);
    procedure bsSkinButton47Click(Sender: TObject);
    procedure ListBox3Click(Sender: TObject);
    procedure FlatEdit1Click(Sender: TObject);
    procedure allocation_Format;
    procedure bsSkinButton58Click(Sender: TObject);
    function checkedcount(bsSkinCheckListBox1: TbsSkinCheckListBox): integer;
    procedure RadioGroup1Click(Sender: TObject);
    procedure bsSkinButton59Click(Sender: TObject);
    procedure FlatEdit26Click(Sender: TObject);
    procedure bsSkinButton49Click(Sender: TObject);
    procedure bsSkinButton28Click(Sender: TObject);
    procedure SetFdText2(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure SetFdText9(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure SetFdText10(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure ADOQuery2AfterOpen(DataSet: TDataSet);
    procedure StatusBar1DrawPanel(StatusBar: TStatusBar;
      Panel: TStatusPanel; const Rect: TRect);
    function co_unit(process, s: string): string; //返回系数单位
    function co_num(process, num, s2: string): double; //返回系数
    function co_unit1(process, s: string): string; //返回系数单位
    function co_num1(process, num, s2: string): double; //返回系数
    function standard_unit(s2: string): string; //转换为标准单位
    function unit_exchange(s2: string): double;
    procedure FlatEdit27KeyPress(Sender: TObject; var Key: Char);
    procedure FlatEdit28KeyPress(Sender: TObject; var Key: Char);
    procedure FlatEdit29KeyPress(Sender: TObject; var Key: Char);
    procedure Label24Click(Sender: TObject);
    procedure bsSkinButton36Click(Sender: TObject);
    procedure bsSkinButton46Click(Sender: TObject);
    procedure Label25Click(Sender: TObject);
    procedure Label27Click(Sender: TObject); //返回单位换算系数
    procedure SetFdText4(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure SetFdText3(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure ADOQuery3AfterOpen(DataSet: TDataSet);
    procedure ADOQuery4AfterOpen(DataSet: TDataSet);
    procedure RadioButton5Click(Sender: TObject);
    procedure RadioButton6Click(Sender: TObject);
    procedure RadioButton7Click(Sender: TObject);
    procedure RadioButton8Click(Sender: TObject);
    procedure bsSkinButton10Click(Sender: TObject);
    procedure bsSkinButton13Click(Sender: TObject);
    procedure bsSkinButton18Click(Sender: TObject);
    procedure bsSkinButton19Click(Sender: TObject);
    procedure DBGridEh7CellClick(Column: TColumnEh);
    procedure Label6Click(Sender: TObject);
    procedure Label20Click(Sender: TObject);
    procedure bsSkinButton24Click(Sender: TObject);
    procedure bsSkinButton23Click(Sender: TObject);
    procedure ADOQuery5AfterOpen(DataSet: TDataSet);
    procedure Label21Click(Sender: TObject);
    procedure Label22Click(Sender: TObject);
    procedure Label23Click(Sender: TObject);
    procedure SetFdText5(Sender: TField; var Text: string; DisplayText: Boolean);
    function CO2(process: string): double;
    function energy(process: string): double;
    procedure direct(process: string);
    procedure Button1Click(Sender: TObject);
    procedure checkmat(mat: TMatrix); //调试用 检查矩阵
    procedure mat_energy;
    procedure mat_mainprocess;
    procedure PageControl1Change(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure ComboBox4Change(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure ADOQuery6AfterOpen(DataSet: TDataSet);
    procedure SetFdText6(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure ComboBox5Change(Sender: TObject);
    procedure DBGridEh8GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure RadioButton9Click(Sender: TObject);
    procedure RadioButton10Click(Sender: TObject);
    procedure bsSkinCheckListBox2ClickCheck(Sender: TObject);
    procedure bsSkinCheckListBox1ClickCheck(Sender: TObject);
    function CopyDir(SrcDir, DesDir: string): Boolean;
    procedure N15Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure Label30Click(Sender: TObject);
    procedure N22Click(Sender: TObject); //复制文件夹
    procedure PaintCtrlToBmp(Ctrl: TWinControl; Bmp: TBitmap);
    procedure RadioButton16Click(Sender: TObject);
    procedure N39Click(Sender: TObject);
    procedure LCA1Click(Sender: TObject);
    procedure LCA2Click(Sender: TObject);
    procedure ComboBox7Change(Sender: TObject);
    procedure bsSkinButton31Click(Sender: TObject);
    procedure bsSkinButton32Click(Sender: TObject);
    procedure bsSkinCheckListBox8ClickCheck(Sender: TObject);
    procedure bsSkinButton33Click(Sender: TObject);
    procedure CheckBox6Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure bsSkinButton42Click(Sender: TObject);
    procedure bsSkinButton60Click(Sender: TObject);
    procedure bsSkinButton61Click(Sender: TObject);
    procedure bsSkinButton62Click(Sender: TObject);
    procedure bsSkinButton63Click(Sender: TObject);
    procedure bsSkinButton64Click(Sender: TObject);
    procedure DBGridEh6TitleClick(Column: TColumnEh);
    procedure bsSkinButton65Click(Sender: TObject);
    procedure DBGridEh9TitleClick(Column: TColumnEh);
    procedure bsSkinButton69Click(Sender: TObject);
    procedure bsSkinButton68Click(Sender: TObject);
    procedure bsSkinButton66Click(Sender: TObject);
    procedure bsSkinButton67Click(Sender: TObject);
    procedure CheckBox7Click(Sender: TObject);
    procedure CheckBox8Click(Sender: TObject);
    function ReturnName(Value: pchar): string;
    procedure N18Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure LCA4Click(Sender: TObject);
    procedure DBGridEh7TitleClick(Column: TColumnEh);
    procedure DBGridEh8TitleClick(Column: TColumnEh);
    procedure bsSkinButton25Click(Sender: TObject);
    procedure DBGridEh4GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState); //返回文件夹名
    procedure SetFdText8(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure ADOQuery8AfterOpen(DataSet: TDataSet);
    procedure bsSkinButton70Click(Sender: TObject);
    procedure Excel1Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure bsSkinButton30Click(Sender: TObject);
    procedure bsSkinButton71Click(Sender: TObject);
    procedure ComboBox8Change(Sender: TObject);
    procedure LCA5Click(Sender: TObject);
    procedure N28Click(Sender: TObject);
    procedure N29Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure LCA6Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure LCI1Click(Sender: TObject);
    procedure LCIA1Click(Sender: TObject);
    procedure bsSkinButton72Click(Sender: TObject);
    procedure bsSkinButton73Click(Sender: TObject);
    procedure bsSkinButton14Click(Sender: TObject);
    procedure ListBox6Click(Sender: TObject);
    procedure ComboBox3Change(Sender: TObject);
    procedure ADOQuery9AfterOpen(DataSet: TDataSet);
    procedure bsSkinEdit2KeyPress(Sender: TObject; var Key: Char);
    procedure bsSkinEdit3KeyPress(Sender: TObject; var Key: Char);
    procedure Label39Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure ListBox7Click(Sender: TObject);
    procedure bsSkinButton78Click(Sender: TObject);
    procedure bsSkinCheckListBox5ClickCheck(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ADOQuery10AfterOpen(DataSet: TDataSet);
    procedure bsSkinButton76Click(Sender: TObject);
    procedure ADOQuery11AfterOpen(DataSet: TDataSet);
    procedure SetFdText11(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure bsSkinButton77Click(Sender: TObject);
    procedure bsSkinCheckListBox55ClickCheck(Sender: TObject);
    procedure bsSkinButton74Click(Sender: TObject);
    procedure bsSkinButton56Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure bsSkinButton2Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
    procedure bsSkinButton55Click(Sender: TObject);
    procedure ADOQuery12AfterOpen(DataSet: TDataSet);
    procedure SetFdText12(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure DBGridEh14TitleClick(Column: TColumnEh);
    procedure DBGridEh14GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure RadioButton11Click(Sender: TObject);
    procedure RadioButton12Click(Sender: TObject);


  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  flag1: boolean;
const
  pdunit: array[0..10] of string = ('千克(kg)', '吨(ton)', '万吨(10,000tons)', '千立方米(km3)', '立方米(m3)', '千瓦时(kWh)', '万千瓦时(10,000kWh)', '吉焦(GJ)', '兆焦(MJ)', '千焦(KJ)', '个(件、罐、台、根...)');




implementation

{$R *.dfm}
uses
  allocation;

procedure TForm1.FormCreate(Sender: TObject);

begin
  bsSkinData1.LoadFromFile(ExtractFilePath(Application.ExeName) + '\skin\19\skin.ini');
// bsSkinPanel1.Visible := true;
  ADOConnection1.Close;
  ADOConnection1.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'base.mdb;Persist Security Info=False';
  ADOConnection1.Open;

  Application.HintHidePause := 30000;

  TreeView1Renew;
  flag1 := false;
end;


procedure TForm1.TreeView1Renew;
var
  i: integer;
  rootnode: ttreenode;
begin

  TreeView1.Hint := '项目管理';
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  ListBox1Renew;
  TreeView1.Items.Clear;
  if ListBox1.Items.Count = 0 then
  begin
    PageControl1.Visible := false;
    exit;
  end;
  for i := 0 to ListBox1.Items.Count - 1 do
    rootnode := TreeView1.items.add(nil, ListBox1.Items[i]);


end;

procedure TForm1.ListBox1Renew; //
var
  s: TStringList;
begin
  ListBox1.Clear;
  s := TStringList.Create;
  ListDirs(ExtractFilePath(Application.ExeName) + '\projects\', s);
  ListBox1.Items.AddStrings(s);
  s.Free;
end;

procedure TForm1.bsSkinButton1Click(Sender: TObject);
begin
  Panel7.Visible := false;
  newprojectForm.ListBox1Renew;
 // bsSkinPanel3.Visible := false;
 // bsSkinPanel1.Visible := true;
 // ListBox1Renew;

  newprojectForm.ShowModal;
end;


procedure TForm1.bsSkinButton16Click(Sender: TObject);
begin
  Label5.Caption := '保存/下一步';
  bsSkinButton15Click(Sender);
  Label5.Caption := '保存';


  if FileExists(ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.flc') then
  begin
    dxFlowChart1.Clear;
    dxFlowChart1.LoadFromFile(ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.flc');
    dxFlowChart1.Zoom := 100;
    PageControl1.ActivePageIndex := 1;
    PageControl1.Pages[1].Highlighted := true;
    PageControl1.Pages[0].Highlighted := false;
  end
  else
  begin
    form2.TreeView1.Items.Clear;
    form2.TreeView1.LoadFromFile(ExtractFilePath(Application.ExeName) + '\tree.txt');
    form2.TreeView1.FullExpand;
    form2.dxFlowChart1.Align := alClient;
    form2.panel2.Visible := true;
    form2.bsskinpanel1.Visible := true;
    form2.dxFlowChart1.Clear;
    form2.ShowModal;
  end;





 { Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection2;
  bsSkinCheckListBox1.Items.Clear;
  bsSkinCheckListBox2.Items.Clear;
  bsSkinCheckListBox3.Items.Clear;
  bsSkinCheckListBox4.Items.Clear;
  bsSkinCheckListBox5.Items.Clear;
  try
    mysql := 'select * from LCIfactor';
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add(mysql);
    Aqr1.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('inf_type').AsString = '资源类' then
        bsSkinCheckListBox1.Items.Add(Aqr1.FieldByName('inf_name').AsString);

    end;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('inf_type').AsString = '大气污染物' then
        bsSkinCheckListBox3.Items.Add(Aqr1.FieldByName('inf_name').AsString);
    end;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('inf_type').AsString = '水体污染物' then
        bsSkinCheckListBox4.Items.Add(Aqr1.FieldByName('inf_name').AsString);
    end;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('inf_type').AsString = '固体废弃物' then
        bsSkinCheckListBox5.Items.Add(Aqr1.FieldByName('inf_name').AsString);
    end;


    mysql := 'select distinct impact_type from char_coef';
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add(mysql);
    Aqr1.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      bsSkinCheckListBox2.Items.Add(Aqr1.FieldByName('impact_type').AsString);
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;



  for i := 0 to bsSkinCheckListBox1.Items.Count - 1 do
    bsSkinCheckListBox1.Checked[i] := true;

  for i := 0 to bsSkinCheckListBox2.Items.Count - 1 do
    bsSkinCheckListBox2.Checked[i] := true;

  for i := 0 to bsSkinCheckListBox3.Items.Count - 1 do
    bsSkinCheckListBox3.Checked[i] := true;

  for i := 0 to bsSkinCheckListBox4.Items.Count - 1 do
    bsSkinCheckListBox4.Checked[i] := true;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
    bsSkinCheckListBox5.Checked[i] := true;

 // bsSkinPanel1.Visible := false;
  panel4.Visible := false;
  panel5.Align := alClient;
  panel5.Visible := true;
  label6.Caption := '<' + newprojectForm.FlatEdit1.Text + '> ' + '选择环境负荷因子/环境影响类型';
  label11.Caption := '选择指标';
 }
end;

function Tform1.AdobeReaderInstalled: boolean;
var
  r: TRegistry;
begin
  r := TRegistry.Create;
  try
    with r do
    begin
      RootKey := HKEY_LOCAL_MACHINE;
      OpenKey('\SOFTWARE\Adobe', False);
      if keyexists('Acrobat Reader') = True then
        Result := True
      else if keyexists('Adobe Acrobat') = True then
        Result := True
      else
        Result := False;
      CloseKey;
    end;
  finally
    r.Free;
  end;
end;





function TForm1.ListDirs(Path: string; List: TStringList): Integer;
//获取某一目录下所有文件夹名称
var
  FindData: TWin32FindData;
  FindHandle: THandle;
  FileName: string;
  AddToList: Boolean;
begin
  Result := 0;
  AddToList := Assigned(List);

  if Path[Length(Path)] <> '\' then
    Path := Path + '\';

  Path := Path + '*.*';

  FindHandle := Windows.FindFirstFile(PChar(Path), FindData);
  while FindHandle <> INVALID_HANDLE_VALUE do
  begin
    FileName := StrPas(FindData.cFileName);
    if (FileName <> '.') and (FileName <> '..') and
      ((FindData.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) then
    begin
      Inc(Result);
      if AddToList then
        List.Add(FileName);
    end;

    if not Windows.FindNextFile(FindHandle, FindData) then
      FindHandle := INVALID_HANDLE_VALUE;
  end;
  Windows.FindClose(FindHandle);
end;



procedure TForm1.bsSkinButton15Click(Sender: TObject); //保存项目基本信息
var
  Aqr1: TAdoQuery;
  mysql: string;
begin
  ADOConnection2.Close;
  ADOConnection2.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.mdb ; Persist Security Info=False';
  ADOConnection2.Open;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection2;
  try
    mysql := 'select * from project_inf';
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add(mysql);
    Aqr1.Open;

    Aqr1.Locate('project_item', 'LCA类别', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatComboBox1.Text;
    Aqr1.Post;




    Aqr1.Locate('project_item', '项目建立时间', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit3.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '项目所属系统', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit23.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', 'LCA联系人信息', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit4.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '产品名称（最后工序）', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit6.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '产品系统描述', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatMemo6.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '功能单位数据', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit8.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '功能单位', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := bsSkinComboBox1.Text;
    Aqr1.Post;


    Aqr1.Locate('project_item', '生产数据时间', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit9.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '数据取舍原则', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatMemo10.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '运输', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit11.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '上游数据', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit14.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '数据质量评价', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatEdit17.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '数据分配', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatMemo13.Text;
    Aqr1.Post;

    Aqr1.Locate('project_item', '循环环境收益', [loCaseInsensitive]);
    Aqr1.Edit;
    Aqr1.FieldByName('item_inf').AsString := FlatMemo12.Text;
    Aqr1.Post;

  finally
    Aqr1.Close;
    Aqr1.Free;
  end;


  if Label5.Caption = '保存' then
    bsSkinMessage1.MessageDlg('  保存完成  ', (mtInformation), [mbOK], 0);

end;

procedure TForm1.bsSkinButton4Click(Sender: TObject); //上游数据管理
var
  i: integer;
  Aqr1: TAdoQuery;
  rootnode: ttreenode;
begin
  Panel7.Visible := false;
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  TreeView1.Hint := '基础信息管理';
  TreeView1.Items.Clear;
  PageControl1.Visible := true;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := false;

  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := false;
  rootnode := TreeView1.items.add(nil, '基础信息管理');
  TreeView1.Items.Addchild(rootnode, '外部LCA数据管理');
  TreeView1.Items.Addchild(rootnode, '运输LCA基础数据管理');
  TreeView1.Items.Addchild(rootnode, 'LCI环境指标管理');
  TreeView1.Items.Addchild(rootnode, '环境影响类型管理');

  rootnode := TreeView1.items.add(nil, '对照表管理');
  TreeView1.Items.Addchild(rootnode, '物料属性名称对照表');
  TreeView1.Items.Addchild(rootnode, '运输名称对照表');
  TreeView1.Items.Addchild(rootnode, '上游数据名称对照表');
  TreeView1.Items.Addchild(rootnode, 'LCI因子名称对照表');
  rootnode := TreeView1.items.add(nil, '工序对标数据管理');

  TreeView1.FullExpand;
  TreeView1.Items[1].Selected := true;
  TreeView1.SetFocus;
  StatusBar1.Panels[1].Text := '基础信息管理： ' + TreeView1.Selected.Text;

  Image1.Visible := false;
  bsSkinPanel2.Visible := true; //功能单位

  Label4.Caption := '外部LCA数据管理';

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection1;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct product from up');
    Aqr1.Open;
    ListBox4.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      ListBox4.Items.Add(Aqr1.FieldByName('product').AsString);
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;

  AdoQuery1.Close;
  ListBox4Click(Sender);
end;

procedure TForm1.ListBox4Click(Sender: TObject);
begin
  if ListBox4.ItemIndex = -1 then exit;

  if Label4.Caption = '外部LCA数据管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    up_Format;
  end;

  if Label4.Caption = '运输LCA基础数据管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    transport_base_Format;
  end;

  if Label4.Caption = 'LCI环境指标管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    LCIfactor_Format;
  end;
  if Label4.Caption = '环境影响类型管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    char_Format;
  end;
  if Label4.Caption = '物料能源属性管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    shuxing_Format;
  end;

  if Label4.Caption = Label1.Caption + '1' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection2;
    up_Format;
  end;

  if Label4.Caption = Label1.Caption + '2' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection2;
    transport_base_Format;
  end;

  if Label4.Caption = Label1.Caption + '3' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection2;
    char_Format;
  end;

  if Label4.Caption = Label1.Caption + '4' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection2;
    LCIfactor_Format;
  end;

  if Label4.Caption = '工序对标数据管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    history_Format;
  end;


end;


procedure TForm1.shuxing_Format; //物料属性 数据格式化
var
  i: integer;
  mysql1: string;
begin
  AdoQuery8.Connection := Form1.ADOConnection2;
  mysql1 := 'select ID,name,shuxing,danwei,shuzhi from shuxing';
  if AdoQuery8.Active then AdoQuery8.Close;
  AdoQuery8.SQL.Clear;
  AdoQuery8.SQL.Add(mysql1);
  ADOQuery8.Open;
  ADOQuery8.Active := true;

  for i := 0 to DBGridEh4.Columns.Count - 1 do
  begin
    DBGridEh4.Columns[i].Width := 80;
    DBGridEh4.Columns[i].Title.Alignment := taCenter;
    DBGridEh4.Columns[i].Alignment := taLeftJustify;
  end;

  DBGridEh4.Columns[0].Alignment := taCenter;
  DBGridEh4.Columns[3].Alignment := taCenter;
  DBGridEh4.Columns[4].Alignment := taRightJustify; ;

  DBGridEh4.Columns[0].Title.caption := '序号';
  DBGridEh4.Columns[1].Title.caption := '物料名称';
  DBGridEh4.Columns[2].Title.caption := '属性';
  DBGridEh4.Columns[3].Title.caption := '单位';
  DBGridEh4.Columns[4].Title.caption := '数量';


  DBGridEh4.Columns[0].Width := 40;
  DBGridEh4.Columns[2].Width := 180;
  DBGridEh4.Columns[4].Width := 100;



  DBGridEh4.Columns.Items[3].PickList.clear;
  DBGridEh4.Columns.Items[3].picklist.Add('MJ/kg');
  DBGridEh4.Columns.Items[3].picklist.Add('MJ/m3');
  DBGridEh4.Columns.Items[3].picklist.Add('kg/kg');
  DBGridEh4.Columns.Items[3].picklist.Add('kg/m3');
  DBGridEh4.Columns.Items[3].picklist.Add('g/kg');
  DBGridEh4.Columns.Items[3].picklist.Add('g/m3');
  DBGridEh4.Columns.Items[3].picklist.Add('%');


  DBGridEh4.Columns.Items[2].PickList.clear;
  DBGridEh4.Columns.Items[2].picklist.Add('CO2排放系数');
  DBGridEh4.Columns.Items[2].picklist.Add('热值');
  DBGridEh4.Columns.Items[2].picklist.Add('电力构成');
  DBGridEh4.Columns.Items[2].picklist.Add('折标煤系数');
  DBGridEh4.Columns.Items[2].picklist.Add('铁含量');


  ADOQuery8.First;

end;

procedure TForm1.char_Format; //特征化系数 数据格式化
var
  i: integer;
  mysql1: string;
  Aqr1: TAdoQuery;
begin
  mysql1 := 'select ID,impact_type ,eq_unit,LCI_factor , equivalent from char_coef where impact_type =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39;
  if AdoQuery1.Active then AdoQuery1.Close;
  AdoQuery1.SQL.Clear;
  AdoQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;
  ADOQuery1.Active := true;

  for i := 0 to DBGridEh1.Columns.Count - 1 do
  begin
    DBGridEh1.Columns[i].Width := 80;
    DBGridEh1.Columns[i].Title.Alignment := taCenter;
    DBGridEh1.Columns[i].Alignment := taLeftJustify;
  end;

  DBGridEh1.Columns[0].Alignment := taCenter;
  DBGridEh1.Columns[4].Alignment := taRightJustify; ;

  DBGridEh1.Columns[0].Title.caption := '序号';
  DBGridEh1.Columns[1].Title.caption := '类别';
  DBGridEh1.Columns[2].Title.caption := '当量单位';
  DBGridEh1.Columns[3].Title.caption := 'LCI因子';
  DBGridEh1.Columns[4].Title.caption := '当量系数';


  DBGridEh1.Columns[0].Width := 40;
  DBGridEh1.Columns[2].Width := 100;
  DBGridEh1.Columns[3].Width := 180;
  DBGridEh1.Columns[4].Width := 100;

  DBGridEh1.Columns.Items[1].PickList.clear;
  DBGridEh1.Columns.Items[1].picklist.Add(ListBox4.Items[ListBox4.ItemIndex]);
  DBGridEh1.Columns.Items[3].PickList.clear;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection1;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('Select * from  LCIfactor ');
    Aqr1.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      DBGridEh1.Columns.Items[3].picklist.Add(Aqr1.FieldByName('inf_name').AsString);
    end;

  finally
    Aqr1.Close;
    Aqr1.Free;
  end;



  ADOQuery1.First;

end;

procedure TForm1.history_Format; //历史数据格式化
var
  i: integer;
  mysql1: string;
begin

  mysql1 := ' select * from historydata where source =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39;


  if ADOQuery1.Active then ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;

  for i := 0 to DBGridEh1.Columns.Count - 1 do
  begin
    DBGridEh1.Columns[i].Width := 80;
    DBGridEh1.Columns[i].Title.Alignment := taCenter;
    DBGridEh1.Columns[i].Alignment := taLeftJustify;
  end;

  DBGridEh1.Columns[0].Alignment := taCenter;
  DBGridEh1.Columns[4].Alignment := taRightJustify; ;

  DBGridEh1.Columns[0].Title.caption := '序号';
  DBGridEh1.Columns[1].Title.caption := '来源';
  DBGridEh1.Columns[2].Title.caption := '工序';
  DBGridEh1.Columns[3].Title.caption := '类别';
  DBGridEh1.Columns[4].Title.caption := '单位';
  DBGridEh1.Columns[5].Title.caption := '数量';
  DBGridEh1.Columns[6].Title.caption := '备注';



  DBGridEh1.Columns[0].Width := 40;
  DBGridEh1.Columns[1].Width := 120;
  DBGridEh1.Columns[2].Width := 100;
  DBGridEh1.Columns[3].Width := 100;
  DBGridEh1.Columns[4].Width := 60;
  DBGridEh1.Columns[5].Width := 60;
  DBGridEh1.Columns[6].Width := 150;

  DBGridEh1.Columns.Items[3].PickList.clear;
  DBGridEh1.Columns.Items[3].picklist.Add('工序能耗');
  DBGridEh1.Columns.Items[3].picklist.Add('碳排放');




  ADOQuery1.First;

end;




procedure TForm1.LCIfactor_Format; //LCIfactor数据格式化
var
  i: integer;
  mysql1: string;
begin
  if listbox4.ItemIndex = 0 then
    mysql1 := 'select ID, inf_type,inf_name,units,price, see_level  from  LCIfactor  '
  else
    mysql1 := 'select ID, inf_type,inf_name,units,price,see_level  from  LCIfactor where inf_type =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39;

  if AdoQuery1.Active then AdoQuery1.Close;
  AdoQuery1.SQL.Clear;
  AdoQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;
  ADOQuery1.Active := true;
  ADOQuery1.Sort := 'see_level ASC';

  for i := 0 to DBGridEh1.Columns.Count - 1 do
  begin
    DBGridEh1.Columns[i].Width := 80;
    DBGridEh1.Columns[i].Title.Alignment := taCenter;
    DBGridEh1.Columns[i].Alignment := taLeftJustify;
  end;

  DBGridEh1.Columns[0].Alignment := taCenter;
  DBGridEh1.Columns[3].Alignment := taCenter;
  DBGridEh1.Columns[4].Alignment := taCenter;

  DBGridEh1.Columns[0].Title.caption := '序号';
  DBGridEh1.Columns[1].Title.caption := '类别';
  DBGridEh1.Columns[2].Title.caption := 'LCI因子';
  DBGridEh1.Columns[3].Title.caption := '单位';
  DBGridEh1.Columns[4].Title.caption := 'LCC单价（元）';

  DBGridEh1.Columns[0].Width := 40;
  DBGridEh1.Columns[2].Width := 160;
  DBGridEh1.Columns[1].Width := 100;
  DBGridEh1.Columns[4].Width := 100;
  DBGridEh1.Columns[5].Visible := false;

  DBGridEh1.Columns.Items[1].PickList.clear;
  if listbox4.ItemIndex = 0 then
  begin
    DBGridEh1.Columns.Items[1].picklist.Add('资源类');
    DBGridEh1.Columns.Items[1].picklist.Add('能源类');
    DBGridEh1.Columns.Items[1].picklist.Add('大气污染物');
    DBGridEh1.Columns.Items[1].picklist.Add('水体污染物');
    DBGridEh1.Columns.Items[1].picklist.Add('固体废弃物');
  end
  else
    DBGridEh1.Columns.Items[1].picklist.Add(ListBox4.Items[ListBox4.ItemIndex]);

  ADOQuery1.First;

end;

procedure TForm1.transport_base_Format; //运输基础数据 数据格式化
var
  i: integer;
  mysql1: string;
begin
  mysql1 := 'select * from transport_base ';
  if AdoQuery1.Active then AdoQuery1.Close;
  AdoQuery1.SQL.Clear;
  AdoQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;
  ADOQuery1.Active := true;

  for i := 0 to DBGridEh1.Columns.Count - 1 do
  begin
    DBGridEh1.Columns[i].Width := 80;
    DBGridEh1.Columns[i].Title.Alignment := taCenter;
    DBGridEh1.Columns[i].Alignment := taLeftJustify;
  end;

  DBGridEh1.Columns[0].Alignment := taCenter;
  DBGridEh1.Columns[3].Alignment := taRightJustify;
  DBGridEh1.Columns[4].Alignment := taRightJustify;
  DBGridEh1.Columns[5].Alignment := taRightJustify;
  DBGridEh1.Columns[6].Alignment := taRightJustify;

  DBGridEh1.Columns[0].Title.caption := '序号';
  DBGridEh1.Columns[1].Title.caption := 'LCI因子';
  DBGridEh1.Columns[2].Title.caption := '单位';
  DBGridEh1.Columns[3].Title.caption := '海轮';
  DBGridEh1.Columns[4].Title.caption := '火车';
  DBGridEh1.Columns[5].Title.caption := '卡车';
  DBGridEh1.Columns[6].Title.caption := '驳船';

  DBGridEh1.Columns[0].Width := 40;
  DBGridEh1.Columns[1].Width := 180;
  DBGridEh1.Columns[4].Width := 100;

  DBGridEh1.Columns.Items[2].PickList.clear;
  DBGridEh1.Columns.Items[2].picklist.Add('KJ/t.km');
  DBGridEh1.Columns.Items[2].picklist.Add('g/t.km');
  DBGridEh1.Columns.Items[2].picklist.Add('mg/t.km');
  DBGridEh1.Columns.Items[2].picklist.Add('ml/t.km');


  ADOQuery1.First;

end;

procedure TForm1.up_Format; //up数据格式化
var
  i: integer;
  mysql1: string;
  Aqr1: TAdoQuery;
begin


  mysql1 := 'select ID,product,LCItype,LCIfactor,unit,sum,function_unit,see_level from  up where product=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39;
  if AdoQuery1.Active then AdoQuery1.Close;
  AdoQuery1.SQL.Clear;
  AdoQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;
  ADOQuery1.Active := true;
  ADOQuery1.Sort := 'ID ASC';

  DBGridEh1.Columns[1].Visible := false;
  DBGridEh1.Columns[6].Visible := false;
  DBGridEh1.Columns[7].Visible := false;
  for i := 0 to DBGridEh1.Columns.Count - 1 do
  begin
    DBGridEh1.Columns[i].Width := 80;
    DBGridEh1.Columns[i].Title.Alignment := taCenter;
    DBGridEh1.Columns[i].Alignment := taCenter;
  end;

  DBGridEh1.Columns[5].Alignment := taRightJustify; //数据
  DBGridEh1.Columns[2].Alignment := taLeftJustify;
  DBGridEh1.Columns[3].Alignment := taLeftJustify;

  DBGridEh1.Columns[0].Title.caption := '序号';
  DBGridEh1.Columns[2].Title.caption := '类别';
  DBGridEh1.Columns[3].Title.caption := 'LCI因子';
  DBGridEh1.Columns[4].Title.caption := '单位';
  DBGridEh1.Columns[5].Title.caption := '数量';

  DBGridEh1.Columns[0].Width := 40;
  DBGridEh1.Columns[2].Width := 100;
  DBGridEh1.Columns[3].Width := 160;

  ComboBox2.ItemIndex := ComboBox2.Items.IndexOf(ADOQuery1.FieldByName('function_unit').AsString);

  DBGridEh1.Columns.Items[2].PickList.clear;
  DBGridEh1.Columns.Items[2].picklist.Add('资源类');
  DBGridEh1.Columns.Items[2].picklist.Add('能源类');
  DBGridEh1.Columns.Items[2].picklist.Add('大气污染物');
  DBGridEh1.Columns.Items[2].picklist.Add('水体污染物');
  DBGridEh1.Columns.Items[2].picklist.Add('固体废弃物');


  DBGridEh1.Columns.Items[3].PickList.clear;


  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection1;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from  LCIfactor ');
    Aqr1.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if AdoQuery1.Locate('LCIfactor', Aqr1.FieldByName('inf_name').AsString, []) = false then
        DBGridEh1.Columns.Items[3].picklist.Add(Aqr1.FieldByName('inf_name').AsString);

    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;


  ADOQuery1.First;

end;


procedure TForm1.DBGridEh9GetCellParams(Sender: TObject; Column: TColumnEh;
  AFont: TFont; var Background: TColor; State: TGridDrawState);
begin
  if not (gdSelected in State) then
    if (Sender as TDBGridEh).SumList.RecNo mod 2 = 0 then
      Background := clInfobk;
end;

procedure TForm1.DBGridEh9KeyPress(Sender: TObject; var Key: Char);
begin

  if key = #13 then
  begin
    keybd_event(vk_tab, 0, 0, 0);
    keybd_event(vk_tab, 0, keyeventf_keyup, 0);
  end;

  if dbgrideh9.SelectedField = adoquery1.FieldByName('processtype') then //只能选择类别
    if key <> #13 then //允许使用回车
      key := #0; //什么也不干

  if dbgrideh9.SelectedField = adoquery1.FieldByName('process') then //只能选择类别
    if key <> #13 then //允许使用回车
      key := #0; //什么也不干

  if dbgrideh9.SelectedField = adoquery1.FieldByName('material_type') then //只能选择类别
    if key <> #13 then //允许使用回车
      key := #0; //什么也不干

  if dbgrideh9.SelectedField = adoquery1.FieldByName('pd_unit') then //只能选择类别
    if key <> #13 then //允许使用回车
      key := #0; //什么也不干



end;


procedure TForm1.SetFdText(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  if ADOQuery1.Active then
  begin
    Text := inttostr(ADOQuery1.RecNo);
    if ADOQuery1.RecordCount = 0 then
      Text := '0';

  end;
end;

procedure TForm1.SetFdText11(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery11.RecNo);
  if ADOQuery11.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.SetFdText2(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery2.RecNo);
  if ADOQuery2.RecordCount = 0 then
    Text := '0';
end;


procedure TForm1.SetFdText3(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery3.RecNo);
  if ADOQuery3.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.SetFdText4(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery4.RecNo);
  if ADOQuery4.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.SetFdText5(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery5.RecNo);
  if ADOQuery5.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.SetFdText6(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery6.RecNo);
  if ADOQuery6.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.SetFdText12(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery12.RecNo);
  if ADOQuery12.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.SetFdText8(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery8.RecNo);
  if ADOQuery8.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.SetFdText9(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery9.RecNo);
  if ADOQuery9.RecordCount = 0 then
    Text := '0';
end;


procedure TForm1.SetFdText10(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  Text := inttostr(ADOQuery10.RecNo);
  if ADOQuery10.RecordCount = 0 then
    Text := '0';
end;

procedure TForm1.ADOQuery1AfterOpen(DataSet: TDataSet);
var
  i: integer;
begin
  ADOQuery1.FieldByName('ID').OnGetText := SetFdText;
  for i := 1 to ADOQuery1.FieldCount - 1 do
    if ADOQuery1.Fields[i].DataType = ftFloat then
      (ADOQuery1.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0';

  ADOQuery1.First;

 { if TreeView1.Hint = '基础信息管理' then
  begin
    if Label4.Caption = '外部LCA数据管理' then
      (ADOQuery1.FieldByName('sum') as Tfloatfield).DisplayFormat := '###,###,##0.##0';
    if Label4.Caption = '环境影响类型管理' then
      (ADOQuery1.FieldByName('equivalent') as Tfloatfield).DisplayFormat := '###,###,##0.##0';
    if Label4.Caption = '运输LCA基础数据管理' then
      (ADOQuery1.FieldByName('shuzhi') as Tfloatfield).DisplayFormat := '###,###,##0.##0';

    if Label4.Caption = '物料能源属性管理' then
      (ADOQuery1.FieldByName('shuzhi') as Tfloatfield).DisplayFormat := '###,###,##0.##0';



  end;

   if TreeView1.Hint = '项目管理' then
      begin
      for i:=1 to ADOQuery1.FieldCount-1 do
        if ADOQuery1.Fields[i].DataType=ftfloat then
          (ADOQuery1.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0';



      end;
                 }


end;

procedure TForm1.bsSkinButton12Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;
  rootnode: ttreenode;
begin
  Panel7.Visible := false;
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  TreeView1.Hint := '基础信息管理';
  TreeView1.Items.Clear;
  PageControl1.Visible := true;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := false;
  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := false;

  rootnode := TreeView1.items.add(nil, '基础信息管理');
  TreeView1.Items.Addchild(rootnode, '外部LCA数据管理');
  TreeView1.Items.Addchild(rootnode, '运输LCA基础数据管理');
  TreeView1.Items.Addchild(rootnode, 'LCI环境指标管理');

  TreeView1.Items.Addchild(rootnode, '环境影响类型管理');

  rootnode := TreeView1.items.add(nil, '对照表管理');
  TreeView1.Items.Addchild(rootnode, '物料属性名称对照表');
  TreeView1.Items.Addchild(rootnode, '运输名称对照表');
  TreeView1.Items.Addchild(rootnode, '上游数据名称对照表');
  TreeView1.Items.Addchild(rootnode, 'LCI因子名称对照表');
  rootnode := TreeView1.items.add(nil, '工序对标数据管理');
  TreeView1.FullExpand;
  TreeView1.Items[4].Selected := true;
  TreeView1.SetFocus;
  StatusBar1.Panels[1].Text := '基础信息管理： ' + TreeView1.Selected.Text;

  Image1.Visible := false;
  bsSkinPanel2.Visible := false;
  label4.Caption := '环境影响类型管理';

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection1;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct impact_type from char_coef');
    Aqr1.Open;
    ListBox4.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      ListBox4.Items.Add(Aqr1.FieldByName('impact_type').AsString);
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;

  ADOQuery1.Close;
  ListBox4Click(Sender);
end;

procedure TForm1.bsSkinButton5Click(Sender: TObject);
var
  i: integer;
  rootnode: ttreenode;
begin
  Panel7.Visible := false;
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  TreeView1.Hint := '基础信息管理';
  TreeView1.Items.Clear;
  PageControl1.Visible := true;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := false;
  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := false;

  rootnode := TreeView1.items.add(nil, '基础信息管理');
  TreeView1.Items.Addchild(rootnode, '外部LCA数据管理');
  TreeView1.Items.Addchild(rootnode, '运输LCA基础数据管理');
  TreeView1.Items.Addchild(rootnode, 'LCI环境指标管理');

  TreeView1.Items.Addchild(rootnode, '环境影响类型管理');

  rootnode := TreeView1.items.add(nil, '对照表管理');
  TreeView1.Items.Addchild(rootnode, '物料属性名称对照表');
  TreeView1.Items.Addchild(rootnode, '运输名称对照表');
  TreeView1.Items.Addchild(rootnode, '上游数据名称对照表');
  TreeView1.Items.Addchild(rootnode, 'LCI因子名称对照表');
  rootnode := TreeView1.items.add(nil, '工序对标数据管理');

  TreeView1.FullExpand;
  TreeView1.Items[2].Selected := true;
  TreeView1.SetFocus;
  StatusBar1.Panels[1].Text := '基础信息管理： ' + TreeView1.Selected.Text;

  Image1.Visible := false;
  bsSkinPanel2.Visible := false;

  label4.Caption := '运输LCA基础数据管理';

  ListBox4.Clear;
  ListBox4.Items.Add('全部');
  ListBox4.Items.Add('海轮');
  ListBox4.Items.Add('火车');
  ListBox4.Items.Add('卡车');
  ListBox4.Items.Add('驳船');

  AdoQuery1.Close;
  ListBox4Click(Sender);
end;

procedure TForm1.bsSkinButton11Click(Sender: TObject);
var
  rootnode: ttreenode;
  i: integer;
begin
  Panel7.Visible := false;
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  TreeView1.Hint := '基础信息管理';
  TreeView1.Items.Clear;
  PageControl1.Visible := true;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := false;
  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := false;

  rootnode := TreeView1.items.add(nil, '基础信息管理');
  TreeView1.Items.Addchild(rootnode, '外部LCA数据管理');
  TreeView1.Items.Addchild(rootnode, '运输LCA基础数据管理');
  TreeView1.Items.Addchild(rootnode, 'LCI环境指标管理');

  TreeView1.Items.Addchild(rootnode, '环境影响类型管理');

  rootnode := TreeView1.items.add(nil, '对照表管理');
  TreeView1.Items.Addchild(rootnode, '物料属性名称对照表');
  TreeView1.Items.Addchild(rootnode, '运输名称对照表');
  TreeView1.Items.Addchild(rootnode, '上游数据名称对照表');
  TreeView1.Items.Addchild(rootnode, 'LCI因子名称对照表');
  rootnode := TreeView1.items.add(nil, '工序对标数据管理');

  TreeView1.FullExpand;
  TreeView1.Items[3].Selected := true;
  TreeView1.SetFocus;
  StatusBar1.Panels[1].Text := '基础信息管理： ' + TreeView1.Selected.Text;

  Image1.Visible := false;
  bsSkinPanel2.Visible := false;

  label4.Caption := 'LCI环境指标管理';

  ListBox4.Clear;
  ListBox4.Items.Add('全部');
  ListBox4.Items.Add('资源类');
  ListBox4.Items.Add('能源类');
  ListBox4.Items.Add('大气污染物');
  ListBox4.Items.Add('水体污染物');
  ListBox4.Items.Add('固体废弃物');

  ADOQuery1.Close;
  ListBox4Click(Sender);

end;

procedure TForm1.bsSkinButton20Click(Sender: TObject);
var
  Aqr1: TAdoQuery;
  mysql1: string;
  i: integer;
begin
  if trim(Edit1.Text) = '' then
  begin
    bsSkinMessage1.MessageDlg('名称不能为空', (mtInformation), [mbOK], 0);
    edit1.SetFocus;
    exit;
  end;

  if ListBox4.Items.IndexOf(trim(Edit1.Text)) >= 0 then
  begin
    bsSkinMessage1.MessageDlg('该名称已存在', (mtInformation), [mbOK], 0);
    edit1.SetFocus;
    exit;
  end;

  listbox4.ClearSelection;
  ListBox4.Items.Add(trim(Edit1.Text));
  ListBox4.ItemIndex := listbox4.Items.IndexOf(trim(Edit1.Text));
  listbox4.Selected[listbox4.Items.IndexOf(trim(Edit1.Text))] := true;

  if Label4.Caption = '外部LCA数据管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    mysql1 := 'select ID,product,LCItype,LCIfactor,unit,sum,function_unit,see_level from  up where product=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39;
    if AdoQuery1.Active then AdoQuery1.Close;
    AdoQuery1.SQL.Clear;
    AdoQuery1.SQL.Add(mysql1);
    ADOQuery1.Open;
    ADOQuery1.Active := true;

    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select * from  LCIfactor ');
      Aqr1.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        ADOQuery1.Append;
        ADOQuery1.FieldByName('product').AsString := trim(Edit1.Text);
        ADOQuery1.FieldByName('LCItype').AsString := Aqr1.FieldByName('inf_type').AsString;
        ADOQuery1.FieldByName('LCIfactor').AsString := Aqr1.FieldByName('inf_name').AsString;
        ADOQuery1.FieldByName('unit').AsString := Aqr1.FieldByName('units').AsString;
        ADOQuery1.FieldByName('see_level').AsString := Aqr1.FieldByName('see_level').AsString;
        ADOQuery1.Post;
      end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
    ADOQuery1.UpdateBatch(arall);
    up_Format;
  end;

  if Label4.Caption = Label1.Caption + '1' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection2;
    mysql1 := 'select ID,product,LCItype,LCIfactor,unit,sum,function_unit,see_level from  up where product=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39;
    if AdoQuery1.Active then AdoQuery1.Close;
    AdoQuery1.SQL.Clear;
    AdoQuery1.SQL.Add(mysql1);
    ADOQuery1.Open;
    ADOQuery1.Active := true;

    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select * from  LCIfactor ');
      Aqr1.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        ADOQuery1.Append;
        ADOQuery1.FieldByName('product').AsString := trim(Edit1.Text);
        ADOQuery1.FieldByName('LCItype').AsString := Aqr1.FieldByName('inf_type').AsString;
        ADOQuery1.FieldByName('LCIfactor').AsString := Aqr1.FieldByName('inf_name').AsString;
        ADOQuery1.FieldByName('unit').AsString := Aqr1.FieldByName('units').AsString;
        ADOQuery1.FieldByName('see_level').AsString := Aqr1.FieldByName('see_level').AsString;
        ADOQuery1.Post;
      end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
    ADOQuery1.UpdateBatch(arall);
    up_Format;

  end;


  if Label4.Caption = '环境影响类型管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;
    char_Format;
  end;

  if Label4.Caption = '工序对标数据管理' then
  begin
    AdoQuery1.Connection := Form1.ADOConnection1;

    AdoQuery1.Append;
    AdoQuery1.FieldByName('source').AsString := trim(Edit1.Text);
    AdoQuery1.Post;
  end;

end;


procedure TForm1.SaveToWPSet(bg: TDBGrideh);
var
  i, row, column, j: integer;
  etapp: OleVariant; //定义金山表格对象
  myworkbook: OleVariant; //定义金山表格的工作簿对象
begin

  if bg.DataSource.DataSet.RecordCount = 0 then exit;
  try
    etapp := CreateOleObject('et.Application');
  except
    form1.bsSkinMessage1.MessageDlg('未安装 Excel \ WPS 或其它原因，保存失败！', (mtInformation), [mbOK], 0);
    Exit;
  end;


  etapp.Workbooks.Close;
  myworkbook := etapp.Workbooks.add;
  myworkbook.WorkSheets['sheet1'].Activate;
  column := 1;
  for j := 0 to bg.FieldCount - 1 do
  begin
    if bg.Columns[j].Visible = true then
    begin
      myworkbook.worksheets['sheet1'].cells[1, column].value := bg.columns.Items[j].Title.caption;
      column := column + 1;
    end;
  end;
  row := 2;
  bg.DataSource.DataSet.First;

  while not (bg.DataSource.DataSet.Eof) do
  begin
    column := 1;
    for i := 0 to bg.Columns.Count - 1 do
    begin
      if BG.Columns[I].Visible = true then
      begin
        myworkbook.worksheets['sheet1'].cells[row, column].value := bg.Columns[i].Field.AsString;
        column := column + 1;
      end;
    end;
    bg.DataSource.DataSet.Next;
    row := row + 1;
  end;
  form1.bsSkinMessage1.MessageDlg('保存完成！', (mtInformation), [mbOK], 0);
  etapp.ActiveWorkbook.saveas(SaveDialog1.FileName + '.et');
  etapp.ActiveWorkbook.Close;
  etapp.quit;
  etapp := Unassigned;

//  end;




end;


procedure TForm1.DBgridToExcel(Grid: TDBGrideh); //输出到Excel
var
  v, sheet: Variant;
  i, j: integer;
begin

  SaveDialog1.Filter := '*.xls|*.xls|*.et|*.et';

  if SaveDialog1.Execute then
  begin
    if SaveDialog1.FilterIndex = 1 then
    begin

      try
        v := CreateOleObject('Excel.Application'); //创建ole对象
      except
        form2.SaveDialog.Filter := '*.et';
        SaveToWPSet(Grid);
        exit;
      end;



      v.WorkBooks.Add;
      Sheet := v.WorkBooks[1].WorkSheets[1];
      Grid.DataSource.DataSet.First;
      i := 1;
      while not (Grid.DataSource.DataSet.EOF) do
      begin
        for j := 1 to Grid.FieldCount do
          if Grid.Columns[j - 1].Visible = True then
          begin
            Sheet.Cells[i + 1, j] := Grid.Fields[j - 1].AsString;

            if Grid.Fields[j - 1].FieldName = 'ID' then
              Sheet.Cells[i + 1, j] := Grid.DataSource.DataSet.RecNo;

      //+1保留列标题
            Sheet.Cells[1, j] := Grid.Columns[j - 1].Title.caption;

          end;
        Grid.DataSource.DataSet.Next;
        i := i + 1;
      end;


      v.WorkBooks[1].saveas(SaveDialog1.FileName + '.xls');
      v.WorkBooks[1].Close;
      v.quit;
      v := Unassigned;
      form1.bsSkinMessage1.MessageDlg('保存完成！', (mtInformation), [mbOK], 0);
    end;
    if SaveDialog1.FilterIndex = 2 then
      SaveToWPSet(Grid);
  end;
end;




procedure TForm1.bsSkinButton27Click(Sender: TObject);
var
  i: integer;
begin
  if DBGridEh1.SelectedRows.Count = 0 then
  begin
    bsSkinMessage1.MessageDlg('请选择需要删除的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if bsSkinMessage1.MessageDlg('确定删除所选记录？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := 0 to DBGridEh1.SelectedRows.Count - 1 do
    begin
      DBGridEh1.DataSource.DataSet.GotoBookmark(pointer(DBGridEh1.SelectedRows[i]));
      DBGridEh1.DataSource.DataSet.Delete;
    end;
  end;

  ADOQuery1.UpdateBatch(arall);

  if Label4.Caption = '外部LCA数据管理' then
    up_Format;
  if Label4.Caption = 'LCI环境指标管理' then
    LCIfactor_Format;

end;

procedure TForm1.bsSkinButton21Click(Sender: TObject);
var
  Aqr1: TAdoQuery;
  i: integer;
begin

  if trim(Edit1.Text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请输入查询名称中至少一个字符！', (mtInformation), [mbOK], 0);
    edit1.SetFocus;
    exit;
  end;


  if Label4.Caption = '外部LCA数据管理' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from up where product like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关产品！', (mtInformation), [mbOK], 0);
        edit1.SetFocus;
        exit;
      end;
      listbox4.ClearSelection;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          listbox4.Selected[listbox4.Items.IndexOf(Aqr1.FieldByName('product').AsString)] := true;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;



  if Label4.Caption = Label1.Caption + '1' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection2;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from up where product like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关产品！', (mtInformation), [mbOK], 0);
        edit1.SetFocus;
        exit;
      end;
      listbox4.ClearSelection;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          listbox4.Selected[listbox4.Items.IndexOf(Aqr1.FieldByName('product').AsString)] := true;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;



  if Label4.Caption = 'LCI环境指标管理' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from LCIfactor where inf_name like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关环境负荷因子！', (mtInformation), [mbOK], 0);
        DBGridEh1.Selection.Clear;
        edit1.SetFocus;
        exit;
      end;

      DBGridEh1.Selection.Clear;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          if ADOQuery1.Locate('inf_name', Aqr1.FieldByName('inf_name').AsString, []) then
            DBGridEh1.SelectedRows.CurrentRowSelected := True;
          DBGridEh1.SetFocus;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;

  if Label4.Caption = Label1.Caption + '4' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection2;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from LCIfactor where inf_name like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关环境负荷因子！', (mtInformation), [mbOK], 0);
        DBGridEh1.Selection.Clear;
        edit1.SetFocus;
        exit;
      end;

      DBGridEh1.Selection.Clear;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          if ADOQuery1.Locate('inf_name', Aqr1.FieldByName('inf_name').AsString, []) then
            DBGridEh1.SelectedRows.CurrentRowSelected := True;
          DBGridEh1.SetFocus;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;




  if Label4.Caption = '环境影响类型管理' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from char_coef where LCI_factor like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关环境负荷！', (mtInformation), [mbOK], 0);
        DBGridEh1.Selection.Clear;
        edit1.SetFocus;
        exit;
      end;

      DBGridEh1.Selection.Clear;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          if ADOQuery1.Locate('LCI_factor', Aqr1.FieldByName('LCI_factor').AsString, []) then
            DBGridEh1.SelectedRows.CurrentRowSelected := True;
          DBGridEh1.SetFocus;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;


  if Label4.Caption = Label1.Caption + '3' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection2;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from char_coef where LCI_factor like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关环境负荷！', (mtInformation), [mbOK], 0);
        DBGridEh1.Selection.Clear;
        edit1.SetFocus;
        exit;
      end;

      DBGridEh1.Selection.Clear;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          if ADOQuery1.Locate('LCI_factor', Aqr1.FieldByName('LCI_factor').AsString, []) then
            DBGridEh1.SelectedRows.CurrentRowSelected := True;
          DBGridEh1.SetFocus;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;

  if Label4.Caption = '运输LCA基础数据管理' then
  begin
   { Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from transport_base where shuxing like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关环境指标！', (mtInformation), [mbOK], 0);
        DBGridEh1.Selection.Clear;
        edit1.SetFocus;

        exit;
      end;

      DBGridEh1.Selection.Clear;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          if ADOQuery1.Locate('shuxing', Aqr1.FieldByName('shuxing').AsString, []) then
            DBGridEh1.SelectedRows.CurrentRowSelected := True;
          DBGridEh1.SetFocus;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;   }
  end;


  if Label4.Caption = '物料能源属性管理' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from shuxing where name like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关产品！', (mtInformation), [mbOK], 0);
        edit1.SetFocus;
        exit;
      end;
      listbox4.ClearSelection;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          listbox4.Selected[listbox4.Items.IndexOf(Aqr1.FieldByName('name').AsString)] := true;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;

  if Label4.Caption = '工序对标数据管理' then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;
    try
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from historydata where source like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        form1.bsSkinMessage1.MessageDlg('未查询到相关产品！', (mtInformation), [mbOK], 0);
        edit1.SetFocus;
        exit;
      end;
      listbox4.ClearSelection;
      if Aqr1.RecordCount > 0 then
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          listbox4.Selected[listbox4.Items.IndexOf(Aqr1.FieldByName('source').AsString)] := true;
        end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;
  end;





end;

procedure TForm1.bsSkinButton22Click(Sender: TObject);
var
  i: integer;

begin
  ListBox4.Selected[ListBox4.ItemIndex] := true;
  if ListBox4.ItemIndex = -1 then
    exit;

  if Label4.Caption = '外部LCA数据管理' then
    if form1.bsSkinMessage1.MessageDlg('确定删除所选产品LCA数据？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      for i := ListBox4.Items.Count - 1 downto 0 do
        if ListBox4.Selected[i] = true then
        begin
          AdoQuery1.Connection := Form1.ADOConnection1;
          if AdoQuery1.Active then AdoQuery1.Close;
          AdoQuery1.SQL.Clear;
          AdoQuery1.SQL.Add('delete * from up where product=' + #39 + ListBox4.Items[i] + #39);
          AdoQuery1.ExecSQL;
          ListBox4.Items.Delete(i);
        end;


  if Label4.Caption = Label1.Caption + '1' then
    if form1.bsSkinMessage1.MessageDlg('确定删除所选产品LCA数据？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      for i := ListBox4.Items.Count - 1 downto 0 do
        if ListBox4.Selected[i] = true then
        begin
          AdoQuery1.Connection := Form1.ADOConnection2;
          if AdoQuery1.Active then AdoQuery1.Close;
          AdoQuery1.SQL.Clear;
          AdoQuery1.SQL.Add('delete * from up where product=' + #39 + ListBox4.Items[i] + #39);
          AdoQuery1.ExecSQL;
          ListBox4.Items.Delete(i);
        end;








  if Label4.Caption = '环境影响类型管理' then
    if form1.bsSkinMessage1.MessageDlg('确定删除所选环境影响？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      AdoQuery1.Connection := Form1.ADOConnection1;
      if AdoQuery1.Active then AdoQuery1.Close;
      AdoQuery1.SQL.Clear;
      AdoQuery1.SQL.Add('delete * from char_coef where impact_type=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      AdoQuery1.ExecSQL;
      ListBox4.Items.Delete(ListBox4.ItemIndex);
    end;

  if Label4.Caption = Label1.Caption + '3' then
    if form1.bsSkinMessage1.MessageDlg('确定删除所选环境影响？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      AdoQuery1.Connection := Form1.ADOConnection2;
      if AdoQuery1.Active then AdoQuery1.Close;
      AdoQuery1.SQL.Clear;
      AdoQuery1.SQL.Add('delete * from char_coef where impact_type=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      AdoQuery1.ExecSQL;
      ListBox4.Items.Delete(ListBox4.ItemIndex);
    end;

 // if Label4.Caption = '运输LCA基础数据管理' then
   { if form1.bsSkinMessage1.MessageDlg('确定删除所选运输方式？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if AdoQuery1.Active then AdoQuery1.Close;
      AdoQuery1.SQL.Clear;
      AdoQuery1.SQL.Add('delete * from transport_base where name=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      AdoQuery1.ExecSQL;
      ListBox4.Items.Delete(ListBox4.ItemIndex);
    end;  }

  if Label4.Caption = '物料能源属性管理' then
    if form1.bsSkinMessage1.MessageDlg('确定删除所选物料或能源？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      AdoQuery1.Connection := Form1.ADOConnection1;
      if AdoQuery1.Active then AdoQuery1.Close;
      AdoQuery1.SQL.Clear;
      AdoQuery1.SQL.Add('delete * from shuxing where name=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      AdoQuery1.ExecSQL;
      ListBox4.Items.Delete(ListBox4.ItemIndex);
    end;

  if Label4.Caption = '工序对标数据管理' then
  begin
    if form1.bsSkinMessage1.MessageDlg('确定删除所选数据源的全部数据？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      AdoQuery1.Connection := Form1.ADOConnection1;
      if AdoQuery1.Active then AdoQuery1.Close;
      AdoQuery1.SQL.Clear;
      AdoQuery1.SQL.Add('delete * from historydata where source=' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      AdoQuery1.ExecSQL;
      ListBox4.Items.Delete(ListBox4.ItemIndex);
      history_Format; //历史数据格式化
    end;

  end;

end;

procedure TForm1.bsSkinButton26Click(Sender: TObject);
var
  Aqr1: TAdoQuery;
  i: integer;
begin

  ADOQuery1.UpdateBatch(arall);

  Aqr1 := TAdoQuery.Create(nil);

  try

    if Label4.Caption = '外部LCA数据管理' then
    begin

      Aqr1.Connection := form1.ADOConnection1;
      if ComboBox2.ItemIndex = -1 then
      begin
        form1.bsSkinMessage1.MessageDlg('请先选择功能单位！', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('LCItype', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('类别不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('LCIfactor', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('名称不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;
      if ADOQuery1.Locate('unit', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('单位不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;


// ID,,LCItype,LCIfactor,unit,sum,function_unit,see_level

      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        ADOQuery1.Edit;
        if ADOQuery1.FieldByName('product').AsString = '' then
          ADOQuery1.FieldByName('product').AsString := ListBox4.Items[ListBox4.ItemIndex];
        if ADOQuery1.FieldByName('see_level').AsString = '' then
          ADOQuery1.FieldByName('see_level').AsString := '4x';

        ADOQuery1.FieldByName('function_unit').AsString := ComboBox2.Text;
        ADOQuery1.Post;
      end;

      ADOQuery1.UpdateBatch(arall);
    end;


    if Label4.Caption = Label1.Caption + '1' then
    begin
      Aqr1.Connection := form1.ADOConnection2;
      if ComboBox2.ItemIndex = -1 then
      begin
        form1.bsSkinMessage1.MessageDlg('请先选择功能单位！', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('LCItype', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('类别不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('LCIfactor', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('名称不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;
      if ADOQuery1.Locate('unit', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('单位不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;


// ID,,LCItype,LCIfactor,unit,sum,function_unit,see_level

      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        ADOQuery1.Edit;
        if ADOQuery1.FieldByName('product').AsString = '' then
          ADOQuery1.FieldByName('product').AsString := ListBox4.Items[ListBox4.ItemIndex];
        if ADOQuery1.FieldByName('see_level').AsString = '' then
          ADOQuery1.FieldByName('see_level').AsString := '4x';

        ADOQuery1.FieldByName('function_unit').AsString := ComboBox2.Text;
        ADOQuery1.Post;
      end;

      ADOQuery1.UpdateBatch(arall);
    end;




    if Label4.Caption = 'LCI环境指标管理' then
    begin

      if ADOQuery1.Locate('inf_type', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('类别不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('inf_name', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('名称不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;
      if ADOQuery1.Locate('units', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('单位不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      ADOQuery1.UpdateBatch(arall);
    end;

    if Label4.Caption = Label1.Caption + '4' then
    begin

      if ADOQuery1.Locate('inf_type', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('类别不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('inf_name', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('名称不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;
      if ADOQuery1.Locate('units', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('单位不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      ADOQuery1.UpdateBatch(arall);
    end;




    if Label4.Caption = '环境影响类型管理' then
    begin

      Aqr1.Connection := form1.ADOConnection1;
      if ADOQuery1.Locate('impact_type', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('环境影响类型不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('LCI_factor', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('LCI环境指标不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('eq_unit', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('当量单位不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;
      if ADOQuery1.Locate('equivalent', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('数量不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select * from  LCIfactor ');
      Aqr1.Open;

      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        if Aqr1.Locate('inf_name', ADOQuery1.FieldByName('LCI_factor').AsString, []) = false then
        begin
          form1.bsSkinMessage1.MessageDlg(ADOQuery1.FieldByName('LCI_factor').AsString + ' 指标不存在，请确认或从"LCI环境指标管理"界面添加！', (mtInformation), [mbOK], 0);
          exit;
        end;
      end;


      ADOQuery1.UpdateBatch(arall);

      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select distinct  LCI_factor  from char_coef where  impact_type =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      Aqr1.Open;
      if Aqr1.RecordCount < ADOQuery1.RecordCount then
      begin
        form1.bsSkinMessage1.MessageDlg('存在重复记录，请修改！', (mtInformation), [mbOK], 0);
        exit;
      end;

      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select distinct eq_unit  from char_coef where  impact_type =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      Aqr1.Open;
      if Aqr1.RecordCount > 1 then
      begin
        form1.bsSkinMessage1.MessageDlg('当量单位不一致，请修改！', (mtInformation), [mbOK], 0);
        exit;
      end;




    end;


    if Label4.Caption = Label1.Caption + '3' then
    begin

      Aqr1.Connection := form1.ADOConnection2;
      if ADOQuery1.Locate('impact_type', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('环境影响类型不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('LCI_factor', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('LCI环境指标不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;

      if ADOQuery1.Locate('eq_unit', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('当量单位不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;
      if ADOQuery1.Locate('equivalent', null, []) then
      begin
        form1.bsSkinMessage1.MessageDlg('数量不能为空', (mtInformation), [mbOK], 0);
        exit;
      end;


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select * from  LCIfactor ');
      Aqr1.Open;

      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        if Aqr1.Locate('inf_name', ADOQuery1.FieldByName('LCI_factor').AsString, []) = false then
        begin
          form1.bsSkinMessage1.MessageDlg(ADOQuery1.FieldByName('LCI_factor').AsString + ' 指标不存在，请确认或从"LCI环境指标管理"界面添加！', (mtInformation), [mbOK], 0);
          exit;
        end;
      end;


      ADOQuery1.UpdateBatch(arall);

      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select distinct  LCI_factor  from char_coef where  impact_type =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      Aqr1.Open;
      if Aqr1.RecordCount < ADOQuery1.RecordCount then
      begin
        form1.bsSkinMessage1.MessageDlg('存在重复记录，请修改！', (mtInformation), [mbOK], 0);
        exit;
      end;

      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select distinct eq_unit  from char_coef where  impact_type =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      Aqr1.Open;
      if Aqr1.RecordCount > 1 then
      begin
        form1.bsSkinMessage1.MessageDlg('当量单位不一致，请修改！', (mtInformation), [mbOK], 0);
        exit;
      end;
    end;


    if Label4.Caption = '运输LCA基础数据管理' then
    begin

      for i := 1 to ADOQuery1.FieldCount - 1 do
        if ADOQuery1.Locate(ADOQuery1.Fields[i].FieldName, null, []) then
        begin
          form1.bsSkinMessage1.MessageDlg(dbgrideh1.Columns[i].Title.Caption + '不能为空', (mtInformation), [mbOK], 0);
          exit;
        end;
      Aqr1.Connection := form1.ADOConnection1;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select * from  LCIfactor ');
      Aqr1.Open;

      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        if Aqr1.Locate('inf_name', ADOQuery1.FieldByName('shuxing').AsString, []) = false then
        begin
          form1.bsSkinMessage1.MessageDlg('LCI环境指标不存在，请确认或从"LCI环境指标管理"界面添加！', (mtInformation), [mbOK], 0);
          exit;
        end;
      end;

      ADOQuery1.UpdateBatch(arall);


    {  Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select distinct shuxing  from transport_base where  name =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      Aqr1.Open;
      if Aqr1.RecordCount < ADOQuery1.RecordCount then
      begin
        form1.bsSkinMessage1.MessageDlg('存在重复记录，请修改！', (mtInformation), [mbOK], 0);
        exit;
      end;  }

    end;

    if Label4.Caption = Label1.Caption + '2' then
    begin

      for i := 1 to ADOQuery1.FieldCount - 1 do
        if ADOQuery1.Locate(ADOQuery1.Fields[i].FieldName, null, []) then
        begin
          form1.bsSkinMessage1.MessageDlg(dbgrideh1.Columns[i].Title.Caption + '不能为空', (mtInformation), [mbOK], 0);
          exit;
        end;
      Aqr1.Connection := form1.ADOConnection2;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select * from  LCIfactor ');
      Aqr1.Open;

      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        if Aqr1.Locate('inf_name', ADOQuery1.FieldByName('shuxing').AsString, []) = false then
        begin
          form1.bsSkinMessage1.MessageDlg('LCI环境指标不存在，请确认或从"LCI环境指标管理"界面添加！', (mtInformation), [mbOK], 0);
          exit;
        end;
      end;

      ADOQuery1.UpdateBatch(arall);


   {   Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select distinct shuxing  from transport_base where  name =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      Aqr1.Open;
      if Aqr1.RecordCount < ADOQuery1.RecordCount then
      begin
        form1.bsSkinMessage1.MessageDlg('存在重复记录，请修改！', (mtInformation), [mbOK], 0);
        exit;
      end;   }
    end;



    if Label4.Caption = '物料能源属性管理' then
    begin
      for i := 1 to ADOQuery1.FieldCount - 1 do
        if ADOQuery1.Locate(ADOQuery1.Fields[i].FieldName, null, []) then
        begin
          form1.bsSkinMessage1.MessageDlg(dbgrideh1.Columns[i].Title.Caption + '不能为空', (mtInformation), [mbOK], 0);
          exit;
        end;

      ADOQuery1.UpdateBatch(arall);
      Aqr1.Connection := form1.ADOConnection1;
      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('Select distinct shuxing  from shuxing where  name =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
      Aqr1.Open;
      if Aqr1.RecordCount < ADOQuery1.RecordCount then
      begin
        form1.bsSkinMessage1.MessageDlg('存在重复记录，请修改！', (mtInformation), [mbOK], 0);
        exit;
      end;

    end;

    if Label4.Caption = '工序对标数据管理' then
    begin

      for i := 1 to ADOQuery1.FieldCount - 2 do
        if ADOQuery1.Locate(ADOQuery1.Fields[i].FieldName, null, []) then
        begin
          form1.bsSkinMessage1.MessageDlg(dbgrideh1.Columns[i].Title.Caption + '不能为空', (mtInformation), [mbOK], 0);
          exit;
        end;

      ADOQuery1.UpdateBatch(arall);

      Aqr1.Connection := form1.ADOConnection1;


      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('Select * from historydata where  source =' + #39 + ListBox4.Items[ListBox4.ItemIndex] + #39);
        Aqr1.SQL.Add(' and process=' + #39 + ADOQuery1.FieldByName('process').AsString + #39);
        Aqr1.SQL.Add(' and inf=' + #39 + ADOQuery1.FieldByName('inf').AsString + #39);
        Aqr1.Open;

        if Aqr1.RecordCount > 1 then
        begin
          form1.bsSkinMessage1.MessageDlg('存在重复记录，请修改！', (mtInformation), [mbOK], 0);
          exit;
        end;
      end;


    end;

  finally
    Aqr1.Close;
    Aqr1.Free;
  end;




end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  if ADOQuery1.Active = false then exit;
  if CheckBox1.Checked = false then
  begin
    ADOQuery1.Filtered := false;
  end;


  if CheckBox1.Checked = true then
  begin
    ADOQuery1.Filtered := false;
    ADOQuery1.Filter := 'see_level like ' + #39 + '0%' + #39;
    ADOQuery1.Filtered := true;
    if Label4.Caption = '外部LCA数据管理' then
      ComboBox2.ItemIndex := ComboBox2.Items.IndexOf(ADOQuery1.FieldByName('function_unit').AsString);
  end;

end;

procedure TForm1.bsSkinButton3Click(Sender: TObject);
begin
  if PageControl1.Visible = false then exit;
  PageControl1.ActivePageIndex := 10;
  PageControl1Change(Sender);

end;


function TForm1.chechedcount(bsSkinCheckListBox1: TbsSkinCheckListBox): integer;
//返回选中个数
var
  i, j: integer;
begin
  j := 0;
  if bsSkinCheckListBox1.Items.Count = 0 then
  begin
    result := 0;
    exit;
  end
  else
    for i := 0 to bsSkinCheckListBox1.Items.Count - 1 do
    begin
      if bsSkinCheckListBox1.Checked[i] = true then
        j := j + 1;
    end;
  result := j;
end;



{procedure TForm1.bsSkinCheckListBox6ClickCheck(Sender: TObject);
var
  i, j: integer;
  s: string;
  Aqr1: TAdoQuery;
begin
 // bsSkinCheckListBox7.Items.Clear;
  if chechedcount(bsSkinCheckListBox6) = 0 then
  begin
  //  bsSkinMessage1.MessageDlg('请选择项目！', (mtInformation), [mbOK], 0);
    exit;
  end;


  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  try
    for i := 0 to bsSkinCheckListBox6.Items.Count - 1 do
      if bsSkinCheckListBox6.Checked[i] then
      begin
        s := 'select distinct process  from data';
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add(s);
        Aqr1.Open;

        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          if bsSkinCheckListBox7.Items.IndexOf(Aqr1.FieldByName('process').AsString) = -1 then
            bsSkinCheckListBox7.Items.Add(Aqr1.FieldByName('process').AsString);
        end;
      end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;


end;       }

{procedure TForm1.CheckBox2Click(Sender: TObject);
var
  i: integer;
begin
  if CheckBox2.Checked then
  begin
    for i := 0 to bsSkinCheckListBox6.Items.Count - 1 do
      bsSkinCheckListBox6.Checked[i] := true;

  end;

  if CheckBox2.Checked = false then
  begin
    for i := 0 to bsSkinCheckListBox6.Items.Count - 1 do
      bsSkinCheckListBox6.Checked[i] := false;

  end;


  bsSkinCheckListBox6ClickCheck(Sender);
end;      }

{procedure TForm1.CheckBox3Click(Sender: TObject);
var
  i: integer;
begin
  if CheckBox3.Checked then
  begin
    for i := 0 to bsSkinCheckListBox7.Items.Count - 1 do
      bsSkinCheckListBox7.Checked[i] := true;

  end;

  if CheckBox3.Checked = false then
  begin
    for i := 0 to bsSkinCheckListBox7.Items.Count - 1 do
      bsSkinCheckListBox7.Checked[i] := false;

  end;


  bsSkinCheckListBox7ClickCheck(Sender);

end;   }

{procedure TForm1.bsSkinCheckListBox7ClickCheck(Sender: TObject);
var
  i, j: integer;
  mysql1: string;
 // Aqr1: TAdoQuery;
begin

  if chechedcount(bsSkinCheckListBox6) = 0 then
  begin
    ADOQuery2.Close;
    exit;
  end;

  if chechedcount(bsSkinCheckListBox7) = 0 then
  begin
    ADOQuery2.Close;
    exit;
  end;



  mysql1 := 'select * from ';
  for i := 0 to bsSkinCheckListBox6.Items.Count - 1 do
    if bsSkinCheckListBox6.Checked[i] then
    begin
      mysql1 := mysql1 + bsSkinCheckListBox6.Items.Strings[i];
      break;
    end;

  if chechedcount(bsSkinCheckListBox6) > 1 then
  begin
    for j := i + 1 to bsSkinCheckListBox6.Items.Count - 1 do
      if bsSkinCheckListBox6.Checked[j] then
        mysql1 := mysql1 + ' union (select * from ' + bsSkinCheckListBox6.Items.Strings[j] + ')';
  end;


  if chechedcount(bsSkinCheckListBox7) > 0 then
  begin
    for i := 0 to bsSkinCheckListBox7.Items.Count - 1 do
      if bsSkinCheckListBox7.Checked[i] then
      begin
        mysql1 := 'select * from (' + mysql1 + ') where (process = ' + #39 + bsSkinCheckListBox7.Items.Strings[i] + #39 + ')';
        break;
      end;

    if chechedcount(bsSkinCheckListBox7) > 1 then
    begin
      for j := i + 1 to bsSkinCheckListBox7.Items.Count - 1 do
        if bsSkinCheckListBox7.Checked[j] then
        begin
          mysql1 := mysql1 + ' or (process =  ' + #39 + bsSkinCheckListBox7.Items.Strings[j] + #39 + ')';
        end;

    end;

  end;

      // ***********************************

 //showmessage(mysql1);
  if ADOQuery2.Active then ADOQuery2.Close;
  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add(mysql1);
  ADOQuery2.Open;

  DataSource1.DataSet := ADOQuery2;

  for i := 0 to DBGridEh1.Columns.Count - 1 do
  begin
    DBGridEh1.Columns[i].Width := 80;
    DBGridEh1.Columns[i].Title.Alignment := taCenter;
    DBGridEh1.Columns[i].Alignment := taLeftJustify;
  end;

  ComboBox1.Items.Clear;

  for i := 1 to ADOQuery2.RecordCount do
  begin
    ADOQuery2.RecNo := i;
    if ComboBox1.Items.IndexOf(ADOQuery2.FieldByName('pd_name').AsString) = -1 then
      ComboBox1.Items.Add(ADOQuery2.FieldByName('pd_name').AsString);
  end;

end;      }

procedure TForm1.bsSkinButton29Click(Sender: TObject);
begin
  if PageControl1.Visible = false then exit;

  SaveDialog1.FileName := '';
  if PageControl1.TabIndex = 2 then
  begin
    if DBGrideh9.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh9);
    DBGridEh9.DataSource.DataSet.First;
  end;

  if PageControl1.TabIndex = 3 then
  begin
    if DBGrideh5.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh5);
    DBGridEh5.DataSource.DataSet.First;
  end;

  if PageControl1.TabIndex = 4 then
  begin
    if DBGrideh4.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh4);
    DBGridEh4.DataSource.DataSet.First;
  end;

  if PageControl1.TabIndex = 5 then
  begin
    if DBGrideh2.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh2);
    DBGridEh2.DataSource.DataSet.First;
  end;

  if PageControl1.TabIndex = 6 then
  begin
    if DBGrideh6.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh6);
    DBGridEh6.DataSource.DataSet.First;
  end;


  if PageControl1.TabIndex = 7 then
  begin
    if DBGrideh8.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh8);
    DBGridEh8.DataSource.DataSet.First;
  end;

  if PageControl1.TabIndex = 8 then
  begin
    if DBGrideh11.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh11);
    DBGridEh11.DataSource.DataSet.First;
  end;

  if PageControl1.TabIndex = 10 then
  begin
  {  if DBGrideh12.DataSource.DataSet.Active = false then exit;
    DBgridToExcel(DBGrideh12);
    DBGridEh12.DataSource.DataSet.First; }
  end;

end;

procedure TForm1.TreeView1Click(Sender: TObject);
var
  Aqr1: TAdoQuery;
  mysql: string;
  i: integer;
begin
  if TreeView1.Hint = '项目管理' then
  begin
    if TreeView1.Selected.Level = 0 then
      Label1.Caption := TreeView1.Selected.Text; //项目名称
    if TreeView1.Selected.Level = 1 then
      Label1.Caption := TreeView1.Selected.Parent.Text;
    if TreeView1.Selected.Level = 2 then
      Label1.Caption := TreeView1.Selected.Parent.Parent.Text;
    StatusBar1.Panels[1].Text := '当前项目： ' + Label1.Caption;
    PageControl1.Visible := true;
    PageControl1.BringToFront;
    for i := 0 to 10 do
      PageControl1.Pages[i].TabVisible := true;

    PageControl1.Pages[11].TabVisible := false;
    PageControl1.Pages[12].TabVisible := false;
  //  ComboBox8.ItemIndex:=-1;
 //   ComboBox8.Text := '筛选工序（按拼音排序）';

  //    ADOQuery5.Filtered := false;
   //   ADOQuery6.Filtered := false;
    if TreeView1.Selected.Level = 0 then
    begin
      PageControl1.ActivePageIndex := 0;
      PageControl1.Pages[0].Highlighted := true;
      for i := 1 to PageControl1.PageCount - 1 do
        PageControl1.Pages[i].Highlighted := false;

      Label3.Caption := '< ' + TreeView1.Selected.Text + ' >  项目基本信息';

      ADOConnection2.Close;
      ADOConnection2.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.mdb ; Persist Security Info=False';
      ADOConnection2.Open;

      Aqr1 := TAdoQuery.Create(nil);
      Aqr1.Connection := form1.ADOConnection2;
      try
        mysql := 'select * from project_inf';
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add(mysql);
        Aqr1.Open;

        Aqr1.Locate('project_item', 'LCA类别', [loCaseInsensitive]);
        FlatComboBox1.ItemIndex := FlatComboBox1.Items.IndexOf(Aqr1.FieldByName('item_inf').AsString);


        Aqr1.Locate('project_item', '项目建立时间', [loCaseInsensitive]);
        FlatEdit3.Text := Aqr1.FieldByName('item_inf').AsString;

        Aqr1.Locate('project_item', '项目所属系统', [loCaseInsensitive]);
        FlatEdit23.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', 'LCA联系人信息', [loCaseInsensitive]);
        FlatEdit4.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '产品名称（最后工序）', [loCaseInsensitive]);
        FlatEdit6.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '产品系统描述', [loCaseInsensitive]);
        FlatMemo6.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '功能单位数据', [loCaseInsensitive]);
        FlatEdit8.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '功能单位', [loCaseInsensitive]);
        bsSkinComboBox1.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '生产数据时间', [loCaseInsensitive]);
        FlatEdit9.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '数据取舍原则', [loCaseInsensitive]);
        FlatMemo10.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '运输', [loCaseInsensitive]);
        FlatEdit11.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '上游数据', [loCaseInsensitive]);
        FlatEdit14.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '数据质量评价', [loCaseInsensitive]);
        FlatEdit17.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '数据分配', [loCaseInsensitive]);
        FlatMemo13.Text := Aqr1.FieldByName('item_inf').AsString;



        Aqr1.Locate('project_item', '循环环境收益', [loCaseInsensitive]);
        FlatMemo12.Text := Aqr1.FieldByName('item_inf').AsString;



      finally
        Aqr1.Close;
        Aqr1.Free;
      end;





    end;

    ADOConnection2.Close;
    ADOConnection2.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.mdb ; Persist Security Info=False';
    ADOConnection2.Open;



  end; //if 项目管理

//**********************************************************************************


  if TreeView1.Hint = '基础信息管理' then
  begin
    StatusBar1.Panels[1].Text := '基础信息管理： ' + TreeView1.Selected.Text;
    Label1.Caption := 'base';

    if TreeView1.Selected.Text = '外部LCA数据管理' then
      bsSkinButton4Click(Sender); //上游数据管理

    if TreeView1.Selected.Text = '运输LCA基础数据管理' then
      bsSkinButton5Click(Sender);

    if TreeView1.Selected.Text = 'LCI环境指标管理' then
      bsSkinButton11Click(Sender);

    if TreeView1.Selected.Text = '环境影响类型管理' then
      bsSkinButton12Click(Sender);

    if TreeView1.Items[6].Selected = true then
    begin
      RadioButton5.Checked := true;
      bsSkinButton46Click(Sender);
    end;
    if TreeView1.Items[7].Selected = true then
    begin
      RadioButton6.Checked := true;
      bsSkinButton46Click(Sender);
    end;
    if TreeView1.Items[8].Selected = true then
    begin
      RadioButton7.Checked := true;
      bsSkinButton46Click(Sender);
    end;
    if TreeView1.Items[9].Selected = true then
    begin
      RadioButton8.Checked := true;
      bsSkinButton46Click(Sender);
    end;

    if TreeView1.Selected.Text = '运输LCA基础数据管理' then
      bsSkinButton5Click(Sender);

    if TreeView1.Selected.Text = '工序对标数据管理' then
      bsSkinButton77Click(Sender);

  end;

end;

procedure TForm1.bsSkinButton6Click(Sender: TObject);
var
  i: integer;
begin
  dxFlowChart1.Zoom := 0;
  form2.TreeView1.Items.Clear;
  form2.TreeView1.LoadFromFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\tree.txt');
  form2.dxFlowChart1.LoadFromFile(ExtractFilePath(Application.ExeName) + 'projects\' + form1.label1.Caption + '\' + label1.Caption + '.flc');

  form2.ListBox1.Clear;
  for i := 0 to form2.dxFlowChart1.ObjectCount - 1 do
    if form2.listbox1.Items.IndexOf(treeview1.Selected.Text) = -1 then
      form2.listbox1.Items.Add(form2.dxFlowChart1.Objects[i].Text);

  form2.dxFlowChart1.Zoom := 100;
  form2.TreeView1.FullExpand;

  form2.dxFlowChart1.Align := alclient;
  form2.ShowModal;
end;

procedure TForm1.dxFlowChart1DblClick(Sender: TObject);
 {var
  i: integer; }
begin
 { if dxFlowChart1.SelectedConnectionCount > 0 then
    for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
      dxFlowChart1.SelectedConnections[i].Selected := false;

  if dxFlowChart1.SelectedObjectCount = 0 then
    exit;

  listbox2.Items.Clear;
  for i := 0 to dxFlowChart1.ObjectCount - 1 do
    if dxFlowChart1.Objects[i].Text <> '主生产工序' then
      if trim(dxFlowChart1.Objects[i].Text) <> '能源公辅工序' then
        if dxFlowChart1.Objects[i].Text <> '至各工序' then
          listbox2.Items.Add(dxFlowChart1.Objects[i].Text);

  ListBox2.ItemIndex := listbox2.Items.IndexOf(dxFlowChart1.SelectedObject.Text);
  edit1.Text := dxFlowChart1.SelectedObject.Text;
 
  ProcessData_Format;
//  PageControl1.ActivePageIndex := 1;
try
  bsSkinButton7Click(Sender );
  except
  end;  }
end;

procedure TForm1.bsSkinButton9Click(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 0;
  PageControl1Change(Sender);
end;

procedure TForm1.bsSkinButton7Click(Sender: TObject);
//var
  //i: integer;
begin
  if dxFlowChart1.SelectedObjectCount <> 0 then
    edit2.Text := dxFlowChart1.SelectedObject.Text;

{  if dxFlowChart1.ObjectCount = 0 then exit;
  listbox2.Items.Clear;
  listbox2.Items.Add('全部工序');
  for i := 0 to dxFlowChart1.ObjectCount - 1 do
    if dxFlowChart1.Objects[i].Text <> '主生产工序' then
      if trim(dxFlowChart1.Objects[i].Text) <> '能源公辅工序' then
        if dxFlowChart1.Objects[i].Text <> '至各工序' then
          listbox2.Items.Add(dxFlowChart1.Objects[i].Text);


 //删除冗余数据  例如载入流程图后，原先残留的数据

  AdoQuery1.Connection := Form1.ADOConnection2;
  if AdoQuery1.Active then AdoQuery1.Close;
  AdoQuery1.SQL.Clear;
  AdoQuery1.SQL.Add('select * from data ');
  ADOQuery1.Open;
  for i := 1 to ADOQuery1.RecordCount do
  begin
    ADOQuery1.RecNo := i;
    if ListBox2.Items.IndexOf(ADOQuery1.FieldByName('process').AsString) = -1 then
      ADOQuery1.Delete;
  end;


  if dxFlowChart1.SelectedObjectCount <> 0 then
  begin
    ListBox2.ItemIndex := listbox2.Items.IndexOf(dxFlowChart1.SelectedObject.Text);
    ListBox2.Selected[ListBox2.ItemIndex] := true;
    edit2.Text := dxFlowChart1.SelectedObject.Text;
  end
  else
  begin
    ListBox2.ItemIndex := 1;
    ListBox2.Selected[1] := true;
    edit2.Text := '';
  end;
  ProcessData_Format;
  DBGridEh9.Columns.Items[3].PickList.clear;
  if ListBox2.ItemIndex <> -1 then
    DBGridEh9.Columns.Items[3].picklist.Add(ListBox2.Items[ListBox2.ItemIndex]);
               }
  PageControl1.ActivePageIndex := 2;
  PageControl1Change(Sender);
end;


procedure TForm1.ProcessData_Format; //工序数据格式化
var
  i: integer;
  mysql1: string;
begin
  AdoQuery1.Connection := Form1.ADOConnection2;
  if ListBox2.ItemIndex = -1 then ListBox2.ItemIndex := 1;

  mysql1 := 'select ID,system_name,	processtype,process,material_type,	pd_name,	pd_unit,pd, remark from data  where process=' + #39 + ListBox2.Items[ListBox2.ItemIndex] + #39;
  if ListBox2.ItemIndex = 0 then
    mysql1 := 'select ID,system_name,	processtype,process,material_type,	pd_name,	pd_unit,pd, remark from data where process is not null';

  if AdoQuery1.Active then AdoQuery1.Close;
  AdoQuery1.SQL.Clear;
  AdoQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;
  DBGridEh9.Columns[1].Visible := false;

  if ADOQuery1.RecordCount > 0 then
    ADOQuery1.Sort := 'processtype ASC,PROCESS ASC, material_type ASC';
  if ListBox2.ItemIndex = 0 then
    ADOQuery1.Sort := 'processtype ASC,PROCESS ASC, material_type ASC';



  for i := 0 to DBGridEh9.Columns.Count - 1 do
  begin
    DBGridEh9.Columns[i].Width := 80;
    DBGridEh9.Columns[i].Title.Alignment := taCenter;
    DBGridEh9.Columns[i].Alignment := taCenter;
  end;

  DBGridEh9.Columns[7].Alignment := taRightJustify; //数据
  DBGridEh9.Columns[3].Alignment := taLeftJustify;
  DBGridEh9.Columns[4].Alignment := taLeftJustify;

  DBGridEh9.Columns[0].Title.caption := '序号';
  DBGridEh9.Columns[1].Title.caption := '系统名';
  DBGridEh9.Columns[2].Title.caption := '工序类别';
  DBGridEh9.Columns[3].Title.caption := '工序名称';
  DBGridEh9.Columns[4].Title.caption := '数据类别';
  DBGridEh9.Columns[5].Title.caption := '数据名称';
  DBGridEh9.Columns[6].Title.caption := '单位';
  DBGridEh9.Columns[7].Title.caption := '数量';
  DBGridEh9.Columns[8].Title.caption := '备注';


  DBGridEh9.Columns[0].Width := 40;
  DBGridEh9.Columns[3].Width := 100;
  DBGridEh9.Columns[4].Width := 120;
  DBGridEh9.Columns[5].Width := 120;
  DBGridEh9.Columns[7].Width := 120;
  DBGridEh9.Columns[8].Width := 160;


  DBGridEh9.Columns.Items[2].PickList.clear;
  DBGridEh9.Columns.Items[2].picklist.Add('主生产工序');
  DBGridEh9.Columns.Items[2].picklist.Add('能源公辅工序');


  DBGridEh9.Columns.Items[4].PickList.clear;
  DBGridEh9.Columns.Items[4].picklist.Add('产品');
  DBGridEh9.Columns.Items[4].picklist.Add('副产品或固体废弃物');
  DBGridEh9.Columns.Items[4].picklist.Add('能源及能源介质');
  DBGridEh9.Columns.Items[4].picklist.Add('原材料');
  DBGridEh9.Columns.Items[4].picklist.Add('辅助材料');
  DBGridEh9.Columns.Items[4].picklist.Add('大气排放');
  DBGridEh9.Columns.Items[4].picklist.Add('水体排放');

  DBGridEh9.Columns.Items[6].PickList.clear;
  DBGridEh9.Columns.Items[6].picklist.Add('千克(kg)');
  DBGridEh9.Columns.Items[6].picklist.Add('吨(ton)');
  DBGridEh9.Columns.Items[6].picklist.Add('万吨(10,000tons)');
  DBGridEh9.Columns.Items[6].picklist.Add('千立方米(km3)');
  DBGridEh9.Columns.Items[6].picklist.Add('立方米(m3)');
  DBGridEh9.Columns.Items[6].picklist.Add('千瓦时(kWh)');
  DBGridEh9.Columns.Items[6].picklist.Add('万千瓦时(10,000kWh)');
  DBGridEh9.Columns.Items[6].picklist.Add('吉焦(GJ)');
  DBGridEh9.Columns.Items[6].picklist.Add('兆焦(MJ)');
  DBGridEh9.Columns.Items[6].picklist.Add('千焦(KJ)');
  DBGridEh9.Columns.Items[6].picklist.Add('个(件、罐、台、根...)');


end;


procedure TForm1.ListBox2Click(Sender: TObject);
var
  i: integer;
begin
  DBGridEh9.DataSource.DataSet.Filtered := false;
  if ADOQuery1.RecordCount > 0 then
    ADOQuery1.UpdateBatch(arall);



  ProcessData_Format;

  DBGridEh9.Columns.Items[3].PickList.clear;
  if ListBox2.ItemIndex > 0 then
    DBGridEh9.Columns.Items[3].picklist.Add(ListBox2.Items[ListBox2.ItemIndex]);
  if ListBox2.ItemIndex = 0 then
    for i := 1 to ListBox2.Items.Count - 1 do
      DBGridEh9.Columns.Items[3].picklist.Add(ListBox2.Items[i]);


 // if RadioButton2.Checked then
 //   TransportData_Format;
end;

procedure TForm1.TransportData_Format; //运输表格格式化
var
  i: integer;
  mysql1: string;
begin
  AdoQuery1.Connection := Form1.ADOConnection2;
  mysql1 := 'select ID, product,  sea_per,	train_per,	trunk_per,	river_per from transport ';


  if AdoQuery1.Active then AdoQuery1.Close;
  AdoQuery1.SQL.Clear;
  AdoQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;
  ADOQuery1.Active := true;

  for i := 0 to DBGridEh2.Columns.Count - 1 do
  begin
    DBGridEh2.Columns[i].Width := 120;
    DBGridEh2.Columns[i].Title.Alignment := taCenter;
    DBGridEh2.Columns[i].Alignment := taCenter;
  end;
  DBGridEh2.Columns[1].Alignment := taLeftJustify;
  DBGridEh2.Columns[0].Width := 40;
  DBGridEh2.Columns[0].Title.caption := '序号';
  DBGridEh2.Columns[1].Title.caption := '运输产品';
  DBGridEh2.Columns[2].Title.caption := '海运距离(km)';
  DBGridEh2.Columns[3].Title.caption := '火车距离(km)';
  DBGridEh2.Columns[4].Title.caption := '卡车距离(km)';
  DBGridEh2.Columns[5].Title.caption := '河运距离(km)';
end;





procedure TForm1.bsSkinButton17Click(Sender: TObject);
var
  i: integer;
begin
  if ListBox2.Count = 0 then exit;

  if trim(Edit2.Text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请输入查询名称中至少一个字符！', (mtInformation), [mbOK], 0);
    edit2.SetFocus;
    exit;
  end;

  listbox2.ClearSelection;
  for i := 0 to ListBox2.Count - 1 do
    if Pos(Edit2.Text, ListBox2.Items.Strings[i]) > 0 then
    begin
      ListBox2.Selected[i] := true;

    end;


  // showmessage('aaa');
end;

procedure TForm1.bsSkinButton35Click(Sender: TObject); // 从数据源批量载入数据
var
//  i: integer;
//  Aqr: TAdoQuery;
  P: TPoint;
begin

  GetCursorPos(P);
  PopupMenu3.Popup(P.X, P.Y);
{


  }

end;

procedure TForm1.bsSkinButton43Click(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 1;
  PageControl1Change(Sender);
end;



function TForm1.co_unit(process, s: string): string; //返回系数单位
//s：物料单位
var
  mysql, s1, s2: string;
  Aqr: TAdoQuery;
begin
  mysql := 'select * from data  where process = ' + #39 + process + #39 + ' and material_type=' + #39 + '产品' + #39;
  Aqr := TAdoQuery.Create(nil);
  Aqr.Connection := form1.ADOConnection2;
  Aqr.Close;
  Aqr.SQL.Clear;
  Aqr.SQL.Add(mysql);
  Aqr.Open;

  if s = '千克(kg)' then
    s2 := 'kg';
  if s = '吨(ton)' then
    s2 := 'kg';
  if s = '万吨(10,000tons)' then
    s2 := 'kg';
  if s = '千立方米(km3)' then
    s2 := 'm3';
  if s = '立方米(m3)' then
    s2 := 'm3';
  if s = '兆立方米(Mm3)' then
    s2 := 'm3';
  if s = '万千瓦时(10,000kWh)' then
    s2 := 'kWh';
  if s = '千瓦时(kWh)' then
    s2 := 'kWh';
  if s = '吉焦(GJ)' then
    s2 := '兆焦(MJ)';
  if s = '兆焦(MJ)' then
    s2 := '兆焦(MJ)';
  if s = '千焦(KJ)' then
    s2 := '兆焦(MJ)';
  if s = '个(件、罐、台...)' then
    s2 := '个(件、罐、台...)';


  if Aqr.RecordCount > 0 then
  begin
    Aqr.First;
    if Aqr.FieldByName('pd_unit').AsString = '千克(kg)' then
      s1 := 'kg';
    if Aqr.FieldByName('pd_unit').AsString = '吨(ton)' then
      s1 := 'kg';
    if Aqr.FieldByName('pd_unit').AsString = '万吨(10,000tons)' then
      s1 := 'kg';
    if Aqr.FieldByName('pd_unit').AsString = '立方米(m3)' then
      s1 := 'm3';
    if Aqr.FieldByName('pd_unit').AsString = '千立方米(km3)' then
      s1 := 'm3';
    if Aqr.FieldByName('pd_unit').AsString = '兆立方米(Mm3)' then
      s1 := 'm3';
    if Aqr.FieldByName('pd_unit').AsString = '万千瓦时(10,000kWh)' then
      s1 := 'kWh';
    if Aqr.FieldByName('pd_unit').AsString = '千瓦时(kWh)' then
      s1 := 'kWh';
    if Aqr.FieldByName('pd_unit').AsString = '吉焦(GJ)' then
      s1 := '兆焦(MJ)';
    if Aqr.FieldByName('pd_unit').AsString = '兆焦(MJ)' then
      s1 := '兆焦(MJ)';
    if Aqr.FieldByName('pd_unit').AsString = '千焦(KJ)' then
      s1 := '兆焦(MJ)';
    if Aqr.FieldByName('pd_unit').AsString = '个(件、罐、台...)' then
      s1 := '个(件、罐、台...)';

    result := s2 + '/' + s1;
  end
  else
    result := s2 + '/' + s2;
  Aqr.Close;
  Aqr.Free;
end;

function TForm1.co_unit1(process, s: string): string; //返回系数单位
//s：物料单位
var
  mysql, s1, s2: string;
  Aqr: TAdoQuery;
begin
  mysql := 'select * from  data_allocation where process = ' + #39 + process + #39 + ' and material_type=' + #39 + '产品' + #39;
  Aqr := TAdoQuery.Create(nil);
  Aqr.Connection := form1.ADOConnection2;
  Aqr.Close;
  Aqr.SQL.Clear;
  Aqr.SQL.Add(mysql);
  Aqr.Open;

  if s = '千克(kg)' then
    s2 := 'kg';
  if s = '吨(ton)' then
    s2 := 'kg';
  if s = '万吨(10,000tons)' then
    s2 := 'kg';
  if s = '千立方米(km3)' then
    s2 := 'm3';
  if s = '立方米(m3)' then
    s2 := 'm3';
  if s = '兆立方米(Mm3)' then
    s2 := 'm3';
  if s = '万千瓦时(10,000kWh)' then
    s2 := 'kWh';
  if s = '千瓦时(kWh)' then
    s2 := 'kWh';
  if s = '吉焦(GJ)' then
    s2 := '兆焦(MJ)';
  if s = '兆焦(MJ)' then
    s2 := '兆焦(MJ)';
  if s = '千焦(KJ)' then
    s2 := '兆焦(MJ)';
  if s = '个(件、罐、台...)' then
    s2 := '个(件、罐、台...)';


  if Aqr.RecordCount > 0 then
  begin
    Aqr.First;
    if Aqr.FieldByName('pd_unit').AsString = '千克(kg)' then
      s1 := 'kg';
    if Aqr.FieldByName('pd_unit').AsString = '吨(ton)' then
      s1 := 'kg';
    if Aqr.FieldByName('pd_unit').AsString = '万吨(10,000tons)' then
      s1 := 'kg';
    if Aqr.FieldByName('pd_unit').AsString = '立方米(m3)' then
      s1 := 'm3';
    if Aqr.FieldByName('pd_unit').AsString = '千立方米(km3)' then
      s1 := 'm3';
    if Aqr.FieldByName('pd_unit').AsString = '兆立方米(Mm3)' then
      s1 := 'm3';
    if Aqr.FieldByName('pd_unit').AsString = '万千瓦时(10,000kWh)' then
      s1 := 'kWh';
    if Aqr.FieldByName('pd_unit').AsString = '千瓦时(kWh)' then
      s1 := 'kWh';
    if Aqr.FieldByName('pd_unit').AsString = '吉焦(GJ)' then
      s1 := '兆焦(MJ)';
    if Aqr.FieldByName('pd_unit').AsString = '兆焦(MJ)' then
      s1 := '兆焦(MJ)';
    if Aqr.FieldByName('pd_unit').AsString = '千焦(KJ)' then
      s1 := '兆焦(MJ)';
    if Aqr.FieldByName('pd_unit').AsString = '个(件、罐、台...)' then
      s1 := '个(件、罐、台...)';

    result := s2 + '/' + s1;
  end
  else
    result := s2 + '/' + s2;
  Aqr.Close;
  Aqr.Free;
end;

function TForm1.standard_unit(s2: string): string; //转换为标准单位
begin
  result := s2;

  if s2 = '吨(ton)' then
    result := '千克(kg)';

  if s2 = '万吨(10,000tons)' then
    result := '千克(kg)';

  if s2 = '千立方米(km3)' then
    result := '立方米(m3)';

  if s2 = '兆立方米(Mm3)' then
    result := '立方米(m3)';

  if s2 = '万千瓦时(10,000kWh)' then
    result := '千瓦时(kWh)';

  if s2 = '千焦(KJ)' then
    result := '兆焦(MJ)';

  if s2 = '吉焦(GJ)' then
    result := '兆焦(MJ)';



end;


function TForm1.unit_exchange(s2: string): double; //返回单位换算系数

begin
  result := 1;

  if s2 = '吨(ton)' then
    result := 1000;
  if s2 = '万吨(10,000tons)' then
    result := 10000000;
  if s2 = '千立方米(km3)' then
    result := 1000;
  if s2 = '兆立方米(Mm3)' then
    result := 1000000;
  if s2 = '万千瓦时(10,000kWh)' then
    result := 10000;
  if s2 = '千焦(KJ)' then
    result := 0.001;
  if s2 = '吉焦(GJ)' then
    result := 1000;

end;


function TForm1.co_num(process, num, s2: string): double; //返回系数
//num：数量  s2:单位
var
  mysql: string;
  Aqr: TAdoQuery;
  s, a: double; //sum:产品产量   s:物料数量
begin
  try
    s := strtofloat(num);
  except
    s := 0;
  end;

  if s2 = '吨(ton)' then
    s := s * 1000;
  if s2 = '万吨(10,000tons)' then
    s := s * 10000000;
  if s2 = '千立方米(km3)' then
    s := s * 1000;
  if s2 = '兆立方米(Mm3)' then
    s := s * 1000000;
  if s2 = '万千瓦时(10,000kWh)' then
    s := s * 10000;
  if s2 = '千焦(KJ)' then
    s := s * 0.001;
  if s2 = '吉焦(GJ)' then
    s := s * 1000;


  mysql := 'select * from data where process = ' + #39 + process + #39 + ' and material_type =' + #39 + '产品' + #39;
  Aqr := TAdoQuery.Create(nil);
  Aqr.Connection := form1.ADOConnection2;
  Aqr.Close;
  Aqr.SQL.Add(mysql);
  Aqr.Open;


  if Aqr.RecordCount > 0 then
  begin
    Aqr.First;
    a := Aqr.FieldByName('pd').AsFloat;

    if Aqr.FieldByName('pd_unit').AsString = '吨(ton)' then
      a := Aqr.FieldByName('pd').AsFloat * 1000;
    if Aqr.FieldByName('pd_unit').AsString = '万吨(10,000tons)' then
      a := Aqr.FieldByName('pd').AsFloat * 10000000;
    if Aqr.FieldByName('pd_unit').AsString = '千立方米(km3)' then
      a := Aqr.FieldByName('pd').AsFloat * 1000;
    if Aqr.FieldByName('pd_unit').AsString = '兆立方米(Mm3)' then
      a := Aqr.FieldByName('pd').AsFloat * 1000000;
    if Aqr.FieldByName('pd_unit').AsString = '万千瓦时(10,000kWh)' then
      a := Aqr.FieldByName('pd').AsFloat * 10000;
    if Aqr.FieldByName('pd_unit').AsString = '千焦(KJ)' then
      a := Aqr.FieldByName('pd').AsFloat * 0.001;

    result := s / a;

  end
  else
    result := 1;
  Aqr.Close;
  Aqr.Free;
end;


function TForm1.co_num1(process, num, s2: string): double; //返回系数
//num：数量  s2:单位
var
  mysql: string;
  Aqr: TAdoQuery;
  s, a: double; //sum:产品产量   s:物料数量
begin
  try
    s := strtofloat(num);
  except
    s := 0;
  end;

  if s2 = '吨(ton)' then
    s := s * 1000;
  if s2 = '万吨(10,000tons)' then
    s := s * 10000000;
  if s2 = '千立方米(km3)' then
    s := s * 1000;
  if s2 = '兆立方米(Mm3)' then
    s := s * 1000000;
  if s2 = '万千瓦时(10,000kWh)' then
    s := s * 10000;
  if s2 = '千焦(KJ)' then
    s := s * 0.001;
  if s2 = '吉焦(GJ)' then
    s := s * 1000;


  mysql := 'select * from data_allocation where process = ' + #39 + process + #39 + ' and material_type =' + #39 + '产品' + #39;
  Aqr := TAdoQuery.Create(nil);
  Aqr.Connection := form1.ADOConnection2;
  Aqr.Close;
  Aqr.SQL.Add(mysql);
  Aqr.Open;


  if Aqr.RecordCount > 0 then
  begin
    Aqr.First;
    a := Aqr.FieldByName('pd').AsFloat;

    if Aqr.FieldByName('pd_unit').AsString = '吨(ton)' then
      a := Aqr.FieldByName('pd').AsFloat * 1000;
    if Aqr.FieldByName('pd_unit').AsString = '万吨(10,000tons)' then
      a := Aqr.FieldByName('pd').AsFloat * 10000000;
    if Aqr.FieldByName('pd_unit').AsString = '千立方米(km3)' then
      a := Aqr.FieldByName('pd').AsFloat * 1000;
    if Aqr.FieldByName('pd_unit').AsString = '兆立方米(Mm3)' then
      a := Aqr.FieldByName('pd').AsFloat * 1000000;
    if Aqr.FieldByName('pd_unit').AsString = '万千瓦时(10,000kWh)' then
      a := Aqr.FieldByName('pd').AsFloat * 10000;
    if Aqr.FieldByName('pd_unit').AsString = '千焦(KJ)' then
      a := Aqr.FieldByName('pd').AsFloat * 0.001;

    result := s / a;

  end
  else
    result := 1;
  Aqr.Close;
  Aqr.Free;
end;

{procedure TForm1.toMJ;   //四种煤气、蒸汽 换算成热值单位
var


begin


焦炉煤气
高炉煤气
转炉煤气
天然气

蒸汽
end;    }





procedure TForm1.bsSkinButton45Click(Sender: TObject);
var
  i, j, k: integer;
  mysql1, s: string;
  Aqr1, Aqr2, Aqr3: TAdoQuery;
  s1, s2: widestring;
  SL: TStrings;
begin
  bsSkinButton44Click(Sender); //保存
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := ListBox2.Items.Count;


  StatusBar1.Panels[3].Text := '检查数据输入的完整性';
  StatusBar1.Refresh;



  if ADOQuery1.Active then ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add('select * from  data ');
  ADOQuery1.Open;
  for j := ADOQuery1.RecordCount downto 1 do
  begin
    ADOQuery1.RecNo := j;
    if ListBox2.Items.IndexOf(ADOQuery1.FieldByName('process').AsString) = -1 then
      ADOQuery1.Delete;
  end; //  删除冗余数据  例如载入流程图后，原先残留的数据


  for i := 1 to ListBox2.Items.Count - 1 do //检查数据输入的完整性
  begin
    ListBox2.ItemIndex := i;
    FlatProgressBar1.Position := i + 1;
    ProcessData_Format;

    if ADOQuery1.RecordCount < 2 then
    begin
      form1.bsSkinMessage1.MessageDlg(ListBox2.Items[i] + '工序数据不足，请补充！', (mtInformation), [mbOK], 0);
      ListBox2.Selected[i] := true;
      ListBox2Click(Sender);
      exit;
    end;

    if ADOQuery1.Locate('material_type', '产品', []) = false then
    begin
      form1.bsSkinMessage1.MessageDlg(ListBox2.Items[i] + '工序没有产品输入！', (mtInformation), [mbOK], 0);
      ListBox2.Selected[i] := true;
      ListBox2Click(Sender);
      exit;
    end;

 //产品名称不能重复







    if ADOQuery1.Locate('pd_name', null, []) = true then
    begin
      form1.bsSkinMessage1.MessageDlg('"' + ListBox2.Items[i] + '" 中名称未输入!', (mtInformation), [mbOK], 0);
      exit;
    end;



    if ADOQuery1.Locate('material_type', '产品', []) then
      if trim(ADOQuery1.FieldByName('processtype').AsString) = '' then
      begin
        form1.bsSkinMessage1.MessageDlg('"' + ListBox2.Items[i] + '" 中，"' + ADOQuery1.FieldByName('pd_name').AsString + '" 工序类别未输入!', (mtInformation), [mbOK], 0);
        exit;
      end;

    if ADOQuery1.Locate('pd_unit', null, []) = true then
    begin
      form1.bsSkinMessage1.MessageDlg('"' + ListBox2.Items[i] + '" 中，"' + ADOQuery1.FieldByName('pd_name').AsString + '" 单位未输入!', (mtInformation), [mbOK], 0);
      exit;
    end;

    if ADOQuery1.Locate('material_type', null, []) = true then
    begin
      form1.bsSkinMessage1.MessageDlg('"' + ListBox2.Items[i] + '" 中，"' + ADOQuery1.FieldByName('pd_name').AsString + '" 类别未输入!', (mtInformation), [mbOK], 0);
      exit;
    end;

    if ADOQuery1.Locate('pd', null, []) = true then
    begin
      form1.bsSkinMessage1.MessageDlg('"' + ListBox2.Items[i] + '" 中，"' + ADOQuery1.FieldByName('pd_name').AsString + '" 数据未输入!', (mtInformation), [mbOK], 0);
      exit;
    end;

    for j := 1 to ADOQuery1.RecordCount do
    begin
      ADOQuery1.RecNo := j;
      for k := 0 to 10 do
      begin
        if pdunit[k] = ADOQuery1.FieldByName('pd_unit').AsString then
          break;
        if k = 10 then
          if pdunit[k] <> ADOQuery1.FieldByName('pd_unit').AsString then
          begin
            form1.bsSkinMessage1.MessageDlg('"' + ListBox2.Items[i] + '" 中，"' + ADOQuery1.FieldByName('pd_name').AsString + '" 单位不在规定范围中!', (mtInformation), [mbOK], 0);
            exit;
          end;
      end;


    end;

  end;


  Aqr1 := TAdoQuery.Create(nil);
  Aqr2 := TAdoQuery.Create(nil);
  Aqr3 := TAdoQuery.Create(nil);

  try
    Aqr1.Connection := Form1.ADOConnection2;
    Aqr2.Connection := Form1.ADOConnection2;
    Aqr3.Connection := Form1.ADOConnection2;
    StatusBar1.Panels[3].Text := '检查数据产品名称是否重复';
    StatusBar1.Refresh;


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from  data where material_type=' + #39 + '产品' + #39);
    Aqr1.Open;





    s1 := trim(FlatEdit6.Text);
    s2 := stringreplace(s1, ';', #13, [rfReplaceAll]);
    SL := tstringList.Create;
    SL.text := s2;
    for i := 0 to SL.count - 1 do
    begin
      if Aqr1.Locate('pd_name', trim(SL.strings[i]), [loCaseInsensitive]) = false then
      begin
        form1.bsSkinMessage1.MessageDlg('项目基本信息中的<产品名称（最后工序）>  :   [' + trim(SL.strings[i]) + ']    未在工序数据中找到，请修改项目基本信息或修改工序数据。', (mtInformation), [mbOK], 0);
        exit;
      end;
    end;




    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from  data where material_type=' + #39 + '产品' + #39 + ' and pd_name=' + #39 + Aqr1.FieldByName('pd_name').AsString + #39);
      Aqr2.Open;
      if Aqr2.RecordCount > 1 then
      begin
        form1.bsSkinMessage1.MessageDlg(Aqr1.FieldByName('process').AsString + ' 的产品名称与其它工序产品名称重复，请检查修改!', (mtInformation), [mbOK], 0);
        ListBox2.ClearSelection;
        for j := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := j;
          ListBox2.Selected[ListBox2.Items.IndexOf(Aqr2.FieldByName('process').AsString)] := true;
        end;
        ListBox2Click(Sender);
        exit;
      end;
    end;

    StatusBar1.Panels[3].Text := '检查物流对应关系';
    StatusBar1.Refresh;


     //检查箭头，物流关系是LCA建模基础       ****************************************************

 //检查是不是每个主工序方块都有箭头连接         // 能源工序 怎么与主工序建立联系？  主生产工序');  能源公辅工序

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from  data where processtype=' + #39 + '主生产工序' + #39 + ' and material_type=' + #39 + '产品' + #39);
    Aqr1.Open;



    for i := 0 to dxFlowChart1.ObjectCount - 1 do //检查主工序是不是每个方块都有箭头连接
      if Aqr1.Locate('process', dxFlowChart1.Objects[i].Text, []) = true then
        if dxFlowChart1.Objects[i].ConnectionCount = 0 then
        begin
          form1.bsSkinMessage1.MessageDlg('"' + dxFlowChart1.Objects[i].Text + '" 过程没有箭头连接到其它过程!', (mtInformation), [mbOK], 0);
          exit;
        end;
                                       //          ('能源及能源介质');    原材料     辅助材料   大气排放 水体排放


    for i := 0 to dxFlowChart1.ConnectionCount - 1 do //检查箭头连接 的物流名称是否对应
    begin
      if dxFlowChart1.Connections[i].ObjectSource = nil then continue;
      if dxFlowChart1.Connections[i].ObjectDest = nil then continue;
      if dxFlowChart1.Connections[i].ObjectDest.Text = '至各工序' then continue;


      mysql1 := 'select * from data where process = ' + #39 + dxFlowChart1.Connections[i].ObjectSource.Text + #39 + ' and (material_type = ' + #39 + '产品' + #39 + ' or material_type = ' + #39 + '副产品或固体废弃物' + #39 + ')';
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add(mysql1);
      Aqr2.Open;

      mysql1 := 'select * from data where process = ' + #39 + dxFlowChart1.Connections[i].ObjectDest.Text + #39 + ' and (material_type = ' + #39 + '原材料' + #39 + ' or material_type = ' + #39 + '辅助材料' + #39 + ' or material_type = ' + #39 + '能源及能源介质' + #39 + ')';
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add(mysql1);
      Aqr1.Open;

      flag := false;
      for j := 1 to Aqr2.RecordCount do
      begin
        Aqr2.RecNo := j;
        if Aqr1.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
        begin
          flag := true;
          break;
        end;
      end;
      if flag = false then
      begin
        ListBox2.ClearSelection;
        form1.bsSkinMessage1.MessageDlg(dxFlowChart1.Connections[i].ObjectSource.Text + '----' + dxFlowChart1.Connections[i].ObjectDest.Text + '  没有物流数据联系，与工序流程图所示关系不符，请检查数据名称是否对应。', (mtInformation), [mbOK], 0);
        ListBox2.ItemIndex := ListBox2.Items.IndexOf(dxFlowChart1.Connections[i].ObjectSource.Text);
        ListBox2.Selected[ListBox2.ItemIndex] := true;
        ListBox2.Selected[ListBox2.Items.IndexOf(dxFlowChart1.Connections[i].ObjectDest.Text)] := true;
        ListBox2Click(Sender);
        exit;
      end
    end;




    StatusBar1.Panels[3].Text := '正在计算单耗与排放系数';
    StatusBar1.Refresh;



 //计算系数  , 并复制到data_allocation  ****************************************************





    Aqr2.Connection := Form1.ADOConnection2;
    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('delete * from data_allocation');
    Aqr2.ExecSQL;


    s := 'processtype';

    FlatProgressBar1.Position := 0;
    FlatProgressBar1.Max := ListBox2.Items.Count - 1;

    for i := 0 to ListBox2.Items.Count - 1 do
    begin
      FlatProgressBar1.Position := i;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from  data where process=' + #39 + ListBox2.Items[i] + #39);
      Aqr1.Open;
      if Aqr1.Locate('material_type', '产品', []) = true then
        s := Aqr1.FieldByName('processtype').AsString;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation  where process=' + #39 + ListBox2.Items[i] + #39);
      Aqr2.Open;

      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        Aqr1.Edit;
        Aqr1.FieldByName('co_unit').AsString := co_unit(ListBox2.Items[i], Aqr1.FieldByName('pd_unit').AsString);
        Aqr1.FieldByName('co_num').AsFloat := co_num(ListBox2.Items[i], Aqr1.FieldByName('pd').AsString, Aqr1.FieldByName('pd_unit').AsString);
        Aqr1.FieldByName('system_name').AsString := FlatEdit23.Text;
        Aqr1.FieldByName('processtype').AsString := s;
        Aqr1.Post;
      end;

      if Aqr1.RecordCount > 0 then
      begin
        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          Aqr2.Append;
          for k := 1 to Aqr1.FieldCount - 1 do
            Aqr2.Fields[k].AsString := Aqr1.Fields[k].AsString;
          Aqr2.Post;
        end;
      end;

    end;




  //  检查单位的一致性****************************************************

    FlatProgressBar1.Position := 0;

    StatusBar1.Panels[3].Text := '检查单位的一致性';
    StatusBar1.Refresh;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation ');
    Aqr1.Open;
    FlatProgressBar1.Max := Aqr1.RecordCount;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      FlatProgressBar1.Position := i;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where pd_name=' + #39 + Aqr1.FieldByName('pd_name').AsString + #39);
      Aqr2.Open;
      if Aqr2.RecordCount > 1 then
      begin
        Aqr2.First;
        s := Aqr2.FieldByName('co_unit').AsString;
        for j := 2 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := j;
          if copy(s, 1, pos('/', s) - 1) <> copy(Aqr2.FieldByName('co_unit').AsString, 1, pos('/', Aqr2.FieldByName('co_unit').AsString) - 1) then
          begin
            form1.bsSkinMessage1.MessageDlg(Aqr2.FieldByName('pd_name').AsString + '单位体系不一致: ', (mtInformation), [mbOK], 0);
            ListBox2.ItemIndex := 0;
            ListBox2Click(Sender);
            DBGridEh9.DataSource.DataSet.Filter := 'pd_name = ' + #39 + Aqr2.FieldByName('pd_name').AsString + #39;
            DBGridEh9.DataSource.DataSet.Filtered := true;
            exit;
          end;

        end;




      end;
    end;






    FlatProgressBar1.Position := 0;




 //******************************************************************


    StatusBar1.Panels[3].Text := '查找需要分配的单元';
    StatusBar1.Refresh;


//列出需要分配的单元
    AdoQuery2.Connection := Form1.ADOConnection2;
    ListBox3.Clear;
    for i := 0 to ListBox2.Items.Count - 1 do
    begin
      mysql1 := 'select * from data where (material_type= ' + #39 + '产品' + #39 + ' or material_type=' + #39 + '副产品或固体废弃物' + #39 + ')' + 'and process=' + #39 + ListBox2.Items[i] + #39;

      if AdoQuery2.Active then AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Add(mysql1);
      ADOQuery2.Open;
      ADOQuery2.Active := true;

      if ADOQuery2.RecordCount > 1 then
      begin
        if ListBox3.Items.IndexOf(ListBox2.Items[i]) = -1 then
          ListBox3.Items.Add(ListBox2.Items[i]);
      end;

    end;


    Aqr2.Connection := Form1.ADOConnection2;
    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from allocation');
    Aqr2.open;

    Aqr1.Connection := Form1.ADOConnection2;
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data');
    Aqr1.open;
    StatusBar1.Panels[3].Text := '删除冗余数据';
    StatusBar1.Refresh;

    for j := Aqr2.RecordCount downto 1 do
    begin
      Aqr2.RecNo := j;
      if Aqr1.Locate('pd_name', Aqr2.FieldByName('albefore').AsString, []) = false then
        Aqr2.Delete;
    end; //删除冗余数据  数据名称变化后，分配表中的冗余数据

    for j := Aqr2.RecordCount downto 1 do
    begin
      Aqr2.RecNo := j;
      if ListBox2.Items.IndexOf(Aqr2.FieldByName('process').AsString) = -1 then
        Aqr2.Delete;
    end; //删除冗余数据   不在listbox工序列表中的数据





    if ListBox3.Items.Count = 0 then
    begin
      PageControl1.ActivePageIndex := 4;
      PageControl1.Pages[2].Highlighted := false;
      PageControl1.Pages[4].Highlighted := true;
      shuxing_Format;
    end
    else
    begin
      PageControl1.ActivePageIndex := 3;
      PageControl1.Pages[2].Highlighted := false;
      PageControl1.Pages[3].Highlighted := true;
      ListBox3.ItemIndex := 0;
      ListBox3Click(Sender);
      allocation_Format;
    end;




  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;

  end;

  StatusBar1.Panels[3].Text := '';
  StatusBar1.Refresh;
end;

procedure TForm1.bsSkinButton48Click(Sender: TObject);
var
  i: integer;
begin
  Panel7.Visible := false;
  if TreeView1.Items.Count = 0 then exit;
  TreeView1Renew;
  TreeView1.Items[0].Selected := true;
  TreeView1Click(Sender);
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := true;

  PageControl1.Pages[11].TabVisible := false;
  PageControl1.Pages[12].TabVisible := false;
end;

procedure TForm1.DBGridEh1GetCellParams(Sender: TObject; Column: TColumnEh;
  AFont: TFont; var Background: TColor; State: TGridDrawState);
begin
  if not (gdSelected in State) then
    if (Sender as TDBGridEh).SumList.RecNo mod 2 = 0 then
      Background := clInfobk;
end;

procedure TForm1.DBGridEh1KeyPress(Sender: TObject; var Key: Char);
begin
  if Label4.Caption = '外部LCA数据管理' then
  begin
    if dbgrideh1.SelectedField = adoquery1.FieldByName('LCItype') then //只能选择
      if key <> #13 then //允许使用回车
        key := #0; //什么也不干


    if dbgrideh1.SelectedField = adoquery1.FieldByName('LCIfactor') then //只能选择
      if key <> #13 then //允许使用回车
        key := #0; //什么也不干
  end;

  if Label4.Caption = 'LCI环境指标管理' then
  begin
    if dbgrideh1.SelectedField = adoquery1.FieldByName('inf_type') then //只能选择
      if key <> #13 then //允许使用回车
        key := #0; //什么也不干

  end;

  if Label4.Caption = '环境影响类型管理' then
  begin
    if dbgrideh1.SelectedField = adoquery1.FieldByName('impact_type') then //只能选择
      if key <> #13 then //允许使用回车
        key := #0; //什么也不干
  end;

  if Label4.Caption = '运输LCA基础数据管理' then
  begin
    if dbgrideh1.SelectedField = adoquery1.FieldByName('shuxing') then //只能选择
      if key <> #13 then //允许使用回车
        key := #0; //什么也不干

    if dbgrideh1.SelectedField = adoquery1.FieldByName('danwei') then //只能选择
      if key <> #13 then //允许使用回车
        key := #0; //什么也不干



  end;
end;

procedure TForm1.Edit2Click(Sender: TObject);
begin
  if Edit2.Text = '输入工序名称进行查询' then
    Edit2.Text := '';
end;

procedure TForm1.bsSkinButton44Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;
  mysql1: string;
begin

  ADOQuery1.UpdateBatch(arall);


  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := Form1.ADOConnection2;

  StatusBar1.Panels[3].Text := '检查是否有重复数据';
  StatusBar1.Refresh;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := ADOQuery1.RecordCount;

  try
    for i := 1 to ADOQuery1.RecordCount do
    begin
      ADOQuery1.RecNo := i;
      Aqr1.Close;
      Aqr1.SQL.Clear;
      mysql1 := 'select * from data where process=' + #39 + ListBox2.Items[ListBox2.ItemIndex] + #39 + ' and pd_name=' + #39 + ADOQuery1.FieldByName('pd_name').AsString + #39 + ' and material_type=' + #39 + ADOQuery1.FieldByName('material_type').AsString + #39;
      Aqr1.SQL.Add(mysql1);
      Aqr1.Open;

      if Aqr1.RecordCount > 1 then
      begin
        form1.bsSkinMessage1.MessageDlg('存在重复数据，请检查！', (mtInformation), [mbOK], 0);
        exit;
      end;

      FlatProgressBar1.Position := i;

    end;



  finally
    Aqr1.Close;
    Aqr1.Free;


  end;
     //     ID 0	system_name1 processtype2 	process3  	material_type4	  pd_name5	  pd_unit6  	pd7  	remark8
// ID	system_name	processtype	process	material_type	pd_name	pd_unit	pd	co_unit	co_num	remark






  ADOQuery1.First;
  FlatProgressBar1.Position := 0;


end;

procedure TForm1.bsSkinButton38Click(Sender: TObject);
var
  i: integer;
begin
  if DBGridEh9.SelectedRows.Count = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('请选择需要删除的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if form1.bsSkinMessage1.MessageDlg('确定删除所选记录？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := DBGridEh9.SelectedRows.Count - 1 downto 0 do
    begin
      DBGridEh9.DataSource.DataSet.GotoBookmark(pointer(DBGridEh9.SelectedRows[i]));
      DBGridEh9.DataSource.DataSet.Delete;
    end;
  end;

  ADOQuery1.UpdateBatch(arall);
 // if LCA_calc.Label6.Caption = '上游数据管理' then
 //    up_Format;
 //  if Label6.Caption = '环境负荷因子管理' then
 //    LCIfactor_Format;

end;

procedure TForm1.bsSkinButton51Click(Sender: TObject);
begin
  ADOQuery1.UpdateBatch(arall);
  ADOQuery1.First;
end;

procedure TForm1.bsSkinButton52Click(Sender: TObject);

begin
  ADOQuery1.UpdateBatch(arall);
  PageControl1.ActivePageIndex := 6;
  PageControl1.Pages[5].Highlighted := false;
  PageControl1.Pages[6].Highlighted := true;
  Label6Click(Sender);

end;

procedure TForm1.bsSkinButton54Click(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 4;
  PageControl1Change(Sender);
end;

procedure TForm1.bsSkinButton50Click(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 2;
  PageControl1Change(Sender);
end;

procedure TForm1.bsSkinButton47Click(Sender: TObject);
begin
  ADOQuery8.UpdateBatch(arall);
  ADOQuery8.First;
end;

procedure TForm1.ListBox3Click(Sender: TObject);

var
  i: integer;
  mysql1: string;
begin
  if ListBox3.ItemIndex = -1 then exit;
  ADOQuery2.Close;
  AdoQuery2.Connection := Form1.ADOConnection2;
  mysql1 := 'select * from data where process=' + #39 + ListBox3.Items[ListBox3.itemindex] + #39 + ' and( material_type= ' + #39 + '产品' + #39 + ' or material_type=' + #39 + '副产品或固体废弃物' + #39 + ')';
  if AdoQuery2.Active then AdoQuery2.Close;
  AdoQuery2.SQL.Clear;
  AdoQuery2.SQL.Add(mysql1);
  ADOQuery2.Open;
  ADOQuery2.Active := true;


  bsSkinCheckListBox9.Clear;
  if ADOQuery2.RecordCount > 1 then
  begin
    for i := 1 to ADOQuery2.RecordCount do
    begin
      ADOQuery2.RecNo := i;
      if bsSkinCheckListBox9.Items.IndexOf(ADOQuery2.FieldbyName('pd_name').AsString) = -1 then
        bsSkinCheckListBox9.Items.Add(ADOQuery2.FieldbyName('pd_name').AsString);
    end;
  end;

  allocation_Format;



end;

procedure TForm1.FlatEdit1Click(Sender: TObject);
begin
  if FlatEdit1.Text = '在此输入合并后名称' then
    FlatEdit1.Text := '';
end;

procedure TForm1.allocation_Format; //分配
var
  i: integer;
  mysql1: string;
begin
  if ListBox3.itemindex = -1 then exit;
  AdoQuery2.Connection := Form1.ADOConnection2;
  mysql1 := 'select * from allocation where process= ' + #39 + ListBox3.Items[ListBox3.itemindex] + #39;
//  ID	process	allocation	alnum	albefore	alafter remark


  if AdoQuery2.Active then AdoQuery2.Close;
  AdoQuery2.SQL.Clear;
  AdoQuery2.SQL.Add(mysql1);
  ADOQuery2.Open;
  ADOQuery2.Active := true;

  for i := 0 to DBGridEh5.Columns.Count - 1 do
  begin
    DBGridEh5.Columns[i].Width := 120;
    DBGridEh5.Columns[i].Title.Alignment := taCenter;
    DBGridEh5.Columns[i].Alignment := taCenter;
  end;
  DBGridEh5.Columns[3].Visible := false;

  DBGridEh5.Columns[0].Width := 40;
  DBGridEh5.Columns[3].Width := 60;
  DBGridEh5.Columns[6].Width := 180;
  DBGridEh5.Columns[0].Title.caption := '序号';
  DBGridEh5.Columns[1].Title.caption := '工序';
  DBGridEh5.Columns[2].Title.caption := '分配方法';
//  DBGridEh5.Columns[3].Title.caption := '分配次数';
  DBGridEh5.Columns[4].Title.caption := '分配前名称';
  DBGridEh5.Columns[5].Title.caption := '分配后名称';
  DBGridEh5.Columns[6].Title.caption := '分配说明';

  for i := 1 to ADOQuery2.RecordCount do
  begin
    ADOQuery2.RecNo := i;
    if bsSkinCheckListBox9.Items.IndexOf(ADOQuery2.FieldbyName('albefore').AsString) <> -1 then
      bsSkinCheckListBox9.Items.Delete(bsSkinCheckListBox9.Items.IndexOf(ADOQuery2.FieldbyName('albefore').AsString));
  end;



end;

procedure TForm1.bsSkinButton58Click(Sender: TObject);
var
  i, max_i: integer;
begin
  if radiogroup1.ItemIndex = 0 then //  设为主产品
  begin
    if ADOQuery2.Locate('allocation', '设为主产品', []) then
    begin
      form1.bsSkinMessage1.MessageDlg('已经设置过主产品了，主产品只能设置一个！', (mtInformation), [mbOK], 0);
      exit;
    end;

    if checkedcount(bsSkinCheckListBox9) <> 1 then
    begin
      form1.bsSkinMessage1.MessageDlg('主产品只能设置一个！', (mtInformation), [mbOK], 0);
      exit;
    end;

    for i := 0 to bsSkinCheckListBox9.Items.Count - 1 do
      if bsSkinCheckListBox9.Checked[i] = true then
      begin
        ADOQuery2.Append;
        ADOQuery2.FieldByName('process').AsString := ListBox3.Items[ListBox3.ItemIndex];
        ADOQuery2.FieldByName('allocation').AsString := RadioGroup1.Items[RadioGroup1.ItemIndex];
        ADOQuery2.FieldByName('alnum').AsInteger := 0;
        ADOQuery2.FieldByName('albefore').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('alafter').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('remark').AsString := FlatEdit25.Text;
        ADOQuery2.Post;
        ADOQuery2.Refresh;
      end;
    allocation_Format;
  end;

 //************************************************************************************

  if radiogroup1.ItemIndex = 1 then //  合并
  begin
    if checkedcount(bsSkinCheckListBox9) < 2 then
    begin
      form1.bsSkinMessage1.MessageDlg('请选择2个及以上的产品进行合并！', (mtInformation), [mbOK], 0);
      exit;
    end;

    if FlatEdit1.Text = '在此输入合并后名称' then
    begin
      form1.bsSkinMessage1.MessageDlg('请输入合并后名称！', (mtInformation), [mbOK], 0);
      exit;
    end;

    if trim(FlatEdit1.Text) = '' then
    begin
      form1.bsSkinMessage1.MessageDlg('请输入合并后名称！', (mtInformation), [mbOK], 0);
      exit;
    end;

    if bsSkinCheckListBox9.Items.IndexOf(trim(FlatEdit1.Text)) <> -1 then
    begin
      form1.bsSkinMessage1.MessageDlg('合并后的名称不能与现有名称重复！', (mtInformation), [mbOK], 0);
      exit;
    end;

    if listbox2.Items.IndexOf(trim(FlatEdit1.Text)) <> -1 then
    begin
      form1.bsSkinMessage1.MessageDlg('合并后的名称不能与现有名称重复！', (mtInformation), [mbOK], 0);
      exit;
    end;

    //  ID	process	allocation	alnum	albefore	alafter remark

    max_i := 0;
    for i := 1 to ADOQuery2.RecordCount do
    begin
      ADOQuery2.RecNo := i;
      if ADOQuery2.FieldByName('alnum').AsInteger > max_i then
        max_i := ADOQuery2.FieldByName('alnum').AsInteger;
    end;


    for i := 0 to bsSkinCheckListBox9.Items.Count - 1 do
      if bsSkinCheckListBox9.Checked[i] = true then
      begin
        ADOQuery2.Append;
        ADOQuery2.FieldByName('process').AsString := ListBox3.Items[ListBox3.ItemIndex];
        ADOQuery2.FieldByName('allocation').AsString := RadioGroup1.Items[RadioGroup1.ItemIndex];
        ADOQuery2.FieldByName('alnum').AsInteger := max_i + 1;
        ADOQuery2.FieldByName('albefore').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('alafter').AsString := trim(FlatEdit1.Text);
        ADOQuery2.FieldByName('remark').AsString := trim(FlatEdit25.Text);
        ADOQuery2.Post;
        ADOQuery2.Refresh;
      end;


    for i := bsSkinCheckListBox9.Items.Count - 1 downto 0 do
    begin
      if bsSkinCheckListBox9.Checked[i] = true then
        bsSkinCheckListBox9.Items.Delete(i);
    end;
    bsSkinCheckListBox9.Items.Add(trim(FlatEdit1.Text));

    allocation_Format;
  end;

 //************************************************************************************

  if radiogroup1.ItemIndex = 2 then //  分拆
  begin

    if bsSkinMessage1.MessageDlg('分拆工序需按分拆后的工序重新绘制流程图，重新编辑流程图？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin

      bsSkinButton6Click(Sender); //重新编辑流程图

    end;

   {
    max_i := 0;
    for i := 1 to ADOQuery2.RecordCount do
    begin
      ADOQuery2.RecNo := i;
      if ADOQuery2.FieldByName('alnum').AsInteger > max_i then
        max_i := ADOQuery2.FieldByName('alnum').AsInteger;
    end;



    listbox2.ClearSelection;
    for i := bsSkinCheckListBox9.Items.Count - 1 downto 0 do
      if bsSkinCheckListBox9.Checked[i] = true then
      begin
        ADOQuery2.Append;
        ADOQuery2.FieldByName('process').AsString := ListBox3.Items[ListBox3.ItemIndex];
        ADOQuery2.FieldByName('allocation').AsString := RadioGroup1.Items[RadioGroup1.ItemIndex];
        ADOQuery2.FieldByName('alnum').AsInteger := max_i + 1;
        ADOQuery2.FieldByName('albefore').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('alafter').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('remark').AsString := trim(FlatEdit25.Text);
        ADOQuery2.Post;
        ADOQuery2.Refresh;
     //   listbox2.Items.Add(bsSkinCheckListBox9.Items[i]);
     //   listbox2.Selected[listbox2.Items.Count - 1] := true;
        bsSkinCheckListBox9.Items.Delete(i);
      end;

 //   PageControl1.ActivePageIndex := 2;

    ProcessData_Format;
                          }

  end;


 //************************************************************************************

  if radiogroup1.ItemIndex = 3 then //  系统扩展法
  begin
    if checkedcount(bsSkinCheckListBox9) <> 1 then
    begin
      form1.bsSkinMessage1.MessageDlg('系统扩展法每次只能设置一个！', (mtInformation), [mbOK], 0);
      exit;
    end;



    if FlatEdit26.Text = '在此输入替代工序名称' then
    begin
      form1.bsSkinMessage1.MessageDlg('请输入替代工序名称！', (mtInformation), [mbOK], 0);
      exit;
    end;

    if trim(FlatEdit26.Text) = '' then
    begin
      form1.bsSkinMessage1.MessageDlg('请输入替代工序名称！', (mtInformation), [mbOK], 0);
      exit;
    end;


    max_i := 0;
    for i := 1 to ADOQuery2.RecordCount do
    begin
      ADOQuery2.RecNo := i;
      if ADOQuery2.FieldByName('alnum').AsInteger > max_i then
        max_i := ADOQuery2.FieldByName('alnum').AsInteger;
    end;

    for i := bsSkinCheckListBox9.Items.Count - 1 downto 0 do
      if bsSkinCheckListBox9.Checked[i] = true then
      begin
        ADOQuery2.Append;
        ADOQuery2.FieldByName('process').AsString := ListBox3.Items[ListBox3.ItemIndex];
        ADOQuery2.FieldByName('allocation').AsString := RadioGroup1.Items[RadioGroup1.ItemIndex];
        ADOQuery2.FieldByName('alnum').AsInteger := max_i + 1;
        ADOQuery2.FieldByName('albefore').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('alafter').AsString := trim(FlatEdit26.Text);
        ADOQuery2.FieldByName('remark').AsString := trim(FlatEdit25.Text);
        ADOQuery2.Post;
        ADOQuery2.Refresh;
        bsSkinCheckListBox9.Items.Delete(i);
      end;

    allocation_Format;
  end;

 //************************************************************************************
  if radiogroup1.ItemIndex = 4 then //  不分配
  begin
    if checkedcount(bsSkinCheckListBox9) <> 1 then
    begin
      form1.bsSkinMessage1.MessageDlg('每次只能选择一个！', (mtInformation), [mbOK], 0);
      exit;
    end;

    max_i := 0;
    for i := 1 to ADOQuery2.RecordCount do
    begin
      ADOQuery2.RecNo := i;
      if ADOQuery2.FieldByName('alnum').AsInteger > max_i then
        max_i := ADOQuery2.FieldByName('alnum').AsInteger;
    end;

    for i := bsSkinCheckListBox9.Items.Count - 1 downto 0 do
      if bsSkinCheckListBox9.Checked[i] = true then
      begin
        ADOQuery2.Append;
        ADOQuery2.FieldByName('process').AsString := ListBox3.Items[ListBox3.ItemIndex];
        ADOQuery2.FieldByName('allocation').AsString := RadioGroup1.Items[RadioGroup1.ItemIndex];
        ADOQuery2.FieldByName('alnum').AsInteger := max_i + 1;
        ADOQuery2.FieldByName('albefore').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('alafter').AsString := bsSkinCheckListBox9.Items[i];
        ADOQuery2.FieldByName('remark').AsString := trim(FlatEdit25.Text);
        ADOQuery2.Post;
        ADOQuery2.Refresh;
        bsSkinCheckListBox9.Items.Delete(i);
      end;

    allocation_Format;

  end;





end;





function TForm1.checkedcount(bsSkinCheckListBox1: TbsSkinCheckListBox): integer;
//返回选中个数
var
  i, j: integer;
begin
  j := 0;
  if bsSkinCheckListBox1.Items.Count = 0 then
  begin
    result := 0;
    exit;
  end
  else
    for i := 0 to bsSkinCheckListBox1.Items.Count - 1 do
    begin
      if bsSkinCheckListBox1.Checked[i] = true then
        j := j + 1;
    end;
  result := j;
end;



procedure TForm1.RadioGroup1Click(Sender: TObject);
begin
  if radiogroup1.ItemIndex = 1 then //  合并
  begin
    FlatEdit1.Text := '在此输入合并后名称';
    FlatEdit1.Visible := true;
  end
  else
    FlatEdit1.Visible := false;


  if radiogroup1.ItemIndex = 3 then //  合并
  begin
    FlatEdit26.Text := '在此输入替代工序名称';
    FlatEdit26.Visible := true;
  end
  else
    FlatEdit26.Visible := false;




end;

procedure TForm1.bsSkinButton59Click(Sender: TObject);
var
  i: integer;
begin
  if ADOQuery2.RecordCount = 0 then exit;

  if bsSkinMessage1.MessageDlg('重新分配将删除原有分配方案，确认重新分配？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin

    for i := ADOQuery2.RecordCount downto 1 do
    begin
      ADOQuery2.RecNo := i;
      ADOQuery2.Delete;

    end;
  end;

  ADOQuery2.Refresh;
  ListBox3Click(Sender);





end;

procedure TForm1.FlatEdit26Click(Sender: TObject);
var
  i: integer;
begin
  if FlatEdit26.Text = '在此输入替代工序名称' then
    for i := 0 to bsSkinCheckListBox9.Items.Count - 1 do
      if bsSkinCheckListBox9.Checked[i] then
        FlatEdit26.Text := bsSkinCheckListBox9.Items[i];
end;



procedure TForm1.bsSkinButton49Click(Sender: TObject);
var
  i: integer;
  a: double;
begin
  ADOQuery8.UpdateBatch(arall);
  a := 0;
  for i := 1 to ADOQuery8.RecordCount do
  begin
    ADOQuery8.RecNo := i;
    if ADOQuery8.FieldByName('shuxing').AsString = '电力构成' then
      a := a + ADOQuery8.FieldByName('shuzhi').AsFloat;
  end;



  if abs(a - 100) > 0.000001 then
  begin
    form1.bsSkinMessage1.MessageDlg('电力构成比例之和必须为100%，请修改！', (mtInformation), [mbOK], 0);
    exit;
  end;

 { Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := Form1.ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := Form1.ADOConnection2;

  try

    for j := 1 to ADOQuery8.RecordCount do //先删除旧数据
    begin
      ADOQuery8.RecNo := j;
      if ADOQuery8.FieldByName('shuxing').AsString = '电力构成' then
      begin
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('delete * from  data_allocation where pd_name= ' + #39 + ADOQuery8.FieldByName('name').AsString + #39 + ' and material_type <> ' + #39 + '产品' + #39);
        Aqr2.ExecSQL;
      end;
    end;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from  data_allocation ');
    Aqr1.Open;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from  data_allocation ');
    Aqr2.Open;



    for i := Aqr1.RecordCount downto 1 do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('pd_name').AsString = '电' then
      begin
        a := 0;
        for j := 1 to ADOQuery8.RecordCount do
        begin
          ADOQuery8.RecNo := j;
          if ADOQuery8.FieldByName('shuxing').AsString = '电力构成' then
          begin
            a := ADOQuery8.FieldByName('shuzhi').AsFloat;

            Aqr2.Append;
            for k := 1 to Aqr1.FieldCount - 1 do
              Aqr2.Fields[k].AsString := Aqr1.Fields[k].AsString;

            Aqr2.FieldByName('pd').AsFloat := a / 100 * Aqr1.FieldByName('pd').AsFloat;
            Aqr2.FieldByName('co_num').AsFloat := a / 100 * Aqr1.FieldByName('co_num').AsFloat;
            Aqr2.FieldByName('pd_name').AsString := ADOQuery8.FieldByName('name').AsString;
            Aqr2.Post;
          end; // '电力构成'

        end; // for j
        Aqr1.Delete;
      end; //    = '电'
    end;




  finally
    Aqr1.Close;
    Aqr1.Free;
  end;
             }

  PageControl1.ActivePageIndex := 5;
  PageControl1.Pages[4].Highlighted := false;
  PageControl1.Pages[5].Highlighted := true;
  TransportData_Format;

end;

procedure TForm1.bsSkinButton28Click(Sender: TObject);
var
  i, j, k, max_i: integer;
  a: double;
  Aqr1, Aqr2, Aqr3: TAdoQuery;
  mysql: string;
begin

  for i := 0 to Listbox3.Items.Count - 1 do
  begin
    ListBox3.ItemIndex := i;
    ListBox3Click(Sender);
    if bsSkinCheckListBox9.Items.Count > 0 then
    begin
      form1.bsSkinMessage1.MessageDlg('分配未完成！', (mtInformation), [mbOK], 0);
      ListBox3.Selected[i] := true;
      exit;
    end;
  end;



  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := Listbox3.Items.Count;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := Form1.ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := Form1.ADOConnection2;
  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := Form1.ADOConnection2;

  try




//data_allocation
//ID0	system_name1	processtype2	process3	material_type4	pd_name5	pd_unit6	pd7	co_unit8	co_num9	remark10	elementflow11	shuxing12	transport13	upstream14
//data
//ID	system_name	processtype	process	material_type	pd_name	pd_unit	pd	co_unit	co_num	remark


    for i := 0 to Listbox3.Items.Count - 1 do
    begin
      FlatProgressBar1.Position := i + 1;
      ListBox3.ItemIndex := i;
      ListBox3Click(Sender);

      max_i := 1;
      for j := 1 to ADOQuery2.RecordCount do
      begin
        ADOQuery2.RecNo := j;
        if ADOQuery2.FieldByName('alnum').AsInteger > max_i then
          max_i := ADOQuery2.FieldByName('alnum').AsInteger;
      end;



      for j := 1 to max_i do //  第j次合并
      begin

        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from  data_allocation where process=' + #39 + ListBox3.Items[i] + #39);
        Aqr1.Open;


        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        mysql := 'select * from allocation where process=' + #39 + ListBox3.Items[i] + #39 + ' and alnum=' + inttostr(j) + ' and allocation=' + #39 + '合并' + #39;
        Aqr2.SQL.Add(mysql);
        Aqr2.Open;

        a := 0;
        if Aqr2.RecordCount > 0 then
        begin
          for k := 1 to Aqr2.RecordCount do
          begin
            Aqr2.RecNo := k;
            if Aqr1.Locate('pd_name', Aqr2.FieldByName('albefore').AsString, []) = true then
            begin
              a := a + Aqr1.FieldByName('pd').AsFloat * unit_exchange(Aqr1.FieldByName('pd_unit').AsString);
             //   Aqr1.Delete;
            end;
          end;
        end; // 合并求和 须考虑单位换算问题 a



        if Aqr3.Active then Aqr3.Close;
        Aqr3.SQL.Clear;
        Aqr3.SQL.Add('select * from  data_allocation where process=' + #39 + ListBox3.Items[i] + #39);
        Aqr3.Open;

        if Aqr2.RecordCount > 0 then
        begin
          if Aqr1.Locate('pd_name', Aqr2.FieldByName('albefore').AsString, []) = true then
          begin
            if Aqr3.Locate('pd_name', Aqr2.FieldByName('alafter').AsString, []) = false then
              Aqr3.Append
            else
              Aqr3.Edit;

            for k := 1 to Aqr1.FieldCount - 1 do
              Aqr3.Fields[k].AsString := Aqr1.Fields[k].AsString;
            Aqr3.FieldByName('pd_name').AsString := Aqr2.FieldByName('alafter').AsString;
            Aqr3.FieldByName('material_type').AsString := '产品';
            Aqr3.FieldByName('pd_unit').AsString := standard_unit(Aqr3.FieldByName('pd_unit').AsString);
            Aqr3.FieldByName('pd').AsFloat := a;
            Aqr3.Post;
          end; //新增一条合并后的名称的记录


          for k := 1 to Aqr2.RecordCount do
          begin
            Aqr2.RecNo := k;
            if Aqr1.Locate('pd_name', Aqr2.FieldByName('albefore').AsString, []) = true then
            begin
              Aqr1.Delete;
            end;
          end;


        end;

      end; //第j次合并
    end; //for listbox


    //计算系数

    for i := 0 to Listbox3.Items.Count - 1 do
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from  data_allocation where process=' + #39 + ListBox3.Items[i] + #39);
      Aqr1.Open;

      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        Aqr1.Edit;
        Aqr1.FieldByName('co_unit').AsString := co_unit1(ListBox3.Items[i], Aqr1.FieldByName('pd_unit').AsString);
        Aqr1.FieldByName('co_num').AsFloat := co_num1(ListBox3.Items[i], Aqr1.FieldByName('pd').AsString, Aqr1.FieldByName('pd_unit').AsString);
        Aqr1.FieldByName('system_name').AsString := FlatEdit23.Text;
        Aqr1.Post;
      end;
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
  end;

  shuxing_Format;
  PageControl1.ActivePageIndex := 4;
  PageControl1.Pages[3].Highlighted := false;
  PageControl1.Pages[4].Highlighted := true;
  FlatProgressBar1.Position := 0;


end;

procedure TForm1.ADOQuery2AfterOpen(DataSet: TDataSet);
begin
  ADOQuery2.FieldByName('ID').OnGetText := SetFdText2;
end;

procedure TForm1.StatusBar1DrawPanel(StatusBar: TStatusBar;
  Panel: TStatusPanel; const Rect: TRect);
begin

  FlatProgressBar1.Parent := StatusBar1;
  FlatProgressBar1.Left := Rect.Left;
  FlatProgressBar1.Top := Rect.Top;
  FlatProgressBar1.Width := Panel.Width;
  FlatProgressBar1.Height := Rect.Bottom - Rect.Top;


end;

procedure TForm1.FlatEdit27KeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9', #8, #13, #46]) then key := #0;
end;

procedure TForm1.FlatEdit28KeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9', #8, #13, #46]) then key := #0;
end;

procedure TForm1.FlatEdit29KeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9', #8, #13, #46]) then key := #0;
end;

procedure TForm1.Label24Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;
begin
  bsSkinPanel2.Visible := true; //功能单位
  Label4.Caption := Label1.Caption + '1';
  for i := 0 to 12 do
    PageControl1.Pages[i].TabVisible := false;

  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := true;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct product from up');
    Aqr1.Open;
    ListBox4.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      ListBox4.Items.Add(Aqr1.FieldByName('product').AsString);
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;

  AdoQuery1.Close;
  ListBox4Click(Sender);


  {  Panel7.Visible := false;
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  TreeView1.Hint := '基础信息管理';
  TreeView1.Items.Clear;
  PageControl1.Visible := true;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := false;

  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := false;
  rootnode := TreeView1.items.add(nil, '基础信息管理');
  TreeView1.Items.Addchild(rootnode, '外部LCA数据管理');
  TreeView1.Items.Addchild(rootnode, '运输LCA基础数据管理');
  TreeView1.Items.Addchild(rootnode, 'LCI环境指标管理');
  TreeView1.Items.Addchild(rootnode, '环境影响类型管理');

  rootnode := TreeView1.items.add(nil, '对照表管理');
  TreeView1.Items.Addchild(rootnode, '物料属性名称对照表');
  TreeView1.Items.Addchild(rootnode, '运输名称对照表');
  TreeView1.Items.Addchild(rootnode, '上游数据名称对照表');
  TreeView1.Items.Addchild(rootnode, 'LCI因子名称对照表');


  TreeView1.FullExpand;
  TreeView1.Items[1].Selected := true;
  TreeView1.SetFocus;
  StatusBar1.Panels[1].Text := '基础信息管理： ' + TreeView1.Selected.Text;

  Image1.Visible := false;
  bsSkinPanel2.Visible := true; //功能单位

  Label4.Caption := '外部LCA数据管理';

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection1;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct product from up');
    Aqr1.Open;
    ListBox4.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      ListBox4.Items.Add(Aqr1.FieldByName('product').AsString);
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;

  AdoQuery1.Close;
  ListBox4Click(Sender);}
end;




function TForm1.CO2(process: string): double; //返回CO2排放系数  g/kg  g/m3  g/kWh;
var
  j: integer;
  Aqr1, Aqr2, Aqr4: TAdoQuery;
  CO2, a: double;
begin
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;

  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection1;


  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from shuxing where shuxing = ' + #39 + 'CO2排放系数' + #39);
    Aqr1.Open;

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, shuxing from comparison where shuxing is not null ');
    Aqr4.Open;



    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
    Aqr2.Open;


    CO2 := 0; a := 0;
    for j := 1 to Aqr2.RecordCount do
    begin
      Aqr2.RecNo := j;
      if Aqr4.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
      begin
        if Aqr1.Locate('name', Aqr4.FieldByName('shuxing').AsString, []) = true then
        begin
          if Aqr2.FieldByName('material_type').AsString = '产品' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '副产品或固体废弃物' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '能源及能源介质' then
            a := Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '原材料' then
            a := Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '辅助材料' then
            a := Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '大气排放' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '水体排放' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;

          CO2 := CO2 + a * Aqr1.FieldByName('shuzhi').AsFloat;
        end;
      end;




    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;

    Aqr4.Close;
    Aqr4.Free;

  end;


  result := CO2 * 1000;

end;


function TForm1.energy(process: string): double; //返回能耗  MJ/kg  MJ/m3  MJ/kWh;
var
  j: integer;
  Aqr1, Aqr2, Aqr4: TAdoQuery;
  e, a: double;
begin
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection1;

  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from shuxing where shuxing = ' + #39 + '热值' + #39);
    Aqr1.Open;

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, shuxing from comparison where shuxing is not null ');
    Aqr4.Open;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
    Aqr2.Open;

    e := 0; a := 0;
    for j := 1 to Aqr2.RecordCount do
    begin
      Aqr2.RecNo := j;
      if Aqr4.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
      begin
        if Aqr1.Locate('name', Aqr4.FieldByName('shuxing').AsString, []) = true then
        begin
          if Aqr2.FieldByName('material_type').AsString = '产品' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '副产品或固体废弃物' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '能源及能源介质' then
            a := Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '原材料' then
            a := Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '辅助材料' then
            a := Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '大气排放' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
          if Aqr2.FieldByName('material_type').AsString = '水体排放' then
            a := (-1) * Aqr2.FieldByName('co_num').AsFloat;

          e := e + a * Aqr1.FieldByName('shuzhi').AsFloat;
        end;
      end;




    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;

    Aqr4.Close;
    Aqr4.Free;

  end;


  result := e;

end;




procedure TForm1.direct(process: string);
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  lci, a: double;
  product_name: string;
begin
  // 输入 － 输出
  //查询 LCIfactor，对照表，        黄河水 除盐水        (r)水耗
  //指标1 查询对应的工序中名称，如存在，则＋1 或－1

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection2;
  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection1;


  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from LCIfactor order by ID ASC');
    Aqr1.Open;

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
    Aqr4.Open;

    if Aqr4.Locate('material_type', '产品', []) = true then
    begin
      product_name := Aqr4.FieldByName('pd_name').AsString;
    end
    else
    begin
      showmessage(process + ' 工序没有产品输入！');
      flag := false;
      exit;
    end;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from LCA where product=' + #39 + product_name + #39);
    Aqr2.Open;



    for i := 1 to Aqr1.RecordCount do //lcifactor
    begin
      Aqr1.RecNo := i;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select id, pd_name, lcifactor from comparison where lcifactor= ' + #39 + Aqr1.FieldByName('inf_name').AsString + #39);
      Aqr3.Open;


      lci := 0; a := 0;
      for j := 1 to Aqr3.RecordCount do //  黄河水 除盐水        (r)水耗
      begin
        Aqr3.RecNo := j;

        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39 + ' and pd_name=' + #39 + Aqr3.FieldByName('pd_name').AsString + #39);
        Aqr4.Open;
        if Aqr4.RecordCount > 0 then
          for k := 1 to Aqr4.RecordCount do
          begin
            Aqr4.RecNo := k;
            if Aqr4.FieldByName('material_type').AsString = '产品' then
              a := (-1) * Aqr4.FieldByName('co_num').AsFloat;
            if Aqr4.FieldByName('material_type').AsString = '副产品或固体废弃物' then
              a := (-1) * Aqr4.FieldByName('co_num').AsFloat;
            if Aqr4.FieldByName('material_type').AsString = '能源及能源介质' then
              a := Aqr4.FieldByName('co_num').AsFloat;
            if Aqr4.FieldByName('material_type').AsString = '原材料' then
              a := Aqr4.FieldByName('co_num').AsFloat;
            if Aqr4.FieldByName('material_type').AsString = '辅助材料' then
              a := Aqr4.FieldByName('co_num').AsFloat;
            if Aqr4.FieldByName('material_type').AsString = '大气排放' then
              a := (-1) * Aqr4.FieldByName('co_num').AsFloat;
            if Aqr4.FieldByName('material_type').AsString = '水体排放' then
              a := (-1) * Aqr4.FieldByName('co_num').AsFloat;

            lci := lci + a;

          end;

      end;

      if Aqr2.Locate('LCIfactor', Aqr1.FieldByName('inf_name').AsString, []) = true then
        Aqr2.Edit
      else
        Aqr2.Append;

      Aqr2.FieldByName('product').AsString := product_name;
      Aqr2.FieldByName('LCItype').AsString := Aqr1.FieldByName('inf_type').AsString;
      Aqr2.FieldByName('LCIfactor').AsString := Aqr1.FieldByName('inf_name').AsString;
      Aqr2.FieldByName('unit').AsString := Aqr1.FieldByName('units').AsString;

      Aqr2.FieldByName('g1').asfloat := lci;
      if Aqr1.FieldByName('inf_type').AsString = '大气污染物' then
        Aqr2.FieldByName('g1').asfloat := (-1000) * lci;
      if Aqr1.FieldByName('inf_type').AsString = '水体污染物' then
        Aqr2.FieldByName('g1').asfloat := (-1000000) * lci;
      if Aqr1.FieldByName('inf_type').AsString = '固体废弃物' then
        Aqr2.FieldByName('g1').asfloat := (-1) * lci;

      if Aqr1.FieldByName('inf_name').AsString = '(a)二氧化碳' then
        Aqr2.FieldByName('g1').asfloat := CO2(process);
      if Aqr1.FieldByName('inf_name').AsString = '(e)总一次能源' then
        Aqr2.FieldByName('g1').asfloat := energy(process);
      Aqr2.FieldByName('see_level').AsString := Aqr1.FieldByName('see_level').AsString;
      Aqr2.Post;

    end;





  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;

end;



procedure TForm1.checkmat(mat: TMatrix); //调试用 检查矩阵
var
  i, j: integer;
begin
  Memo11.Clear;
  for j := 1 to mat.NrOfRows do
    for i := 1 to mat.NrOfColumns do
    begin
      Memo11.Text := Memo11.Text + '  ' + formatfloat('##0.##0', mat.Elem[i, j]);
      if i = mat.NrOfColumns then
        Memo11.Text := Memo11.Text + #13#10;
    end;

  Memo11.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + '1\' + mat.Name + '.txt');



 // showmessage(inttostr(mat.NrOfRows) + '    X    ' + inttostr(mat.NrOfColumns));
 {
  Memo11.Text := Memo11.Text + 'g1 :  ' + inttostr(g1.NrOfRows) + '    X    ' + inttostr(g1.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g2 :  ' + inttostr(g2.NrOfRows) + '    X    ' + inttostr(g2.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g3 :  ' + inttostr(g3.NrOfRows) + '    X    ' + inttostr(g3.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g4 :  ' + inttostr(g4.NrOfRows) + '    X    ' + inttostr(g4.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g5 :  ' + inttostr(g5.NrOfRows) + '    X    ' + inttostr(g5.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g6 :  ' + inttostr(g6.NrOfRows) + '    X    ' + inttostr(g6.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g7 :  ' + inttostr(g7.NrOfRows) + '    X    ' + inttostr(g7.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g_onsite :  ' + inttostr(g_onsite.NrOfRows) + '    X    ' + inttostr(g_onsite.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g_outsite  :  ' + inttostr(g_outsite.NrOfRows) + '    X    ' + inttostr(g_outsite.NrOfColumns) + #13#10;
  Memo11.Text := Memo11.Text + 'g_route  :  ' + inttostr(g_route.NrOfRows) + '    X    ' + inttostr(g_route.NrOfColumns) + #13#10;
 }
end;




procedure TForm1.mat_energy;
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4, Aqr5: TAdoQuery;
//  product_name: string;
begin
  //能源公辅工序 能源工序个数 确定矩阵行列数   单位矩阵   直接消耗系数矩阵

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection2;
  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection2;
  Aqr5 := TAdoQuery.Create(nil);
  Aqr5.Connection := ADOConnection2;

  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '计算能源工序LCA';
  StatusBar1.Refresh;

  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr1.Open;
    if Aqr1.RecordCount = 0 then exit;

    PB2.NrOfColumns := Aqr1.RecordCount;
    PB2.NrOfRows := Aqr1.RecordCount;
    for i := 1 to Aqr1.RecordCount do
      for j := 1 to Aqr1.RecordCount do
        PB2.Elem[i, j] := 0; //  elem[列，行]       PB2 行 工序  列 产品

    FlatProgressBar1.Position := 2;
//   ID	system_name	processtype	process	material_type	pd_name	pd_unit	pd	co_unit	co_num	remark	elementflow	shuxing	transport	upstream


    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
        Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
        Aqr2.Open;
        if Aqr2.RecordCount > 0 then
        begin
          Aqr2.First;
          PB2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
        end;

      end;

    end;


 //   {
    checkmat(PB2);
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      Memo11.Lines.Add(Aqr1.FieldByName('process').AsString);
    end;
    Memo11.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + '1\PB2' + '.txt');

 //         }



    FlatProgressBar1.Position := 5;



    I_PB2.NrOfColumns := PB2.NrOfColumns;
    I_PB2.NrOfRows := PB2.NrOfRows;



    for i := 1 to PB2.NrOfColumns do // I-PB2
      for j := 1 to PB2.NrOfRows do
        if i = j then
          I_PB2.Elem[i, j] := 1 - PB2.Elem[i, j]
        else
          I_PB2.Elem[i, j] := -PB2.Elem[i, j];


    I_PB2.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


    FlatProgressBar1.Position := 10;


           //    memo1.Lines.Add( );
    I_PB2_I.NrOfColumns := I_PB2.NrOfColumns;
    I_PB2_I.NrOfRows := I_PB2.NrOfRows;

    for i := 1 to I_PB2.NrOfColumns do // (I-PB2)-1-I
      for j := 1 to I_PB2.NrOfRows do
        if i = j then
          I_PB2_I.Elem[i, j] := I_PB2.Elem[i, j] - 1
        else
          I_PB2_I.Elem[i, j] := I_PB2.Elem[i, j];

    FlatProgressBar1.Position := 12;


    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCIfactor order by inf_name ASC');
    Aqr3.Open;

    g1_e.NrOfColumns := Aqr3.RecordCount;
    g1_e.NrOfRows := Aqr1.RecordCount;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      direct(Aqr1.FieldByName('process').AsString);
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCA where  product = ' + #39 + Aqr1.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr3.Open;
      for j := 1 to Aqr3.RecordCount do
      begin
        Aqr3.RecNo := j;
        g1_e.Elem[j, i] := Aqr3.FieldByName('g1').AsFloat; //  elem[列，行]
      end;
    end;

    g2_e.NrOfColumns := g1_e.NrOfColumns;
    g2_e.NrOfRows := g1_e.NrOfRows;

    I_PB2_I.Transpose; //  ((I-PB2)-1-I )T


   //  checkmat(g1_e);
   //  checkmat(I_PB2_I);

    I_PB2_I.Multiply(g1_e, g2_e); //g2-e计算出来了           * 2*2     2*53



    g2_e.Transpose; //53*2



    FlatProgressBar1.Position := 15;



// ************************************************************************************



    //三种煤气  RB2         查询 主生产系统的副产品 作为能源系统的原材料的


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr1.Open;

    RB2.NrOfRows := Aqr1.RecordCount;



    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '原材料' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '副产品或固体废弃物' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;

    FlatProgressBar1.Position := 18;
    listbox1.Clear;
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = true then
        if listbox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
          listbox1.Items.Add(Aqr4.FieldByName('pd_name').AsString);
    end;

    RB2.NrOfColumns := listbox1.Items.Count;

    for i := 1 to RB2.NrOfColumns do
      for j := 1 to RB2.NrOfRows do
        RB2.Elem[i, j] := 0; //  elem[列，行]

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr1.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);

        Aqr4.Open;
        RB2.Elem[j + 1, i] := Aqr4.FieldByName('co_num').AsFloat;
      end;
    end; //RB2

    RB2.Transpose;

    RB2_WQ.NrOfColumns := RB2.NrOfColumns;
    RB2_WQ.NrOfRows := RB2.NrOfRows;


    RB2.Multiply(I_PB2, RB2_WQ);


    gBCL.NrOfRows := listbox1.Items.Count;

    FlatProgressBar1.Position := 20;


    Aqr4.Connection := ADOConnection1;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null ');
    Aqr4.Open;



    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from lcifactor order by inf_name ASC');
    Aqr2.Open; //     if Aqr1.Locate('LCIfactor',Aqr2.FieldByName('inf_name').AsString ,[]) then

    gBCL.NrOfColumns := Aqr2.RecordCount;

    for i := 0 to listbox1.Items.Count - 1 do
    begin
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr4.FieldByName('pd_name').AsString = listbox1.Items[i] then
        begin
          if Aqr1.Active then Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39 + ' order by LCIfactor ASC');
          Aqr1.Open;

          for k := 1 to Aqr2.RecordCount do
          begin
            Aqr2.RecNo := k;
            if Aqr1.Locate('LCIfactor', Aqr2.FieldByName('inf_name').AsString, []) then
              gBCL.Elem[k, i + 1] := Aqr1.FieldByName('sum').AsFloat;
          end;

        end;
      end;
    end;

    RB2_WQ.Transpose;


    g3_e.NrOfRows := RB2_WQ.NrOfRows;
    g3_e.NrOfColumns := gBCL.NrOfColumns;
    RB2_WQ.Multiply(gBCL, g3_e);
    g3_e.Transpose;



    FlatProgressBar1.Position := 30;








 //*****************************************************************************

   //g4  副产品内部收益g4       副产品产出系数


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr1.Open;

    RB.NrOfColumns := Aqr1.RecordCount;
    g4_e.NrOfRows := Aqr1.RecordCount;

    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
    Aqr4.Open; //   showmessage(  Aqr4.SQL.Text  +  '    '+inttostr(Aqr4.RecordCount) );

    listbox1.Clear; //分配表没有设置工序类别，所以这里辨识一下。
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
        if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
          listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
    end;
 //      showmessage( listbox1.Items[0]);

    //判断副产品是内部使用的   某工序产出的副产品，在其它工序作为输入；     名称不一致怎么办？

    for i := listbox1.Items.Count - 1 downto 0 do
    begin
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + listbox1.Items[i] + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;
      if Aqr4.RecordCount = 0 then
        listbox1.Items.Delete(i);
    end;

    FlatProgressBar1.Position := 35;

    RB.NrOfRows := listbox1.Items.Count;

    for i := 1 to RB.NrOfColumns do
      for j := 1 to RB.NrOfRows do
        RB.Elem[i, j] := 0; //  elem[列，行]


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
        Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
        Aqr4.Open;
        RB.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;
      end;
    end; //RB
    FlatProgressBar1.Position := 38;

 // RA*(I-PB2)-1

    RB_WQ.NrOfColumns := RB.NrOfColumns;
    RB_WQ.NrOfRows := RB.NrOfRows;

    RB.Multiply(I_PB2, RB_WQ);

    RB_WQ.Transpose;


 //rg2

    RG2.NrOfRows := listbox1.Items.Count;

    RG2.NrOfColumns := g2_e.NrOfRows;

    Aqr4.Connection := ADOConnection1;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
    Aqr4.Open;

    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from lcifactor order by inf_name ASC');
    Aqr3.Open; //     if Aqr1.Locate('LCIfactor',Aqr2.FieldByName('inf_name').AsString ,[]) then

    for j := 0 to listbox1.Items.Count - 1 do
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
      Aqr1.Open;

      if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then // ccc
      begin
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39 + ' order by LCIfactor ASC');
        Aqr2.Open;
        for k := 1 to Aqr3.RecordCount do
        begin
          Aqr3.RecNo := k;
          if Aqr2.Locate('LCIfactor', Aqr3.FieldByName('inf_name').AsString, []) then
            RG2.Elem[k, j + 1] := Aqr2.FieldByName('sum').AsFloat;
        end;
      end;
    end;

    FlatProgressBar1.Position := 40;

    g4_e.NrOfColumns := g2_e.NrOfRows;


   //    （RA*(I-PB2)-1）T*RG


    Rb_WQ.Multiply(RG2, g4_e); //g4


    g4_e.Transpose;
    for i := 1 to g4_e.NrOfColumns do
      for j := 1 to g4_e.NrOfRows do
        g4_e.Elem[i, j] := (-1) * g4_e.Elem[i, j];

    FlatProgressBar1.Position := 45;

//******************************************************************************

    g_onsite_e.NrOfColumns := g3_e.NrOfColumns;
    g_onsite_e.NrOfRows := g3_e.NrOfRows;
    g1_e.Transpose;



    for i := 1 to g_onsite_e.NrOfColumns do
      for j := 1 to g_onsite_e.NrOfRows do
        g_onsite_e.Elem[i, j] := g1_e.Elem[i, j] + g2_e.Elem[i, j] + g3_e.Elem[i, j] + g4_e.Elem[i, j];

    FlatProgressBar1.Position := 50;
//********************************************************************************
   //

      //g5  运输     工序名称 对照表 运输产品名称

        //外购产品带入运输负荷  ＋  能源产品带入运输负荷



    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr1.Open; //所有能源工序


    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '能源公辅工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
    Aqr4.Open;
    juli_e.NrOfColumns := 4;
    juli_e.NrOfRows := Aqr4.RecordCount;
    for i := 1 to juli_e.NrOfColumns do
      for j := 1 to juli_e.NrOfRows do
        juli_e.Elem[i, j] := 0;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from transport ');
    Aqr2.Open;

    Aqr3.Connection := ADOConnection1;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from comparison where transport is not null ');
    Aqr3.open;

    ListBox1.Clear;
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr3.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = true then
      begin
        if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
          ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString); //在对照表中能找到的数据，存入listbox
        if Aqr2.Locate('product', Aqr3.FieldByName('transport').AsString, []) then
          for j := 1 to 4 do
            juli_e.Elem[j, i] := Aqr2.Fields[j + 2].AsFloat;
      end;

    end; // m* 4

    FlatProgressBar1.Position := 55;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from transport_base ');
    Aqr2.Open;

    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCIfactor order by inf_name ASC');
    Aqr3.Open;

    T_base.NrOfRows := 4;
    T_base.NrOfColumns := Aqr3.RecordCount; //4种运输方式的LCI清单
    for i := 1 to T_base.NrOfColumns do
      for j := 1 to T_base.NrOfRows do
        T_base.Elem[i, j] := 0;


    for i := 1 to Aqr3.RecordCount do
    begin
      Aqr3.RecNo := i;
      if Aqr2.Locate('shuxing', Aqr3.FieldByName('inf_name').AsString, []) then
        for j := 1 to 4 do
          T_base.Elem[i, j] := Aqr2.Fields[j + 2].AsFloat;
    end; // 4* 53

    T_km_e.NrOfColumns := T_base.NrOfColumns;
    T_km_e.NrOfRows := juli_e.NrOfRows;

    juli_e.Multiply(T_base, T_km_e); // T_km  外购产品距离 带入的负荷 还需乘以运输量     m*53

    FlatProgressBar1.Position := 60;

  // 对有运输数据的产品的单耗    OF

    OF_2.NrOfColumns := Aqr1.RecordCount;
    OF_2.NrOfRows := juli_e.NrOfRows;
    for i := 1 to OF_2.NrOfColumns do
      for j := 1 to OF_2.NrOfRows do
        OF_2.Elem[i, j] := 0;


    for i := 1 to Aqr1.RecordCount do //每个工序
    begin
      Aqr1.RecNo := i;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype = ' + #39 + '能源公辅工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;

      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr4.FieldByName('process').AsString = Aqr1.FieldByName('process').AsString then
          OF_2.Elem[i, j] := Aqr4.FieldByName('co_num').AsFloat;
      end;
    end;
    OF_2.Transpose;

    T_e.NrOfRows := OF_2.NrOfRows;
    T_e.NrOfColumns := T_km_e.NrOfColumns;
    OF_2.Multiply(T_km_e, T_e); // T 各工序直接运输负荷





    g5_e.NrOfRows := T_e.NrOfRows;
    g5_e.NrOfColumns := T_e.NrOfColumns;
    I_PB2.Transpose;
    I_PB2.Multiply(T_e, g5_e);



    g5_e.Transpose;

    for i := 1 to g5_e.NrOfColumns do
      for j := 1 to g5_e.NrOfRows do
        g5_e.Elem[i, j] := g5_e.Elem[i, j] / 1000; //单位换算  t.km  --  kg

    FlatProgressBar1.Position := 65;
//********************************************************************************

    // g6_e  上游      OB ：外购产品  查询所有产品、副产品，不在这里面的全部属于外购产品


    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '能源公辅工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
    Aqr4.Open;



    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' or material_type=' + #39 + '副产品或固体废弃物' + #39);
    Aqr1.Open;

    ListBox1.Clear;
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
        if Aqr1.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = false then //非产品、副产品，为外购产品
          ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString);
    end;

    OB.NrOfColumns := g_onsite_e.NrOfColumns;
    OB.NrOfRows := ListBox1.Items.Count;


    FlatProgressBar1.Position := 70;

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '能源公辅工序' + #39 + ' and material_type = ' + #39 + '产品' + #39);
    Aqr4.Open;

    for j := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := j;
      for i := 0 to ListBox1.Items.Count - 1 do
      begin
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr4.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + ListBox1.Items[i] + #39);
        Aqr1.Open;
        if Aqr1.RecordCount = 0 then
          OB.Elem[j, i + 1] := 0
        else
          OB.Elem[j, i + 1] := Aqr1.FieldByName('co_num').AsFloat;
      end;

    end;
    FlatProgressBar1.Position := 75;

   //                   OB(I-PB2)-1            L


    OB_WQ.NrOfColumns := OB.NrOfColumns;
    OB_WQ.NrOfRows := OB.NrOfRows;
    I_PB2.Transpose;
    OB.Multiply(I_PB2, OB_WQ);
    OB_WQ.Transpose;


    OG.NrOfRows := listbox1.Items.Count;
    Aqr4.Connection := ADOConnection1;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null ');
    Aqr4.Open;



    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from lcifactor order by inf_name ASC ');
    Aqr3.Open; //     if Aqr1.Locate('LCIfactor',Aqr2.FieldByName('inf_name').AsString ,[]) then
    OG.NrOfColumns := Aqr3.RecordCount;

    for i := 1 to OG.NrOfColumns do
      for j := 1 to OG.NrOfRows do
        OG.Elem[i, j] := 0;

    for i := 0 to listbox1.Items.Count - 1 do
    begin
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr4.FieldByName('pd_name').AsString = listbox1.Items[i] then
        begin
          if Aqr1.Active then Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39 + ' order by LCIfactor ASC');
          Aqr1.Open;
          for k := 1 to Aqr3.RecordCount do
          begin
            Aqr3.RecNo := k;
            if Aqr1.Locate('LCIfactor', Aqr3.FieldByName('inf_name').AsString, []) then
              OG.Elem[k, i + 1] := Aqr1.FieldByName('sum').AsFloat;
          end;

        end;
      end;
    end;



    g6_e.NrOfColumns := g_onsite_e.NrOfRows;
    g6_e.NrOfRows := g_onsite_e.NrOfColumns;

    OB_WQ.Multiply(OG, g6_e); //    (OB(I-PB2)-1)TOG
    g6_e.Transpose;

    FlatProgressBar1.Position := 80;
 //********************************************************************************
      //g7


    //g7

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr1.Open;

    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
    Aqr4.Open;


    listbox1.Clear;
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
      begin
        if Aqr3.Active then Aqr3.Close;
        Aqr3.SQL.Clear;
        Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('albefore').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
        Aqr3.Open;
        if Aqr3.RecordCount = 0 then
        begin
          if Aqr3.Active then Aqr3.Close;
          Aqr3.SQL.Clear;
          Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('alafter').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
          Aqr3.Open;
          if Aqr3.RecordCount = 0 then
            if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
              listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
        end;
      end;
    end;
        //  TRT电  -- 电网电     水渣－－水泥           listbox1.Items.Delete(i);


    RB_out.NrOfColumns := Aqr1.RecordCount;
    RB_out.NrOfRows := listbox1.Items.Count;
    for i := 1 to RB_out.NrOfColumns do
      for j := 1 to RB_out.NrOfRows do
        RB_out.Elem[i, j] := 0;

    for i := 1 to RA_out.NrOfColumns do
      for j := 1 to RA_out.NrOfRows do
        RB_out.Elem[i, j] := 0; //  elem[列，行]

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
        Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
        Aqr4.Open;
        RB_out.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;

      end;
    end; //RA_out

    FlatProgressBar1.Position := 85;

 // RA_out_WQ*(I-PA1)-1

    RB_out_WQ.NrOfColumns := RB_out.NrOfColumns;
    RB_out_WQ.NrOfRows := RB_out.NrOfRows;

    RB_out.Multiply(I_PB2, RB_out_WQ);

    RB_out_WQ.Transpose;

  //rg




    RG2_out.NrOfRows := listbox1.Items.Count;
    RG2_out.NrOfColumns := g6_e.NrOfRows;
    for i := 1 to RG2_out.NrOfColumns do
      for j := 1 to RG2_out.NrOfRows do
        RG2_out.Elem[i, j] := 0;

    Aqr4.Connection := ADOConnection1;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
    Aqr4.Open;


    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from lcifactor order by inf_name ASC');
    Aqr3.Open; //     if Aqr1.Locate('LCIfactor',Aqr2.FieldByName('inf_name').AsString ,[]) then



    for j := 0 to listbox1.Items.Count - 1 do
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
      Aqr1.Open;

      if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then // ccc
      begin
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39 + ' order by LCIfactor ASC');
        Aqr2.Open;
        for k := 1 to Aqr3.RecordCount do
        begin
          Aqr3.RecNo := k;
          if Aqr2.Locate('LCIfactor', Aqr3.FieldByName('inf_name').AsString, []) then
            RG2_out.Elem[k, j + 1] := Aqr2.FieldByName('sum').AsFloat;
        end;
      end;
    end;

    g7_e.NrOfRows := RB_out_WQ.NrOfRows;
    g7_e.NrOfColumns := RG2_out.NrOfColumns;

    RB_out_WQ.Multiply(RG2_out, g7_e); //    （RA*(I-PA1)-1）T*RG       g7_1


    g7_e.Transpose;


    for i := 1 to g7_e.NrOfColumns do
      for j := 1 to g7_e.NrOfRows do
        g7_e.Elem[i, j] := (-1) * g7_e.Elem[i, j];








    FlatProgressBar1.Position := 90;




//********************************************************************************



    g_outsite_e.NrOfColumns := g_onsite_e.NrOfColumns;
    g_outsite_e.NrOfRows := g_onsite_e.NrOfRows;
    g_route_e.NrOfColumns := g_onsite_e.NrOfColumns;
    g_route_e.NrOfRows := g_onsite_e.NrOfRows;

    for i := 1 to g_onsite_e.NrOfColumns do
      for j := 1 to g_onsite_e.NrOfRows do
      begin
        g_outsite_e.Elem[i, j] := g5_e.Elem[i, j] + g6_e.Elem[i, j] + g7_e.Elem[i, j];
        g_route_e.Elem[i, j] := g_outsite_e.Elem[i, j] + g_onsite_e.Elem[i, j];
      end;
    FlatProgressBar1.Position := 100;


    StatusBar1.Panels[3].Text := '能源LCA结果写入数据库';
    StatusBar1.Refresh;

    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close; //写入数据库
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    FlatProgressBar1.Position := 0;
    FlatProgressBar1.Max := Aqr4.RecordCount;

    for i := 1 to Aqr4.RecordCount do //2个产品
    begin
      FlatProgressBar1.Position := i;
      Aqr4.RecNo := i;
      if Aqr5.Active then Aqr5.Close;
      Aqr5.SQL.Clear;
      Aqr5.SQL.Add('select * from LCA where product=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr5.Open;
      for j := 1 to Aqr5.RecordCount do
      begin
        Aqr5.RecNo := j;
        Aqr5.Edit;
        Aqr5.FieldByName('g2').AsFloat := g2_e.Elem[i, j];
        Aqr5.FieldByName('g3').AsFloat := g3_e.Elem[i, j];
        Aqr5.FieldByName('g4').AsFloat := g4_e.Elem[i, j];
        Aqr5.FieldByName('g_site').AsFloat := g_onsite_e.Elem[i, j];
        Aqr5.FieldByName('g5').AsFloat := g5_e.Elem[i, j];
        Aqr5.FieldByName('g6').AsFloat := g6_e.Elem[i, j];
        Aqr5.FieldByName('g7').AsFloat := g7_e.Elem[i, j];
        Aqr5.FieldByName('g_outsite').AsFloat := g_outsite_e.Elem[i, j];
        Aqr5.FieldByName('g_route').AsFloat := g_route_e.Elem[i, j];
        Aqr5.Post;
      end;
    end;






  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;
    Aqr5.Close;
    Aqr5.Free;
  end;

end;


procedure TForm1.mat_mainprocess;
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4, Aqr5: TAdoQuery;
begin
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;
  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection2;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection2;
  Aqr5 := TAdoQuery.Create(nil);
  Aqr5.Connection := ADOConnection2;

  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '计算主生产工序LCA';
  StatusBar1.Refresh;


  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;
    if Aqr1.RecordCount = 0 then exit;
    FlatProgressBar1.Position := 3;

    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCIfactor ORDER BY inf_name ASC');
    Aqr3.Open;

    g1.NrOfColumns := Aqr3.RecordCount;
    g1.NrOfRows := Aqr1.RecordCount;
    FlatProgressBar1.Position := 5;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      direct(Aqr1.FieldByName('process').AsString);
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCA where  product = ' + #39 + Aqr1.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr3.Open;
      for j := 1 to Aqr3.RecordCount do
      begin
        Aqr3.RecNo := j;
        g1.Elem[j, i] := Aqr3.FieldByName('g1').AsFloat; //  elem[列，行]
      end;
    end;

  //  g1.Transpose;


    FlatProgressBar1.Position := 10;
//*****************************************************************************************





    PA1.NrOfColumns := Aqr1.RecordCount;
    PA1.NrOfRows := Aqr1.RecordCount;
    for i := 1 to Aqr1.RecordCount do
      for j := 1 to Aqr1.RecordCount do
        PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品


//   ID	system_name	processtype	process	material_type	pd_name	pd_unit	pd	co_unit	co_num	remark	elementflow	shuxing	transport	upstream


    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr4.Open;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
        Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
        Aqr2.Open;
        if Aqr2.RecordCount > 0 then
        begin
          Aqr2.First;
          PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
        end;
      end;
    end;

    FlatProgressBar1.Position := 13;



    I_PA1.NrOfColumns := PA1.NrOfColumns;
    I_PA1.NrOfRows := PA1.NrOfRows;



    for i := 1 to PA1.NrOfColumns do // I-PA1
      for j := 1 to PA1.NrOfRows do
        if i = j then
          I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
        else
          I_PA1.Elem[i, j] := -PA1.Elem[i, j];


    I_PA1.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


    FlatProgressBar1.Position := 18;


    I_PA1_I.NrOfColumns := I_PA1.NrOfColumns;
    I_PA1_I.NrOfRows := I_PA1.NrOfRows;

    for i := 1 to I_PA1.NrOfColumns do // (I-PA1)-1-I
      for j := 1 to I_PA1.NrOfRows do
        if i = j then
          I_PA1_I.Elem[i, j] := I_PA1.Elem[i, j] - 1
        else
          I_PA1_I.Elem[i, j] := I_PA1.Elem[i, j];



     // g3  前道工序带入负荷     [(I-PA1)-1-I]Tg1

    g3.NrOfColumns := g1.NrOfColumns;
    g3.NrOfRows := g1.NrOfRows;

    I_PA1_I.Transpose; //  ((I-PB2)-1-I )T
    I_PA1_I.Multiply(g1, g3); //g3计算出来了

//   g1.Transpose;
    g3.Transpose;
    FlatProgressBar1.Position := 20;
// ************************************************************************************


     // PA2     各主工序对自产能源的直接消耗系数矩阵

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;

    PA2.NrOfColumns := Aqr1.RecordCount;
    PA2.NrOfRows := Aqr4.RecordCount;

    for i := 1 to PA2.NrOfColumns do
      for j := 1 to PA2.NrOfRows do
        PA2.Elem[i, j] := 0;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
        Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
        Aqr2.Open;
        if Aqr2.RecordCount > 0 then
        begin
          Aqr2.First;
          PA2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
        end;
      end;
    end;
    FlatProgressBar1.Position := 25;





   // PA2*（(I-PA1)-1）       对自产能源的完全消耗系数矩阵

    PA2_WQ.NrOfColumns := PA2.NrOfColumns;
    PA2_WQ.NrOfRows := PA2.NrOfRows;

    PA2.Multiply(I_PA1, PA2_WQ);
    PA2_WQ.Transpose;




    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    PE2.NrOfRows := Aqr4.RecordCount;
    PE2.NrOfColumns := g3.NrOfRows;


    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr1.Open;

      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        PE2.Elem[j, i] := Aqr1.FieldByName('g_site').AsFloat;
      end;
    end;



    g2.NrOfColumns := PE2.NrOfColumns;
    g2.NrOfRows := PA2_WQ.NrOfRows;
    PA2_WQ.Multiply(PE2, g2);


    FlatProgressBar1.Position := 30;
 // ************************************************************************************


     //g4  副产品内部收益g4       副产品产出系数


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;

    RA.NrOfColumns := Aqr1.RecordCount;
    g4.NrOfRows := Aqr1.RecordCount;


    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
    Aqr4.Open;

    listbox1.Clear; //分配表没有设置工序类别，所以这里辨识一下。
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
      begin
        if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
          listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
        if listbox1.Items.IndexOf(Aqr4.FieldByName('alafter').AsString) = -1 then
          listbox1.Items.Add(Aqr4.FieldByName('alafter').AsString);
      end; //把分配前的名称和分配后的名称，都写入到listbox中
    end;


 // 判断副产品是内部使用的   某工序产出的副产品，在其它工序作为输入；     名称不一致怎么办？

    for i := listbox1.Items.Count - 1 downto 0 do
    begin
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + listbox1.Items[i] + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;
      if Aqr4.RecordCount = 0 then
        listbox1.Items.Delete(i);
    end; //删掉不在内部使用的物料



    for i := listbox1.Items.Count - 1 downto 0 do
    begin
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
      Aqr4.Open;
      if Aqr4.Locate('alafter', listbox1.Items[i], []) then
        if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) <> -1 then
          if Aqr4.FieldByName('albefore').AsString <> Aqr4.FieldByName('alafter').AsString then
            Listbox1.Items.Delete(i);
    end; //删掉



    FlatProgressBar1.Position := 33;


    RA.NrOfRows := listbox1.Items.Count;

    for i := 1 to RA.NrOfColumns do
      for j := 1 to RA.NrOfRows do
        RA.Elem[i, j] := 0; //  elem[列，行]


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
        Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
        Aqr4.Open;
        if Aqr4.RecordCount = 0 then //分配后的名称在其它工序使用的情况
        begin
          Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and alafter=' + #39 + listbox1.Items[j] + #39);
          Aqr2.SQL.Add(' and process = ' + #39 + Aqr1.FieldByName('process').AsString + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr4.Close;
            Aqr4.SQL.Clear;
            Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr2.FieldByName('albefore').AsString + #39);
            Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);

            Aqr4.Open;
          end;
        end;

        RA.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;

      end;
    end; //RA



  // RA*(I-PA1)-1

    RA_WQ.NrOfColumns := RA.NrOfColumns;
    RA_WQ.NrOfRows := RA.NrOfRows;

    RA.Multiply(I_PA1, RA_WQ);

    RA_WQ.Transpose;

    FlatProgressBar1.Position := 36;
 //rg

    RG.NrOfRows := listbox1.Items.Count;

    RG.NrOfColumns := g2.NrOfColumns;

    Aqr4.Connection := ADOConnection1;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
    Aqr4.Open;
   // showmessage( Aqr4.SQL.Text);

    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from lcifactor order by inf_name ASC ');
    Aqr3.Open;

    for j := 0 to listbox1.Items.Count - 1 do
    begin


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
      Aqr1.Open;
       //  showmessage( Aqr1.SQL.Text);

      if Aqr1.RecordCount = 0 then
      begin
        Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and alafter=' + #39 + listbox1.Items[j] + #39);
        Aqr1.Open;
      end;



      if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then //   comparison
      begin
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39 + ' order by LCIfactor ASC');
        Aqr2.Open; //  showmessage( Aqr2.SQL.Text);
        for k := 1 to Aqr3.RecordCount do
        begin
          Aqr3.RecNo := k;
          if Aqr2.Locate('LCIfactor', Aqr3.FieldByName('inf_name').AsString, []) then
            RG.Elem[k, j + 1] := Aqr2.FieldByName('sum').AsFloat;
        end;
      end;


                    // CDQ电 --- 电网电          燃煤电


    end;

    g4.NrOfColumns := g2.NrOfColumns;



    RA_WQ.Multiply(RG, g4); //g4        RA.Multiply(I_PA1, RA_WQ);

    g2.Transpose;
    g4.Transpose;


    for i := 1 to g4.NrOfColumns do
      for j := 1 to g4.NrOfRows do
        g4.Elem[i, j] := (-1) * g4.Elem[i, j];



    FlatProgressBar1.Position := 40;
//********************************************************************************

    g_onsite.NrOfColumns := g4.NrOfColumns;
    g_onsite.NrOfRows := g4.NrOfRows;
    g1.Transpose;



    for i := 1 to g_onsite.NrOfColumns do
      for j := 1 to g_onsite.NrOfRows do
        g_onsite.Elem[i, j] := g1.Elem[i, j] + g2.Elem[i, j] + g3.Elem[i, j] + g4.Elem[i, j];

    FlatProgressBar1.Position := 50;
//********************************************************************************
        //g5  运输     工序名称 对照表 运输产品名称

        //外购产品带入运输负荷  ＋  能源产品带入运输负荷



    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open; //所有主工序


    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
    Aqr4.Open;
    juli.NrOfColumns := 4;
    juli.NrOfRows := Aqr4.RecordCount;
    for i := 1 to juli.NrOfColumns do
      for j := 1 to juli.NrOfRows do
        juli.Elem[i, j] := 0;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from transport ');
    Aqr2.Open;

    Aqr3.Connection := ADOConnection1;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from comparison where transport is not null ');
    Aqr3.open;

    ListBox1.Clear;
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr3.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = true then
      begin
        if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
          ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString); //在对照表中能找到的数据，存入listbox
        if Aqr2.Locate('product', Aqr3.FieldByName('transport').AsString, []) then
          for j := 1 to 4 do
            juli.Elem[j, i] := Aqr2.Fields[j + 2].AsFloat;
      end;
    end; // m* 4
    FlatProgressBar1.Position := 55;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from transport_base ');
    Aqr2.Open;

    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCIfactor order by inf_name ASC');
    Aqr3.Open;


    T_base.NrOfRows := 4;
    T_base.NrOfColumns := Aqr3.RecordCount; //4种运输方式的LCI清单
    for i := 1 to T_base.NrOfColumns do
      for j := 1 to T_base.NrOfRows do
        T_base.Elem[i, j] := 0;

    for i := 1 to Aqr3.RecordCount do
    begin
      Aqr3.RecNo := i;
      if Aqr2.Locate('shuxing', Aqr3.FieldByName('inf_name').AsString, []) then
        for j := 1 to 4 do
          T_base.Elem[i, j] := Aqr2.Fields[j + 2].AsFloat;
    end; // 4* 53

    T_km.NrOfColumns := T_base.NrOfColumns;
    T_km.NrOfRows := juli.NrOfRows;

    juli.Multiply(T_base, T_km); // T_km  外购产品距离 带入的负荷 还需乘以运输量     m*53

  // 对有运输数据的产品的单耗    OF

    OF_1.NrOfColumns := Aqr1.RecordCount;
    OF_1.NrOfRows := juli.NrOfRows;
    for i := 1 to OF_1.NrOfColumns do
      for j := 1 to OF_1.NrOfRows do
        OF_1.Elem[i, j] := 0;


    for i := 1 to Aqr1.RecordCount do //每个工序
    begin
      Aqr1.RecNo := i;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype = ' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;

      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr4.FieldByName('process').AsString = Aqr1.FieldByName('process').AsString then
          OF_1.Elem[i, j] := Aqr4.FieldByName('co_num').AsFloat;
      end;
    end;
    OF_1.Transpose;

    T.NrOfRows := OF_1.NrOfRows;
    T.NrOfColumns := T_km.NrOfColumns;
    OF_1.Multiply(T_km, T); // T 各工序直接运输负荷




    g5_1.NrOfRows := T.NrOfRows; //外购产品带入运输负荷，  还需加上能源带入运输负荷
    g5_1.NrOfColumns := T.NrOfColumns;
    I_PA1.Transpose;


    I_PA1.Multiply(T, g5_1);

    for i := 1 to g5_1.NrOfColumns do
      for j := 1 to g5_1.NrOfRows do
        g5_1.Elem[i, j] := g5_1.Elem[i, j] / 1000;




    FlatProgressBar1.Position := 57;








   //能源使用 带入的运输负荷

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    PF2.NrOfRows := Aqr4.RecordCount;
    PF2.NrOfColumns := g5_1.NrOfColumns; //能源运输LCI结果

    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr1.Open;

      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        PF2.Elem[j, i] := Aqr1.FieldByName('g5').AsFloat;
      end;

    end;

    g5_2.NrOfColumns := g5_1.NrOfColumns;
    g5_2.NrOfRows := g5_1.NrOfRows;

    PA2_WQ.Multiply(PF2, g5_2);

    g5.NrOfColumns := g5_1.NrOfColumns;
    g5.NrOfRows := g5_1.NrOfRows;

    for i := 1 to g5.NrOfColumns do
      for j := 1 to g5.NrOfRows do
        g5.Elem[i, j] := g5_1.Elem[i, j] + g5_2.Elem[i, j];

    FlatProgressBar1.Position := 60;
//********************************************************************************************

//g6


    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
    Aqr4.Open;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
//    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' or material_type=' + #39 + '副产品或固体废弃物' + #39);

    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39);

    Aqr1.Open;


    ListBox1.Clear;
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
        if Aqr1.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = false then //非产品、副产品，为外购产品
          ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString);
    end;

    FlatProgressBar1.Position := 63;
    OA.NrOfColumns := g_onsite.NrOfColumns;
    OA.NrOfRows := ListBox1.Items.Count;

    for i := 1 to OA.NrOfColumns do
      for j := 1 to OA.NrOfRows do
        OA.Elem[i, j] := 0;

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and material_type = ' + #39 + '产品' + #39);
    Aqr4.Open;

    for j := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := j;
      for i := 0 to ListBox1.Items.Count - 1 do
      begin
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr4.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + ListBox1.Items[i] + #39);
        Aqr1.Open;
        if Aqr1.RecordCount = 0 then
          OA.Elem[j, i + 1] := 0
        else
          OA.Elem[j, i + 1] := Aqr1.FieldByName('co_num').AsFloat;
      end;

    end;


 // OA(I-PA1)-1            L

    I_PA1.Transpose;

    OA_WQ.NrOfColumns := OA.NrOfColumns;
    OA_WQ.NrOfRows := OA.NrOfRows;

    OA.Multiply(I_PA1, OA_WQ);
    OA_WQ.Transpose;
    FlatProgressBar1.Position := 65;

    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCIfactor order by inf_name  ASC ');
    Aqr3.Open;
    OG.NrOfColumns := Aqr3.RecordCount;

    OG.NrOfRows := listbox1.Items.Count;

    for i := 1 to OG.NrOfColumns do
      for j := 1 to OG.NrOfRows do
        OG.Elem[i, j] := 0;

    Aqr4.Connection := ADOConnection1;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null ');
    Aqr4.Open;

    for i := 0 to listbox1.Items.Count - 1 do
    begin
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr4.FieldByName('pd_name').AsString = listbox1.Items[i] then
        begin
          if Aqr1.Active then Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39 + ' order by LCIfactor ASC');
          Aqr1.Open;

          for k := 1 to Aqr3.RecordCount do
          begin
            Aqr3.RecNo := k;
            if Aqr1.Locate('LCIfactor', Aqr3.FieldByName('inf_name').AsString, []) = true then
              OG.Elem[k, i + 1] := Aqr1.FieldByName('sum').AsFloat; //上游lci

          end;

        end;
      end;
    end; //      exit;

    g6_1.NrOfColumns := OG.NrOfColumns;
    g6_1.NrOfRows := OA_WQ.NrOfRows;

    OA_WQ.Multiply(OG, g6_1); //    (OA(I-PA2)-1)TOG
 //   g6_1.Transpose;





     //能源使用 带入的上游负荷
    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    PG2.NrOfRows := Aqr4.RecordCount;
    PG2.NrOfColumns := g6_1.NrOfColumns; //能源上游LCI结果

    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr1.Open;

      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        PG2.Elem[j, i] := Aqr1.FieldByName('g6').AsFloat;
      end;

    end;

    g6_2.NrOfColumns := PG2.NrOfColumns;
    g6_2.NrOfRows := PA2_WQ.NrOfRows;

    PA2_WQ.Multiply(PG2, g6_2);

    g6.NrOfColumns := g6_1.NrOfColumns;
    g6.NrOfRows := g6_1.NrOfRows;
    for i := 1 to g6.NrOfColumns do
      for j := 1 to g6.NrOfRows do
        g6.Elem[i, j] := g6_1.Elem[i, j] + g6_2.Elem[i, j];





    g6.Transpose;
    FlatProgressBar1.Position := 70;
//********************************************************************************
      //g7

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
    Aqr4.Open;



    listbox1.Clear;
    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
      begin
        if Aqr3.Active then Aqr3.Close;
        Aqr3.SQL.Clear;
        Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('albefore').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
        Aqr3.Open;
        if Aqr3.RecordCount = 0 then
        begin
          if Aqr3.Active then Aqr3.Close;
          Aqr3.SQL.Clear;
          Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('alafter').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
          Aqr3.Open;
          if Aqr3.RecordCount = 0 then
            if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
              listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
        end;
      end;
    end;
        //  TRT电  -- 电网电     水渣－－水泥           listbox1.Items.Delete(i);






    for i := listbox1.Items.Count - 1 downto 0 do
    begin
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[i] + #39);
      Aqr4.Open;

      Aqr3.Connection := ADOConnection2;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('alafter').AsString + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr3.Open;

      if Aqr3.RecordCount > 0 then
        if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) <> -1 then
//    if Aqr4.FieldByName('albefore').AsString <> Aqr4.FieldByName('alafter').AsString then
          Listbox1.Items.Delete(i);

    end;

    FlatProgressBar1.Position := 72;




    RA_out.NrOfColumns := Aqr1.RecordCount;
    RA_out.NrOfRows := listbox1.Items.Count;

    for i := 1 to RA_out.NrOfColumns do
      for j := 1 to RA_out.NrOfRows do
        RA_out.Elem[i, j] := 0; //  elem[列，行]

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 0 to listbox1.Items.Count - 1 do
      begin
        Aqr3.Connection := ADOConnection2;
        if Aqr3.Active then Aqr3.Close;
        Aqr3.SQL.Clear;
        Aqr3.SQL.Add('select * from LCIfactor order by inf_name ASC');
        Aqr3.Open;


        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
        Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
        Aqr4.Open;
        RA_out.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;

      end;
    end; //RA_out

    FlatProgressBar1.Position := 75;

 // RA_out_WQ*(I-PA1)-1

    RA_out_WQ.NrOfColumns := RA_out.NrOfColumns;
    RA_out_WQ.NrOfRows := RA_out.NrOfRows;

    RA_out.Multiply(I_PA1, RA_out_WQ);

    RA_out_WQ.Transpose;

  //rg




    RG_out.NrOfRows := listbox1.Items.Count;
    RG_out.NrOfColumns := g6.NrOfRows;


    Aqr4.Connection := ADOConnection1;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
    Aqr4.Open;

    Aqr3.Connection := ADOConnection2;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCIfactor order by inf_name ASC');
    Aqr3.Open;

    for j := 0 to listbox1.Items.Count - 1 do
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
      Aqr1.Open;

      if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then // ccc
      begin
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39 + ' order by LCIfactor ASC');
        Aqr2.Open;
        for k := 1 to Aqr3.RecordCount do
        begin
          Aqr3.RecNo := k;
          if Aqr2.Locate('LCIfactor', Aqr3.FieldByName('inf_name').AsString, []) = true then
            RG_out.Elem[k, j + 1] := Aqr2.FieldByName('sum').AsFloat;
        end;
      end;
    end;

    g7_1.NrOfRows := RA_out_WQ.NrOfRows;
    g7_1.NrOfColumns := RG_out.NrOfColumns;

    RA_out_WQ.Multiply(RG_out, g7_1); //    （RA*(I-PA1)-1）T*RG       g7_1

    FlatProgressBar1.Position := 78;


      //能源使用 带入副产品回用的收益


    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    PH2.NrOfRows := Aqr4.RecordCount;
    PH2.NrOfColumns := g7_1.NrOfColumns; //能源 g7   LCI结果

    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr1.Open;

      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        PH2.Elem[j, i] := (-1) * Aqr1.FieldByName('g7').AsFloat;
      end;

    end;


    g7_2.NrOfColumns := PH2.NrOfColumns;
    g7_2.NrOfRows := PA2_WQ.NrOfRows;

    PA2_WQ.Multiply(PH2, g7_2);


    g7.NrOfColumns := g7_1.NrOfColumns;
    g7.NrOfRows := g7_1.NrOfRows;
    for i := 1 to g7.NrOfColumns do
      for j := 1 to g7.NrOfRows do
        g7.Elem[i, j] := g7_1.Elem[i, j] + g7_2.Elem[i, j];

    g7.Transpose;
    g5.Transpose;


    for i := 1 to g7.NrOfColumns do
      for j := 1 to g7.NrOfRows do
        g7.Elem[i, j] := (-1) * g7.Elem[i, j];


    FlatProgressBar1.Position := 85;
//********************************************************************************




    g_outsite.NrOfColumns := g_onsite.NrOfColumns;
    g_outsite.NrOfRows := g_onsite.NrOfRows;
    g_route.NrOfColumns := g_onsite.NrOfColumns;
    g_route.NrOfRows := g_onsite.NrOfRows;

    for i := 1 to g_onsite.NrOfColumns do
      for j := 1 to g_onsite.NrOfRows do
      begin
        g_outsite.Elem[i, j] := g5.Elem[i, j] + g6.Elem[i, j] + g7.Elem[i, j];
        g_route.Elem[i, j] := g_outsite.Elem[i, j] + g_onsite.Elem[i, j];
      end;
    FlatProgressBar1.Position := 100;
//********************************************************************************

    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close; //写入数据库
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr4.Open;


    StatusBar1.Panels[3].Text := '主产品LCA结果写入数据库';
    StatusBar1.Refresh;

    FlatProgressBar1.Position := 0;
    FlatProgressBar1.Max := Aqr4.RecordCount;






    for i := 1 to Aqr4.RecordCount do //2个产品
    begin
      Aqr4.RecNo := i;
      FlatProgressBar1.Position := i;
      if Aqr5.Active then Aqr5.Close;
      Aqr5.SQL.Clear;
      Aqr5.SQL.Add('select * from LCA where product=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' order by LCIfactor ASC');
      Aqr5.Open;
      for j := 1 to Aqr5.RecordCount do
      begin
        Aqr5.RecNo := j;
        Aqr5.Edit;
        Aqr5.FieldByName('g2').AsFloat := g2.Elem[i, j];
        Aqr5.FieldByName('g3').AsFloat := g3.Elem[i, j];
        Aqr5.FieldByName('g4').AsFloat := g4.Elem[i, j];
        Aqr5.FieldByName('g_site').AsFloat := g_onsite.Elem[i, j];
        Aqr5.FieldByName('g5').AsFloat := g5.Elem[i, j];
        Aqr5.FieldByName('g6').AsFloat := g6.Elem[i, j];
        Aqr5.FieldByName('g7').AsFloat := g7.Elem[i, j];
        Aqr5.FieldByName('g_outsite').AsFloat := g_outsite.Elem[i, j];
        Aqr5.FieldByName('g_route').AsFloat := g_route.Elem[i, j];
        Aqr5.Post;
      end;
    end;


//********************************************************************************




  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;
    Aqr5.Close;
    Aqr5.Free;
  end;



end;



procedure TForm1.bsSkinButton36Click(Sender: TObject);
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;

const
  BPEI_LCI: array[0..6] of string = ('(r)铁矿石', '(r)水', '(e)总一次能源', '(a)粉尘', '(a)硫氧化物(二氧化硫)', '(a)氮氧化物(二氧化氮)', '(w)化学需氧量');
  BPEI_LCI_data: array[0..6] of double = (0.15, 0.09, 0.27, 0.14, 0.13, 0.1, 0.13);

begin
  bsSkinButton23Click(Sender);


  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;
  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection2;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection2;



  try

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('delete * from LCA ');
    Aqr1.ExecSQL;

    mat_energy;
    mat_mainprocess;

   { if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from LCA ');
    Aqr1.Open;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('LCItype').AsString = '大气污染物' then
        if Aqr1.FieldByName('LCIfactor').AsString <> '(a)二氧化碳' then
        begin
          Aqr1.Edit;
          for j := 5 to 9 do
            Aqr1.Fields[j].AsFloat := Aqr1.Fields[j].AsFloat * 1000;
          Aqr1.Post;
        end;

      if Aqr1.FieldByName('LCItype').AsString = '水体污染物' then
      begin
        Aqr1.Edit;
        for j := 5 to 9 do
          Aqr1.Fields[j].AsFloat := Aqr1.Fields[j].AsFloat * 1000000;
        Aqr1.Post;
      end;



    end;    }







    StatusBar1.Panels[3].Text := '正在进行LCIA计算';
    StatusBar1.Refresh;

    FlatProgressBar1.Position := 0;
    FlatProgressBar1.Max := 100;


    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select distinct impact_type, func_unit from char_coef ');
    Aqr2.Open;


    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCIfactor order by inf_name ASC');
    Aqr3.Open;
    lci.NrOfRows := Aqr3.RecordCount;
    FlatProgressBar1.Position := 10;

    char.NrOfColumns := Aqr3.RecordCount;
    char.NrOfRows := Aqr2.RecordCount;

    for i := 1 to Aqr2.RecordCount do
    begin
      Aqr2.RecNo := i; //lcia type
      for j := 1 to Aqr3.RecordCount do
      begin
        Aqr3.RecNo := j; //lci
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from char_coef where impact_type=' + #39 + Aqr2.FieldByName('impact_type').AsString + #39 + ' and LCI_factor=' + #39 + Aqr3.FieldByName('inf_name').AsString + #39);
        Aqr4.Open;
        if Aqr4.RecordCount > 0 then
          char.Elem[j, i] := Aqr4.FieldByName('equivalent').AsFloat
        else
          char.Elem[j, i] := 0;
      end;
    end;
    FlatProgressBar1.Position := 30;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct product from LCA ');
    Aqr1.Open;
    LCC.NrOfRows := Aqr1.RecordCount;

    for i := 1 to lcc.NrOfColumns do
      for j := 1 to lcc.NrOfRows do
        LCC.Elem[i, j] := 0;


    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select distinct impact_type, func_unit from char_coef ');
    Aqr4.Open;
    FlatProgressBar1.Position := 50;
    for i := 1 to Aqr1.RecordCount do // 产品
    begin
      Aqr1.RecNo := i;
      if i < 50 then
        FlatProgressBar1.Position := 50 + i;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from LCA where product = ' + #39 + Aqr1.FieldByName('product').AsString + #39 + ' order by LCIfactor ASC');
      Aqr2.Open;
      for j := 1 to Aqr3.RecordCount do
      begin
        Aqr3.RecNo := j;
        if Aqr2.Locate('LCIfactor', Aqr3.FieldByName('inf_name').AsString, []) = true then
          for k := 1 to 10 do
          begin
            lci.Elem[k, j] := Aqr2.Fields[k + 4].AsFloat; //  53行 * 10 列
            lcc.Elem[k, i] := lcc.Elem[k, i] + lci.Elem[k, j] * Aqr3.FieldByName('price').AsFloat;
          end;
      end;


      lcia.NrOfColumns := lci.NrOfColumns;
      lcia.NrOfRows := char.NrOfRows;
      char.Multiply(lci, lcia); // lcia 8行  * 10 列

      for j := 1 to lcia.NrOfRows do
      begin
        Aqr4.RecNo := j;
        Aqr2.Append;
        Aqr2.FieldByName('product').AsString := Aqr1.FieldByName('product').AsString;
        Aqr2.FieldByName('LCItype').AsString := 'LCIA';
        Aqr2.FieldByName('LCIfactor').AsString := Aqr4.FieldByName('impact_type').AsString;
        Aqr2.FieldByName('unit').AsString := Aqr4.FieldByName('func_unit').AsString;
        for k := 1 to 10 do
          Aqr2.Fields[k + 4].AsFloat := lcia.Elem[k, j];
        Aqr2.Post;
      end;



      Aqr2.Append;
      Aqr2.FieldByName('product').AsString := Aqr1.FieldByName('product').AsString;
      Aqr2.FieldByName('LCItype').AsString := 'LCC';
      Aqr2.FieldByName('LCIfactor').AsString := '生命周期成本';
      Aqr2.FieldByName('unit').AsString := '元';
      for k := 1 to 10 do
        Aqr2.Fields[k + 4].AsFloat := lcc.Elem[k, i];
      Aqr2.Post;


    end;





    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from scrapeLCA where inf_name=' + #39 + 'S' + #39);
    Aqr2.open;
    if Aqr2.RecordCount > 0 then
    begin
      for i := 1 to Aqr2.RecordCount do
      begin
        Aqr2.RecNo := i;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from data_allocation where process= ' + #39 + Aqr2.FieldByName('inf').AsString + #39 + ' and material_type=' + #39 + '原材料' + #39 + ' and (pd_name=' + #39 + '废钢' + #39 + ' or pd_name=' + #39 + '处理后废钢' + #39 + ')');
        Aqr1.Open;
        if Aqr1.RecordCount > 0 then
        begin
          Aqr2.Edit;
          Aqr2.FieldByName('num').AsFloat := Aqr1.FieldByName('co_num').AsFloat;
          Aqr2.Post;
        end;
      end;
    end;




    StatusBar1.Panels[3].Text := '正在进行BPEI计算';
    StatusBar1.Refresh;

    FlatProgressBar1.Position := 0;


    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from BPEIbase ');
    Aqr4.Open;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct product from LCA');
    Aqr1.Open;
    BPEI.NrOfRows := Aqr1.RecordCount;
    FlatProgressBar1.Max := Aqr1.RecordCount;
    for i := 1 to 2 do
      for j := 1 to BPEI.NrOfRows do
        BPEI.Elem[i, j] := 0;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from  LCA ');
    Aqr2.open;




    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      FlatProgressBar1.Position := i;
      for j := 0 to 6 do
      begin
        if Aqr2.Locate('product; LCIfactor ', VarArrayOf([Aqr1.FieldByName('product').AsString, BPEI_LCI[j]]), []) then
          if Aqr4.Locate('product; LCIfactor ', VarArrayOf([Aqr1.FieldByName('product').AsString, BPEI_LCI[j]]), []) then
          begin
            //if Aqr2.FieldByName('g_site').AsFloat <> 0.0 then
            try
              BPEI.Elem[1, i] := BPEI.Elem[1, i] + Aqr2.FieldByName('g_site').AsFloat / Aqr4.FieldByName('g_site').AsFloat * BPEI_LCI_data[j];
            except
              BPEI.Elem[1, i] := BPEI.Elem[1, i];
            end;
          //  if Aqr2.FieldByName('g_route').AsFloat <> 0.0 then

            try
              BPEI.Elem[2, i] := BPEI.Elem[2, i] + Aqr2.FieldByName('g_route').AsFloat / Aqr4.FieldByName('g_route').AsFloat * BPEI_LCI_data[j];
            except
              BPEI.Elem[2, i] := BPEI.Elem[2, i];
            end;


          end;
      end;
      Aqr2.Append;
      Aqr2.FieldByName('product').AsString := Aqr1.FieldByName('product').AsString;
      Aqr2.FieldByName('LCItype').AsString := 'BPEI';
      Aqr2.FieldByName('LCIfactor').AsString := 'basedonLCI';
      Aqr2.FieldByName('unit').AsString := '无量纲';
      Aqr2.FieldByName('g_site').AsFloat := BPEI.Elem[1, i] * 100;
      Aqr2.FieldByName('g_route').AsFloat := BPEI.Elem[2, i] * 100;
      Aqr2.Post;
    end;




    StatusBar1.Panels[3].Text := '正在进行全厂CO2计算';

    StatusBar1.Refresh;
    FlatProgressBar1.Position := 0;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('delete * from CO2emission ');
    Aqr1.ExecSQL;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type =' + #39 + '产品' + #39); //  + ' and processtype=' + #39 +'主生产工序' + #39
    Aqr1.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from CO2emission ');
      Aqr4.Open;

      Aqr4.Append;
      Aqr4.FieldByName('product').AsString := Aqr1.FieldByName('pd_name').AsString;

      Aqr4.FieldByName('chanliang').AsFloat := Aqr1.FieldByName('pd').AsFloat * unit_exchange(Aqr1.FieldByName('pd_unit').AsString);

      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from data_allocation where material_type <>' + #39 + '产品' + #39 + ' and pd_name=' + #39 + Aqr1.FieldByName('pd_name').AsString + #39);
      Aqr3.Open;
      Aqr4.FieldByName('haoliang').AsFloat := 0;
      for j := 1 to Aqr3.RecordCount do
      begin
        Aqr3.RecNo := j;
        if Aqr3.FieldByName('material_type').AsString <> '副产品或固体废弃物' then
          Aqr4.FieldByName('haoliang').AsFloat := Aqr4.FieldByName('haoliang').AsFloat + Aqr3.FieldByName('pd').AsFloat * unit_exchange(Aqr3.FieldByName('pd_unit').AsString)
        else
          Aqr4.FieldByName('haoliang').AsFloat := Aqr4.FieldByName('haoliang').AsFloat - Aqr3.FieldByName('pd').AsFloat * unit_exchange(Aqr3.FieldByName('pd_unit').AsString);
      end;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from  LCA where LCIfactor=' + #39 + '(a)二氧化碳' + #39 + ' and product=' + #39 + Aqr1.FieldByName('pd_name').AsString + #39);
      Aqr2.open;
      if Aqr2.RecordCount > 0 then
        Aqr4.FieldByName('LCIsite').AsFloat := Aqr2.FieldByName('g_site').AsFloat
      else
        Aqr4.FieldByName('LCIsite').AsFloat := 0;


      Aqr4.FieldByName('waimailiang').AsFloat := Aqr4.FieldByName('chanliang').AsFloat - Aqr4.FieldByName('haoliang').AsFloat;


      Aqr4.FieldByName('siteTotal').AsFloat := Aqr4.FieldByName('waimailiang').AsFloat * Aqr4.FieldByName('LCIsite').AsFloat / 1000000;



      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where pd_name =' + #39 + '电网电' + #39 + ' and process=' + #39 + Aqr1.FieldByName('process').AsString + #39);
      Aqr2.open;
      if Aqr2.RecordCount > 0 then
        Aqr4.FieldByName('gridElectricity').AsFloat := Aqr2.FieldByName('pd').AsFloat * unit_exchange(Aqr2.FieldByName('pd_unit').AsString) / 10000
      else
        Aqr4.FieldByName('gridElectricity').AsFloat := 0;

      Aqr4.FieldByName('gridElectricityCO2').AsFloat := Aqr4.FieldByName('gridElectricity').AsFloat * 788 / 1000000 * 10000; //外购电带入
      Aqr4.FieldByName('CO2Total').AsFloat := Aqr4.FieldByName('siteTotal').AsFloat - Aqr4.FieldByName('gridElectricityCO2').AsFloat;

      Aqr4.Post;

    end;


    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select sum(CO2Total) as CO2_Total from CO2emission ');
    Aqr3.Open;
 //   if Aqr3.RecordCount > 0 then
  //    FlatEdit18.Text := formatfloat('###,##0', Aqr3.FieldByName('CO2_Total').AsFloat) + '  吨';









    FlatProgressBar1.Position := 0;
    StatusBar1.Panels[3].Text := '计算完成';
    StatusBar1.Refresh;





    PageControl1.ActivePageIndex := 7;
    PageControl1.Pages[6].Highlighted := false;
    PageControl1.Pages[7].Highlighted := true;
    bsSkinCheckListBox1.Clear;
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39);
    Aqr1.Open;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      bsSkinCheckListBox1.Items.Add(Aqr1.FieldByName('pd_name').AsString);
    end;

    ComboBox5.ItemIndex := 0;
    ComboBox4.ItemIndex := 0;

    ComboBox5Change(Sender);
    ComboBox4Change(Sender);

    CheckBox4.Checked := true;
    CheckBox5.Checked := true;

    CheckBox4Click(Sender);
    RadioButton9.Checked := true;
    RadioButton9Click(Sender);



 //   PageControl1Change(Sender);





  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;
  end;






end;

procedure TForm1.bsSkinButton46Click(Sender: TObject);
var
  rootnode: ttreenode;
  mysql: string;
  i: integer;
begin
  Panel7.Visible := false;
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  TreeView1.Hint := '基础信息管理';
  TreeView1.Items.Clear;
  PageControl1.Visible := true;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := false;
  PageControl1.ActivePageIndex := 11;


  rootnode := TreeView1.items.add(nil, '基础信息管理');
  TreeView1.Items.Addchild(rootnode, '外部LCA数据管理');
  TreeView1.Items.Addchild(rootnode, '运输LCA基础数据管理');
  TreeView1.Items.Addchild(rootnode, 'LCI环境指标管理');

  TreeView1.Items.Addchild(rootnode, '环境影响类型管理');

  rootnode := TreeView1.items.add(nil, '对照表管理');
  TreeView1.Items.Addchild(rootnode, '物料属性名称对照表');
  TreeView1.Items.Addchild(rootnode, '运输名称对照表');
  TreeView1.Items.Addchild(rootnode, '上游数据名称对照表');
  TreeView1.Items.Addchild(rootnode, 'LCI因子名称对照表');
  rootnode := TreeView1.items.add(nil, '工序对标数据管理');
  TreeView1.FullExpand;
  TreeView1.Items[5].Selected := true;
  TreeView1.SetFocus;


  Image1.Visible := false;
  bsSkinPanel2.Visible := false; //功能单位

  ADOQuery3.Connection := Form1.ADOConnection1;
  ADOQuery3.Close;
  ADOQuery3.SQL.Clear;


  ADOQuery4.Connection := Form1.ADOConnection1;
  ADOQuery4.Close;
  ADOQuery4.SQL.Clear;



  if RadioButton5.Checked = true then
  begin
    mysql := 'select ID, pd_name, shuxing from  comparison  where shuxing is not null';
    Label26.Caption := '物料属性数据名称';
    ADOQuery4.SQL.Add('select ID,name from shuxing ');
    TreeView1.Items[6].Selected := true;
  end;

  if RadioButton6.Checked = true then
  begin
    mysql := 'select ID, pd_name, transport from comparison where transport is not null';
    Label26.Caption := '运输数据名称';
    ADOQuery4.SQL.Add('select ID,product from transport');
    TreeView1.Items[7].Selected := true;
  end;

  if RadioButton7.Checked = true then
  begin
    mysql := 'select ID, 	pd_name, upstream from  comparison where upstream is not null';
    Label26.Caption := '上游LCA数据名称';
    ADOQuery4.SQL.Add('select ID,product from up');
    TreeView1.Items[8].Selected := true;
  end;

  if RadioButton8.Checked = true then
  begin
    mysql := 'select ID, 	pd_name,  lcifactor from comparison where lcifactor is not null';
    Label26.Caption := 'LCI因子数据名称';
    ADOQuery4.SQL.Add('select ID, inf_name  from lcifactor');
    TreeView1.Items[9].Selected := true;
  end;


  ADOQuery3.SQL.Add(mysql);

  ADOQuery3.Open;
  ADOQuery4.Open;


  for i := 0 to DBGridEh7.Columns.Count - 1 do
  begin
    DBGridEh7.Columns[i].Width := 140;
    DBGridEh7.Columns[i].Title.Alignment := taCenter;
    DBGridEh7.Columns[i].Alignment := taLeftJustify;
  end;
  DBGridEh7.Columns[0].Width := 50;

  DBGridEh7.Columns[0].Title.caption := '序号';
  DBGridEh7.Columns[1].Title.caption := '工序中的名称';
  DBGridEh7.Columns[2].Title.caption := '数据库中的名称';

  DBGridEh7.Columns.Items[2].PickList.clear;
  // for i:=1 to  ADOQuery4
 // DBGridEh7.Columns.Items[2].picklist.Add('资源类');


  ListBox5.Items.Clear;
  for i := 1 to ADOQuery4.RecordCount do
  begin
    ADOQuery4.RecNo := i;
    if ListBox5.Items.IndexOf(ADOQuery4.Fields[1].AsString) = -1 then
      ListBox5.Items.Add(ADOQuery4.Fields[1].AsString);
  end;






end;

procedure TForm1.Label25Click(Sender: TObject);
var
  i: integer;
begin
  bsSkinPanel2.Visible := false; //功能单位
  Label4.Caption := Label1.Caption + '2';
  for i := 0 to 12 do
    PageControl1.Pages[i].TabVisible := false;

  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := true;

  ListBox4.Clear;
  ListBox4.Items.Add('海轮');
  ListBox4.Items.Add('火车');
  ListBox4.Items.Add('卡车');
  ListBox4.Items.Add('驳船');



  AdoQuery1.Close;
  ListBox4Click(Sender);

end;

procedure TForm1.Label27Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;
begin
  bsSkinPanel2.Visible := false; //功能单位
  Label4.Caption := Label1.Caption + '3';
  for i := 0 to 12 do
    PageControl1.Pages[i].TabVisible := false;

  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := true;
  bsSkinPanel2.Visible := false;


  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct impact_type from char_coef');
    Aqr1.Open;
    ListBox4.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      ListBox4.Items.Add(Aqr1.FieldByName('impact_type').AsString);
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;

  ADOQuery1.Close;
  ListBox4Click(Sender);
end;

procedure TForm1.ADOQuery3AfterOpen(DataSet: TDataSet);
begin
  ADOQuery3.FieldByName('ID').OnGetText := SetFdText3;
end;

procedure TForm1.ADOQuery4AfterOpen(DataSet: TDataSet);
begin
  ADOQuery4.FieldByName('ID').OnGetText := SetFdText4;
end;

procedure TForm1.RadioButton5Click(Sender: TObject);
begin
  bsSkinButton46Click(Sender);
end;

procedure TForm1.RadioButton6Click(Sender: TObject);
begin
  bsSkinButton46Click(Sender);
end;

procedure TForm1.RadioButton7Click(Sender: TObject);
begin
  bsSkinButton46Click(Sender);
end;

procedure TForm1.RadioButton8Click(Sender: TObject);
begin
  bsSkinButton46Click(Sender);
end;

procedure TForm1.bsSkinButton10Click(Sender: TObject);
begin
  ADOQuery3.UpdateBatch(arall);
end;

procedure TForm1.bsSkinButton13Click(Sender: TObject);
var
  i: integer;
begin
  if DBGridEh7.SelectedRows.Count = 0 then
  begin
    bsSkinMessage1.MessageDlg('请选择需要删除的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if bsSkinMessage1.MessageDlg('确定删除所选记录？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := 0 to DBGridEh7.SelectedRows.Count - 1 do
    begin
      DBGridEh7.DataSource.DataSet.GotoBookmark(pointer(DBGridEh7.SelectedRows[i]));
      DBGridEh7.DataSource.DataSet.Delete;
    end;
  end;

  ADOQuery3.UpdateBatch(arall);
end;

procedure TForm1.bsSkinButton18Click(Sender: TObject);
var
  i: integer;
begin
  if ListBox5.Count = 0 then exit;
  if trim(Edit3.Text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请输入查询名称中至少一个字符！', (mtInformation), [mbOK], 0);
    edit2.SetFocus;
    exit;
  end;

  listbox5.ClearSelection;
  for i := 0 to ListBox5.Count - 1 do
    if Pos(Edit3.Text, ListBox5.Items.Strings[i]) > 0 then
    begin
      ListBox5.Selected[i] := true;

    end;



end;

procedure TForm1.bsSkinButton19Click(Sender: TObject);
begin
  if ListBox5.SelCount = 0 then exit;


  if ListBox5.SelCount > 1 then
  begin
    form1.bsSkinMessage1.MessageDlg('名称匹配不能多选！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if DBGridEh7.SelectedRows.Count = 0 then
  begin
    bsSkinMessage1.MessageDlg('请选择一条需要匹配的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if DBGridEh7.SelectedRows.Count > 1 then
  begin
    bsSkinMessage1.MessageDlg('只能选择一条需要匹配的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  ADOQuery3.Edit;
  ADOQuery3.Fields[2].AsString := ListBox5.Items[ListBox5.ItemIndex];
  ADOQuery3.Post;


  ADOQuery3.UpdateBatch(arall);


end;

procedure TForm1.DBGridEh7CellClick(Column: TColumnEh);
begin
  if ADOQuery3.Fields[1].AsString <> null then
    edit3.Text := ADOQuery3.Fields[1].AsString;
end;

procedure TForm1.Label6Click(Sender: TObject);
var
  i, j, k: integer;
  mysql: string;
  Aqr1, Aqr2: TAdoQuery;
  a: double;
begin
  //**********************************************电力构成
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := Form1.ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := Form1.ADOConnection2;

  try

    AdoQuery8.Connection := Form1.ADOConnection2;
    mysql := 'select ID,name,shuxing,danwei,shuzhi from shuxing where shuxing= ' + #39 + '电力构成' + #39;
    if AdoQuery8.Active then AdoQuery8.Close;
    AdoQuery8.SQL.Clear;
    AdoQuery8.SQL.Add(mysql);
    ADOQuery8.Open;

    for j := 1 to ADOQuery8.RecordCount do //先删除旧数据
    begin
      ADOQuery8.RecNo := j;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('delete * from  data_allocation where pd_name= ' + #39 + ADOQuery8.FieldByName('name').AsString + #39 + ' and material_type <> ' + #39 + '产品' + #39);
      Aqr2.ExecSQL;

    end;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('delete * from  data_allocation where pd_name=' + #39 + '电' + #39);
    Aqr1.ExecSQL;


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from  data where pd_name=' + #39 + '电' + #39);
    Aqr1.Open;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from  data_allocation ');
    Aqr2.Open;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      a := 0;
      for j := 1 to ADOQuery8.RecordCount do
      begin
        ADOQuery8.RecNo := j;
        if ADOQuery8.FieldByName('shuxing').AsString = '电力构成' then
        begin
          a := ADOQuery8.FieldByName('shuzhi').AsFloat;


          Aqr2.Append;

          for k := 1 to Aqr1.FieldCount - 1 do
            Aqr2.Fields[k].AsString := Aqr1.Fields[k].AsString;

          Aqr2.FieldByName('pd').AsFloat := a / 100 * Aqr1.FieldByName('pd').AsFloat;
          Aqr2.FieldByName('co_num').AsFloat := a / 100 * Aqr1.FieldByName('co_num').AsFloat;
          Aqr2.FieldByName('pd_name').AsString := ADOQuery8.FieldByName('name').AsString;
          Aqr2.Post;
        end; // '电力构成'

      end; // for j


    end;



    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select distinct process from  data_allocation ');
    Aqr2.Open;

    ComboBox8.Items.Clear;
    ComboBox8.Items.Add('全部工序');
    for i := 1 to Aqr2.RecordCount do
    begin
      Aqr2.RecNo := i;
      ComboBox8.Items.Add(Aqr2.FieldByName('process').AsString);
    end;

  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
  end;

//********************************************************************




  AdoQuery5.Connection := Form1.ADOConnection2;
  if AdoQuery5.Active then AdoQuery5.Close;
  AdoQuery5.SQL.Clear;
  mysql := 'select ID, processtype, process,material_type,pd_name,	pd_unit,	pd,	co_unit,	co_num,	remark from  data_allocation';
  AdoQuery5.SQL.Add(mysql);
  ADOQuery5.Open;
  ADOQuery5.Sort := 'processtype ASC,PROCESS ASC, material_type ASC';


  for i := 0 to DBGridEh6.Columns.Count - 1 do
  begin
    DBGridEh6.Columns[i].Width := 80;
    DBGridEh6.Columns[i].Title.Alignment := taCenter;
  //  DBGridEh6.Columns[i].Alignment := taCenter;
  end;

  DBGridEh6.Columns[6].Alignment := taRightJustify; //数据
  DBGridEh6.Columns[3].Alignment := taLeftJustify;
  DBGridEh6.Columns[4].Alignment := taLeftJustify;

  DBGridEh6.Columns[0].Title.caption := '序号';
  DBGridEh6.Columns[1].Title.caption := '工序类别';
  DBGridEh6.Columns[2].Title.caption := '工序名称';
  DBGridEh6.Columns[3].Title.caption := '数据类别';
  DBGridEh6.Columns[4].Title.caption := '数据名称';
  DBGridEh6.Columns[5].Title.caption := '单位';
  DBGridEh6.Columns[6].Title.caption := '数量';
  DBGridEh6.Columns[7].Title.caption := '单耗单位';
  DBGridEh6.Columns[8].Title.caption := '单耗';
  DBGridEh6.Columns[9].Title.caption := '备注';
  (ADOQuery5.Fields[8] as Tfloatfield).DisplayFormat := '###,###,##0.##0';
  (ADOQuery5.Fields[6] as Tfloatfield).DisplayFormat := '###,###,##0';
  DBGridEh6.Columns[0].Width := 40;
  DBGridEh6.Columns[3].Width := 100;
  DBGridEh6.Columns[4].Width := 120;
  DBGridEh6.Columns[5].Width := 80;
  DBGridEh6.Columns[7].Width := 60;
  DBGridEh6.Columns[9].Width := 160;

end;

procedure TForm1.Label20Click(Sender: TObject);
var
  mysql: string;
  i: integer;
begin
  AdoQuery5.Connection := Form1.ADOConnection2;
  if AdoQuery5.Active then AdoQuery5.Close;
  AdoQuery5.SQL.Clear;
  mysql := 'select ID, processtype, process,material_type,pd_name, shuxing from  data_allocation where shuxing is not null ';
  AdoQuery5.SQL.Add(mysql);
  ADOQuery5.Open;
  ADOQuery5.Sort := 'PROCESS ASC';

  for i := 0 to DBGridEh6.Columns.Count - 1 do
  begin
    DBGridEh6.Columns[i].Title.Alignment := taCenter;
    DBGridEh6.Columns[i].Alignment := taCenter;
  end;
  DBGridEh6.Columns[0].Title.caption := '序号';
  DBGridEh6.Columns[1].Title.caption := '工序类别';
  DBGridEh6.Columns[2].Title.caption := '工序名称';
  DBGridEh6.Columns[3].Title.caption := '数据类别';
  DBGridEh6.Columns[4].Title.caption := '数据名称';
  DBGridEh6.Columns[5].Title.caption := '物料属性数据库中的名称';

  DBGridEh6.Columns[0].Width := 40;
  DBGridEh6.Columns[1].Width := 120;
  DBGridEh6.Columns[2].Width := 140;
  DBGridEh6.Columns[3].Width := 100;
  DBGridEh6.Columns[4].Width := 120;
  DBGridEh6.Columns[5].Width := 180;




end;

procedure TForm1.bsSkinButton24Click(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to 10 do
    PageControl1.Pages[i].TabVisible := true;

  PageControl1.Pages[11].TabVisible := false;
  PageControl1.Pages[12].TabVisible := false;
  PageControl1.ActivePageIndex := 6;
  Label6Click(Sender);
end;

procedure TForm1.bsSkinButton23Click(Sender: TObject);
var
  i: integer;
  Aqr1, Aqr2: TAdoQuery;
  mysql: string;
begin
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection1;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;

  try
    StatusBar1.Panels[3].Text := '正在进行匹配对照表';
    StatusBar1.Refresh;




    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    mysql := 'select * from  data_allocation';
    Aqr2.SQL.Add(mysql);
    Aqr2.Open;
    FlatProgressBar1.Position := 0;
    FlatProgressBar1.Max := Aqr2.RecordCount;


    for i := 1 to Aqr2.RecordCount do
    begin
      FlatProgressBar1.Position := i;
      Aqr2.RecNo := i;
      Aqr2.Edit;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      mysql := 'select * from  comparison where pd_name=' + #39 + Aqr2.FieldByName('pd_name').AsString + #39 + ' and shuxing is not null';
      Aqr1.SQL.Add(mysql);
      Aqr1.Open;
      if Aqr1.RecordCount > 0 then
        Aqr2.FieldByName('shuxing').AsString := Aqr1.FieldByName('shuxing').AsString;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      mysql := 'select * from  comparison where pd_name=' + #39 + Aqr2.FieldByName('pd_name').AsString + #39 + ' and lcifactor is not null';
      Aqr1.SQL.Add(mysql);
      Aqr1.Open;
      if Aqr1.RecordCount > 0 then
        Aqr2.FieldByName('elementflow').AsString := Aqr1.FieldByName('lcifactor').AsString;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      mysql := 'select * from  comparison where pd_name=' + #39 + Aqr2.FieldByName('pd_name').AsString + #39 + ' and transport is not null';
      Aqr1.SQL.Add(mysql);
      Aqr1.Open;
      if Aqr1.RecordCount > 0 then
        Aqr2.FieldByName('transport').AsString := Aqr1.FieldByName('transport').AsString;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      mysql := 'select * from  comparison where pd_name=' + #39 + Aqr2.FieldByName('pd_name').AsString + #39 + ' and upstream is not null';
      Aqr1.SQL.Add(mysql);
      Aqr1.Open;
      if Aqr1.RecordCount > 0 then
        Aqr2.FieldByName('upstream').AsString := Aqr1.FieldByName('upstream').AsString;

      Aqr2.Post;

    end;


    StatusBar1.Panels[3].Text := '数据名称匹配完成';
    StatusBar1.Refresh;
    FlatProgressBar1.Position := 0;
  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
  end;

  flag1 := true;
 // Label20Click(Sender);



end;

procedure TForm1.ADOQuery5AfterOpen(DataSet: TDataSet);
begin
  ADOQuery5.FieldByName('ID').OnGetText := SetFdText5;
end;

procedure TForm1.Label21Click(Sender: TObject);
var
  mysql: string;
  i: integer;
begin
  AdoQuery5.Connection := Form1.ADOConnection2;
  if AdoQuery5.Active then AdoQuery5.Close;
  AdoQuery5.SQL.Clear;
  mysql := 'select ID, processtype, process,material_type,pd_name, transport from  data_allocation where transport is not null ';
  AdoQuery5.SQL.Add(mysql);
  ADOQuery5.Open;
  ADOQuery5.Sort := 'PROCESS ASC';

  for i := 0 to DBGridEh6.Columns.Count - 1 do
  begin
    DBGridEh6.Columns[i].Title.Alignment := taCenter;
    DBGridEh6.Columns[i].Alignment := taCenter;
  end;
  DBGridEh6.Columns[0].Title.caption := '序号';
  DBGridEh6.Columns[1].Title.caption := '工序类别';
  DBGridEh6.Columns[2].Title.caption := '工序名称';
  DBGridEh6.Columns[3].Title.caption := '数据类别';
  DBGridEh6.Columns[4].Title.caption := '数据名称';
  DBGridEh6.Columns[5].Title.caption := '运输数据库中的名称';

  DBGridEh6.Columns[0].Width := 40;
  DBGridEh6.Columns[1].Width := 120;
  DBGridEh6.Columns[2].Width := 140;
  DBGridEh6.Columns[3].Width := 100;
  DBGridEh6.Columns[4].Width := 120;
  DBGridEh6.Columns[5].Width := 180;

end;

procedure TForm1.Label22Click(Sender: TObject);
var
  mysql: string;
  i: integer;
begin
  AdoQuery5.Connection := Form1.ADOConnection2;
  if AdoQuery5.Active then AdoQuery5.Close;
  AdoQuery5.SQL.Clear;
  mysql := 'select ID, processtype, process,material_type,pd_name, upstream from  data_allocation where upstream is not null ';
  AdoQuery5.SQL.Add(mysql);
  ADOQuery5.Open;
  ADOQuery5.Sort := 'PROCESS ASC';

  for i := 0 to DBGridEh6.Columns.Count - 1 do
  begin
    DBGridEh6.Columns[i].Title.Alignment := taCenter;
    DBGridEh6.Columns[i].Alignment := taCenter;
  end;
  DBGridEh6.Columns[0].Title.caption := '序号';
  DBGridEh6.Columns[1].Title.caption := '工序类别';
  DBGridEh6.Columns[2].Title.caption := '工序名称';
  DBGridEh6.Columns[3].Title.caption := '数据类别';
  DBGridEh6.Columns[4].Title.caption := '数据名称';
  DBGridEh6.Columns[5].Title.caption := '上游LCA数据库中的名称';

  DBGridEh6.Columns[0].Width := 40;
  DBGridEh6.Columns[1].Width := 120;
  DBGridEh6.Columns[2].Width := 140;
  DBGridEh6.Columns[3].Width := 100;
  DBGridEh6.Columns[4].Width := 120;
  DBGridEh6.Columns[5].Width := 180;

end;

procedure TForm1.Label23Click(Sender: TObject);
var
  mysql: string;
  i: integer;
begin
  AdoQuery5.Connection := Form1.ADOConnection2;
  if AdoQuery5.Active then AdoQuery5.Close;
  AdoQuery5.SQL.Clear;
  mysql := 'select ID, processtype, process,material_type,pd_name, elementflow from  data_allocation where elementflow is not null ';
  AdoQuery5.SQL.Add(mysql);
  ADOQuery5.Open;
  ADOQuery5.Sort := 'PROCESS ASC';

  for i := 0 to DBGridEh6.Columns.Count - 1 do
  begin
    DBGridEh6.Columns[i].Title.Alignment := taCenter;
    DBGridEh6.Columns[i].Alignment := taCenter;
  end;
  DBGridEh6.Columns[0].Title.caption := '序号';
  DBGridEh6.Columns[1].Title.caption := '工序类别';
  DBGridEh6.Columns[2].Title.caption := '工序名称';
  DBGridEh6.Columns[3].Title.caption := '数据类别';
  DBGridEh6.Columns[4].Title.caption := '数据名称';
  DBGridEh6.Columns[5].Title.caption := 'LCI因子数据库中的名称';

  DBGridEh6.Columns[0].Width := 40;
  DBGridEh6.Columns[1].Width := 120;
  DBGridEh6.Columns[2].Width := 140;
  DBGridEh6.Columns[3].Width := 100;
  DBGridEh6.Columns[4].Width := 120;
  DBGridEh6.Columns[5].Width := 180;

end;

procedure TForm1.Button1Click(Sender: TObject);
//var
// i, j: integer;
 // Aqr1, Aqr2: TAdoQuery;
//  s1, s2, s3, s4: string;
begin
  checkmat(OA_WQ);
  //showmessage(floattostr(energy('热电厂-CCPP')));
 { Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection1;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection1;
  try
    s1 := '';
    s2 := '';
    s3 := '';
    s4 := '';
    Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from comparison');
    Aqr1.Open;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from comparison where pd_name=' + #39 + Aqr1.FieldByName('pd_name').AsString + #39);
      Aqr2.Open;
      if Aqr2.RecordCount > 1 then
      begin
        for j := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := j;
          if Aqr2.FieldByName('shuxing').AsString <> null then
            s1 := Aqr2.FieldByName('shuxing').AsString;
          if Aqr2.FieldByName('transport').AsString <> null then
            s2 := Aqr2.FieldByName('transport').AsString;
          if Aqr2.FieldByName('upstream').AsString <> null then
            s3 := Aqr2.FieldByName('upstream').AsString;
          if Aqr2.FieldByName('lcifactor').AsString <> null then
            s4 := Aqr2.FieldByName('lcifactor').AsString;
        end;

        Aqr2.First;
        Aqr2.Edit;
        Aqr2.FieldByName('shuxing').AsString := s1;
        Aqr2.FieldByName('transport').AsString := s2;
        Aqr2.FieldByName('upstream').AsString := s3;
        Aqr2.FieldByName('lcifactor').AsString := s4;
        Aqr2.Post;



        for j := Aqr2.RecordCount downto 2 do
        begin
          Aqr2.RecNo := j;
          Aqr2.Delete;
        end;


      end;


    end;


  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
  end;

  showmessage('  pppp   ');    }


//  showmessage('  pppp   ' + floattostr(CO2('压缩空气')));
//  showmessage('  pppp   ' + floattostr(energy('压缩空气')));

 // direct('压缩空气');

 // mat_energy;
//  mat_mainprocess;

//  checkmat(OA_WQ);

//CopyDir('c:\aa','d:\'  );



end;






procedure TForm1.PageControl1Change(Sender: TObject);
var
  i, j: integer;
  Aqr1, Aqr2: TAdoQuery;
  sheet: TTabSheet;
//  rootnode, subnode: ttreenode;
begin
  Panel7.Visible := false;
  FlatProgressBar1.Position := 0;
  {if TreeView1.Selected.Level = 0 then
  begin
    rootnode := TreeView1.Selected;
    if rootnode.HasChildren then
      rootnode.DeleteChildren;
  end;      }


  if PageControl1.TabIndex = 10 then
    CheckBox7.Visible := false
  else
    CheckBox7.Visible := true;

  if PageControl1.TabIndex = 10 then
    CheckBox8.Visible := false
  else
    CheckBox8.Visible := true;

  for i := 0 to PageControl1.PageCount - 1 do
  begin
    sheet := PageControl1.Pages[i];
    if sheet = PageControl1.ActivePage then
      sheet.Highlighted := true
    else
      sheet.Highlighted := false;
  end;

  if FileExists(ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.flc') then
  begin
    dxFlowChart1.Clear;
    dxFlowChart1.LoadFromFile(ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.flc');
    dxFlowChart1.Zoom := 0;
  end;



  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;

  try

    if PageControl1.TabIndex = 0 then
      TreeView1Click(Sender);


    if PageControl1.TabIndex = 1 then
      bsSkinButton16Click(Sender);

    if PageControl1.TabIndex = 2 then
    begin
      if FileExists(ExtractFilePath(Application.ExeName) + 'projects\' + label1.Caption + '\' + label1.Caption + '.flc') then
      begin

        if dxFlowChart1.ObjectCount = 0 then exit;
        listbox2.Items.Clear;
        listbox2.Items.Add('全部工序');
        for i := 0 to dxFlowChart1.ObjectCount - 1 do
          if dxFlowChart1.Objects[i].Text <> '主生产工序' then
            if trim(dxFlowChart1.Objects[i].Text) <> '能源公辅工序' then
              if dxFlowChart1.Objects[i].Text <> '至各工序' then
                listbox2.Items.Add(dxFlowChart1.Objects[i].Text);

        DBGridEh9.DataSource.Enabled := false;

       //删除冗余数据  例如载入流程图后，原先残留的数据

        Aqr1.Connection := Form1.ADOConnection2;
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from data ');
        Aqr1.Open;
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          if ListBox2.Items.IndexOf(Aqr1.FieldByName('process').AsString) = -1 then
            Aqr1.Delete;
        end;

        if listbox2.Items.IndexOf(edit2.Text) <> -1 then
        begin
          ListBox2.ItemIndex := listbox2.Items.IndexOf(edit2.Text);
          ListBox2.Selected[ListBox2.ItemIndex] := true;
        end
        else
        begin
          ListBox2.ItemIndex := 1;
          ListBox2.Selected[1] := true;
          edit2.Text := '';
        end;


        DBGridEh9.DataSource.Enabled := true;
        ProcessData_Format;
        DBGridEh9.Columns.Items[3].PickList.clear;
        if ListBox2.ItemIndex <> -1 then
          DBGridEh9.Columns.Items[3].picklist.Add(ListBox2.Items[ListBox2.ItemIndex]);
      end;

    end;

    if PageControl1.TabIndex = 3 then
    begin
      if dxFlowChart1.ObjectCount = 0 then exit;
      listbox2.Items.Clear;
      listbox2.Items.Add('全部工序');
      for i := 0 to dxFlowChart1.ObjectCount - 1 do
        if dxFlowChart1.Objects[i].Text <> '主生产工序' then
          if trim(dxFlowChart1.Objects[i].Text) <> '能源公辅工序' then
            if dxFlowChart1.Objects[i].Text <> '至各工序' then
              listbox2.Items.Add(dxFlowChart1.Objects[i].Text);


      if ListBox2.Items.Count = 0 then
        exit;
      StatusBar1.Panels[3].Text := '查找需要分配的单元';
      StatusBar1.Refresh;

//列出需要分配的单元
      AdoQuery2.Connection := Form1.ADOConnection2;
      ListBox3.Clear;
      for i := 0 to ListBox2.Items.Count - 1 do
      begin
        if AdoQuery2.Active then AdoQuery2.Close;
        AdoQuery2.SQL.Clear;
        AdoQuery2.SQL.Add('select * from data where (material_type= ' + #39 + '产品' + #39 + ' or material_type=' + #39 + '副产品或固体废弃物' + #39 + ')' + 'and process=' + #39 + ListBox2.Items[i] + #39);
        ADOQuery2.Open;
        ADOQuery2.Active := true;

        if ADOQuery2.RecordCount > 1 then
        begin
          if ListBox3.Items.IndexOf(ListBox2.Items[i]) = -1 then
            ListBox3.Items.Add(ListBox2.Items[i]);
        end;

      end;

      Aqr2.Connection := Form1.ADOConnection2;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from allocation');
      Aqr2.open;

      Aqr1.Connection := Form1.ADOConnection2;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data');
      Aqr1.open;
      StatusBar1.Panels[3].Text := '删除冗余数据';
      StatusBar1.Refresh;

      for j := Aqr2.RecordCount downto 1 do
      begin
        Aqr2.RecNo := j;
        if Aqr1.Locate('pd_name', Aqr2.FieldByName('albefore').AsString, []) = false then
          Aqr2.Delete;
      end; //删除冗余数据  数据名称变化后，分配表中的冗余数据

      ListBox3.ItemIndex := 0;
      ListBox3Click(Sender);
      allocation_Format;


    end;



    if PageControl1.TabIndex = 4 then
      shuxing_Format;

    if PageControl1.TabIndex = 5 then
      TransportData_Format;



    if PageControl1.TabIndex = 6 then
    begin

      //电力按构成比例分配



   //  bsSkinButton23Click(Sender);

      Label6Click(Sender);

    end;

    if PageControl1.TabIndex = 7 then
    begin

      if DBGridEh8.DataSource.DataSet.Active = false then
      begin
        bsSkinCheckListBox1.Clear;
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39);
        Aqr1.Open;
        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          bsSkinCheckListBox1.Items.Add(Aqr1.FieldByName('pd_name').AsString);
        end;

        ComboBox5.ItemIndex := 0;
        ComboBox4.ItemIndex := 0;

        ComboBox5Change(Sender);
        ComboBox4Change(Sender);

        CheckBox4.Checked := true;
        CheckBox5.Checked := true;

        CheckBox4Click(Sender);
        RadioButton9.Checked := true;
        RadioButton9Click(Sender);
      end;


  {  if TreeView1.Selected.Level = 0 then
    begin
      rootnode := TreeView1.Selected;
      rootnode.DeleteChildren;
      subnode := TreeView1.Items.Addchild(rootnode, '主生产工序');

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select distinct process,material_type, processtype,pd_name from data_allocation where processtype = ' + #39 + '主生产工序' + #39 + ' and material_type=' + #39 + '产品' + #39);
      Aqr1.Open;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        TreeView1.Items.Addchild(subnode, Aqr1.FieldByName('pd_name').AsString);
      end;


      subnode := TreeView1.Items.Addchild(rootnode, '能源公辅工序');
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select distinct process,material_type, processtype,pd_name from data_allocation where processtype = ' + #39 + '能源公辅工序' + #39 + ' and material_type=' + #39 + '产品' + #39);
      Aqr1.Open;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        TreeView1.Items.Addchild(subnode, Aqr1.FieldByName('pd_name').AsString);
      end;

   //   TreeView1.FullExpand;
        for i:=0 to TreeView1.Items[0].Count-1 do
    TreeView1.Items[0].Item[i].Collapse(True);
  TreeView1.Items[0].Expand(False)   ;
    end;
                  }



    end;






    if PageControl1.TabIndex = 9 then
    begin
      newprojectForm.ListBox1Renew;
      bsSkinCheckListBox8.Clear;
      for i := 0 to newprojectForm.ListBox1.Items.Count - 1 do
        bsSkinCheckListBox8.Items.Add(newprojectForm.ListBox1.Items[i]);

      bsSkinCheckListBox5.Clear;

      Panel7.BringToFront;
      Panel7.Visible := true;
      ComboBox7.Visible := true;
      ComboBox7.Items.Clear;
      ComboBox7.Items.Add('全部LCI类别');
      ComboBox7.Items.Add('能源类');
      ComboBox7.Items.Add('资源类');
      ComboBox7.Items.Add('大气污染物');
      ComboBox7.Items.Add('水体污染物');
      ComboBox7.Items.Add('固体废弃物');
      ComboBox7.Items.Add('LCIA环境影响评价类型');


      ComboBox7.ItemIndex := -1;
      ComboBox7Change(Sender);
      chart1.Series[1].Clear;
      chart1.Series[0].Clear;
      chart1.Title.Text.Clear;

    end;


    if PageControl1.TabIndex = 8 then
    begin
    {  image2.Hint := 'X：  未包含废钢回收的LCA结果（见前面LCA计算结果）' + #13 + #13;
      image2.Hint := image2.Hint + 'RR： 废钢回收率' + #13 + #13;
      image2.Hint := image2.Hint + 'S： 本项目中转炉炼钢废钢加入量（废钢单耗）' + #13 + #13;
      image2.Hint := image2.Hint + 'Xpr： 利用全铁矿石生产钢材的LCA结果（不含废钢循环）' + #13 + #13;
      image2.Hint := image2.Hint + 'Xre： 利用全废钢生产钢材的LCA结果（不含废钢循环）' + #13 + #13;
      image2.Hint := image2.Hint + 'Y： 利用全废钢炼钢过程中废钢利用效率（即，电炉炼钢1kg废钢可以炼Y kg钢水）' + #13 + #13;
      image2.Hint := image2.Hint + '废钢循环仅对转炉之后（含转炉）工序LCA结果产生影响' + #13;
        }
      ComboBox1.Items.Clear;
      ComboBox9.Items.Clear;
      Aqr1.Connection := Form1.ADOConnection2;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select distinct process from data_allocation where processtype = ' + #39 + '主生产工序' + #39);
      Aqr1.open;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        ComboBox1.Items.Add(Aqr1.FieldByName('process').AsString);
        ComboBox9.Items.Add(Aqr1.FieldByName('process').AsString);
      end;


      Aqr1.Connection := Form1.ADOConnection2;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from scrapeLCA ');
      Aqr1.open;
      bsSkinEdit2.Text := '';
      if Aqr1.Locate('inf_name', 'RR', []) = true then
        bsSkinEdit2.Text := Aqr1.FieldByName('num').AsString;

      bsSkinEdit3.Text := '';
      if Aqr1.Locate('inf_name', 'Y', []) = true then
        bsSkinEdit3.Text := Aqr1.FieldByName('num').AsString;

      if Aqr1.Locate('inf_name', 'Xre_source', []) = true then
      begin
        if Aqr1.FieldByName('num').AsFloat = -1 then
          RadioButton1.Checked := true;
        if Aqr1.FieldByName('num').AsFloat > -1 then
        begin
          RadioButton2.Checked := true;
          ComboBox9.ItemIndex := Aqr1.FieldByName('num').AsInteger;
        end;
      end;

      ListBox6.Clear;
      ComboBox3.Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from scrapeLCA where inf_name=' + #39 + 'S' + #39);
      Aqr1.open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        ListBox6.Items.Add(Aqr1.FieldByName('inf').AsString);
        ComboBox3.Items.Add(Aqr1.FieldByName('num').AsString);
        ComboBox3.Text := Aqr1.FieldByName('num').AsString;
      end;


      if ListBox6.Items.Count > 0 then
      begin
        ListBox6.ItemIndex := 0;
        ListBox6.Selected[0] := true;
        ComboBox3.ItemIndex := 0;
      end;

      ListBox7.Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select distinct inf from scrapeLCA where inf_name=' + #39 + 'LCIincludingEOL' + #39);
      Aqr1.Open;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if ListBox7.Items.IndexOf(Aqr1.FieldByName('inf').AsString) = -1 then
          ListBox7.Items.Add(Aqr1.FieldByName('inf').AsString);
      end;


    end;

    if PageControl1.TabIndex = 10 then
    begin
      newprojectForm.ListBox1Renew;
      bsSkinCheckListBox8.Clear;
      for i := 0 to newprojectForm.ListBox1.Items.Count - 1 do
        bsSkinCheckListBox8.Items.Add(newprojectForm.ListBox1.Items[i]);

      bsSkinCheckListBox5.Clear;

      Panel7.BringToFront;
      Panel7.Visible := true;
      ComboBox7.Items.Clear;
      ComboBox7.Items.Add('工序能耗分析比较');
      ComboBox7.Items.Add('碳排放分析比较');
      ComboBox7.Items.Add('铁平衡');
      ComboBox7.Items.Add('全流程物料/能源平衡');
      ComboBox7.Text := '选择数据分析内容';
      bsSkinCheckListBox55.Clear;
    end;


  finally
    aqr1.Close;
    aqr1.Free;
    aqr2.Close;
    aqr2.Free;
    StatusBar1.Panels[3].Text := '';
    StatusBar1.Refresh;
  end;

end;

procedure TForm1.Button2Click(Sender: TObject); //LCA结果查询
var
  i, j: integer;
  mysql: string;
  Aqr3: TAdoQuery;
begin

  if checkedcount(bsSkinCheckListBox1) = 0 then
  begin
    AdoQuery6.Close;
    exit;
  end;
  if checkedcount(bsSkinCheckListBox2) = 0 then
  begin
    AdoQuery6.Close;
    exit;
  end;


  mysql := 'select * from LCA where (product= ';

  for i := 0 to bsSkinCheckListBox1.Items.Count - 1 do
    if bsSkinCheckListBox1.Checked[i] then
    begin
      mysql := mysql + #39 + bsSkinCheckListBox1.Items.Strings[i] + #39;
      break;
    end;

  for j := i + 1 to bsSkinCheckListBox1.Items.Count - 1 do
    if bsSkinCheckListBox1.Checked[j] then
      mysql := mysql + ' or  product= ' + #39 + bsSkinCheckListBox1.Items.Strings[j] + #39;

  mysql := mysql + ')';



  for i := 0 to bsSkinCheckListBox2.Items.Count - 1 do
    if bsSkinCheckListBox2.Checked[i] then
    begin
      mysql := mysql + ' and ( LCIfactor= ' + #39 + bsSkinCheckListBox2.Items.Strings[i] + #39;
      break;
    end;

  for j := i + 1 to bsSkinCheckListBox2.Items.Count - 1 do
    if bsSkinCheckListBox2.Checked[j] then
      mysql := mysql + ' or (LCIfactor= ' + #39 + bsSkinCheckListBox2.Items.Strings[j] + #39 + ')';


  if checkedcount(bsSkinCheckListBox2) <> 0 then
    mysql := mysql + ')';


  if RadioButton9.Checked = true then
    mysql := mysql + ' and LCItype <> ' + #39 + 'LCIA' + #39 + ' and LCItype <> ' + #39 + 'LCC' + #39;
  if RadioButton10.Checked = true then
    mysql := mysql + ' and LCItype = ' + #39 + 'LCIA' + #39;
  if RadioButton16.Checked = true then
    mysql := mysql + ' and LCItype = ' + #39 + 'LCC' + #39;
  if RadioButton3.Checked = true then
    mysql := mysql + ' and LCItype = ' + #39 + 'BPEI' + #39;




   //   showmessage( mysql);

  if AdoQuery6.Active then AdoQuery6.Close;
  AdoQuery6.SQL.Clear;
  AdoQuery6.SQL.Add(mysql);
  ADOQuery6.Open;

  DBGridEh8.Columns[15].Visible := false;
  for i := 0 to 14 do
    DBGridEh8.Columns[i].Visible := true;


  for i := 0 to DBGridEh8.Columns.Count - 2 do
  begin
    DBGridEh8.Columns[i].Width := 70;
    DBGridEh8.Columns[i].Title.Alignment := taCenter;
    DBGridEh8.Columns[i].Alignment := taCenter;
  end;
  DBGridEh8.Columns[0].Width := 30;
  DBGridEh8.Columns[2].Width := 80;
  DBGridEh8.Columns[3].Width := 120;
  DBGridEh8.Columns[4].Width := 40;

  for i := 5 to 14 do
    DBGridEh8.Columns[i].Alignment := taRightJustify; //数据

  DBGridEh8.Columns[1].Alignment := taLeftJustify;
  DBGridEh8.Columns[2].Alignment := taLeftJustify;
  DBGridEh8.Columns[3].Alignment := taLeftJustify;

  DBGridEh8.Columns[0].Title.caption := 'ID';
  DBGridEh8.Columns[1].Title.caption := '产品';
  DBGridEh8.Columns[2].Title.caption := 'LCI类别';
  DBGridEh8.Columns[3].Title.caption := 'LCI指标';
  DBGridEh8.Columns[4].Title.caption := '单位';
  DBGridEh8.Columns[5].Title.caption := '直接负荷';
  DBGridEh8.Columns[6].Title.caption := '能源负荷';
  DBGridEh8.Columns[7].Title.caption := '中间产品负荷';
  DBGridEh8.Columns[8].Title.caption := '副产品内部收益';
  DBGridEh8.Columns[9].Title.caption := '内部合计';
  DBGridEh8.Columns[10].Title.caption := '运输负荷';
  DBGridEh8.Columns[11].Title.caption := '上游负荷';
  DBGridEh8.Columns[12].Title.caption := '副产品外部收益';
  DBGridEh8.Columns[13].Title.caption := '外部合计';
  DBGridEh8.Columns[14].Title.caption := '生命周期';

  if RadioButton9.Checked = true then
  begin
    DBGridEh8.Columns[3].Title.caption := 'LCI指标';
    DBGridEh8.Columns[2].Visible := true;
  end;

  if RadioButton10.Checked = true then
  begin
    DBGridEh8.Columns[3].Title.caption := 'LCIA指标';
    DBGridEh8.Columns[2].Visible := false;
    DBGridEh8.Columns[4].Width := 120;
  end;

  if RadioButton16.Checked = true then
  begin
    DBGridEh8.Columns[3].Title.caption := '指标';
    DBGridEh8.Columns[2].Visible := false;
    DBGridEh8.Columns[4].Width := 120;
  end;

  if RadioButton3.Checked = true then
  begin
    DBGridEh8.Columns[3].Title.caption := 'LCI指标';
    DBGridEh8.Columns[2].Visible := true;

    for i := 5 to 14 do
      DBGridEh8.Columns[i].Visible := false;

    DBGridEh8.Columns[9].Visible := true;
    DBGridEh8.Columns[14].Visible := true;
    DBGridEh8.Columns[9].Title.caption := '宝钢内部';
    DBGridEh8.Columns[14].Title.caption := '生命周期';
  end;




end;

procedure TForm1.CheckBox4Click(Sender: TObject);
var
  i: integer;
begin

  if CheckBox4.Checked = true then
    for i := 0 to bsSkinCheckListBox1.Items.Count - 1 do
      bsSkinCheckListBox1.Checked[i] := true;
  if CheckBox4.Checked = false then
    for i := 0 to bsSkinCheckListBox1.Items.Count - 1 do
      bsSkinCheckListBox1.Checked[i] := false;
  Button2Click(Sender); //LCA结果查询

end;

procedure TForm1.ComboBox4Change(Sender: TObject);
var
  mysql1: string;
  Aqr1: TAdoQuery;
  i: integer;
begin

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection2;

  try
    Aqr1.Close;
    Aqr1.SQL.Clear;
    mysql1 := 'select * from lcifactor ';
    if ComboBox4.ItemIndex > 0 then
      mysql1 := mysql1 + 'where inf_type = ' + #39 + ComboBox4.Text + #39;
    Aqr1.SQL.Add(mysql1);
    Aqr1.Open;

    bsSkinCheckListBox2.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if bsSkinCheckListBox2.Items.IndexOf(Aqr1.FieldByName('inf_name').AsString) = -1 then
        bsSkinCheckListBox2.Items.Add(Aqr1.FieldByName('inf_name').AsString);
    end;

    CheckBox5.Checked := false;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;


end;

procedure TForm1.CheckBox5Click(Sender: TObject);
var
  i: integer;
begin

  if CheckBox5.Checked = true then
    for i := 0 to bsSkinCheckListBox2.Items.Count - 1 do
      bsSkinCheckListBox2.Checked[i] := true;
  if CheckBox5.Checked = false then
    for i := 0 to bsSkinCheckListBox2.Items.Count - 1 do
      bsSkinCheckListBox2.Checked[i] := false;
  Button2Click(Sender); //LCA结果查询
end;

procedure TForm1.ADOQuery6AfterOpen(DataSet: TDataSet);
var
  i: integer;
begin
  ADOQuery6.FieldByName('ID').OnGetText := SetFdText6;
  for i := 1 to ADOQuery6.FieldCount - 1 do
    if ADOQuery6.Fields[i].DataType = ftFloat then
    begin
      if RadioButton10.Checked = false then
        (ADOQuery6.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0'
      else
        (ADOQuery6.Fields[i] as Tfloatfield).DisplayFormat := '###,##0.####0';
    end;
  ADOQuery6.First;


end;

procedure TForm1.ComboBox5Change(Sender: TObject);
var
  mysql1: string;
  Aqr1: TAdoQuery;
  i: integer;
begin

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection2;

  try
    Aqr1.Close;
    Aqr1.SQL.Clear;
    mysql1 := 'select * from data_allocation where material_type = ' + #39 + '产品' + #39;
    if ComboBox5.ItemIndex > 0 then
      mysql1 := mysql1 + ' and processtype=' + #39 + ComboBox5.Text + #39;
    Aqr1.SQL.Add(mysql1);
    Aqr1.Open;

    bsSkinCheckListBox1.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      bsSkinCheckListBox1.Items.Add(Aqr1.FieldByName('pd_name').AsString);
    end;

    CheckBox4.Checked := false;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;

end;

procedure TForm1.DBGridEh8GetCellParams(Sender: TObject; Column: TColumnEh;
  AFont: TFont; var Background: TColor; State: TGridDrawState);
begin
  if not (gdSelected in State) then
    if (Sender as TDBGridEh).SumList.RecNo mod 2 = 0 then
      Background := clInfobk;
end;

procedure TForm1.RadioButton9Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;
begin
  if RadioButton9.Checked then
  begin
    ComboBox4.Enabled := true;
    bsSkinCheckListBox2.Clear;
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := ADOConnection2;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from lcifactor');
    Aqr1.Open;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      bsSkinCheckListBox2.Items.Add(Aqr1.FieldByName('inf_name').AsString);
    end;
    CheckBox5.Checked := true;
    CheckBox5Click(Sender);
    Aqr1.Close;
    Aqr1.Free;
  end;
  Button2Click(Sender); //LCA结果查询
end;

procedure TForm1.RadioButton10Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;
begin
  if RadioButton10.Checked then
  begin
    ComboBox4.Enabled := false;
    bsSkinCheckListBox2.Clear;
    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := ADOConnection2;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct impact_type from char_coef');
    Aqr1.Open;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      bsSkinCheckListBox2.Items.Add(Aqr1.FieldByName('impact_type').AsString);
    end;
    CheckBox5.Checked := true;
    CheckBox5Click(Sender);
    Aqr1.Close;
    Aqr1.Free;
  end;
  Button2Click(Sender); //LCA结果查询
end;

procedure TForm1.bsSkinCheckListBox2ClickCheck(Sender: TObject);
begin
  Button2Click(Sender); //LCA结果查询
end;

procedure TForm1.bsSkinCheckListBox1ClickCheck(Sender: TObject);
begin
  Button2Click(Sender); //LCA结果查询
end;

function TForm1.CopyDir(SrcDir, DesDir: string): Boolean; //复制文件夹
var
  ss: TSHFileOpStruct;
begin
  Result := False;
  if not DirectoryExists(SrcDir) then Exit;
  FillChar(ss, SizeOf(ss), 0);
  ss.Wnd := Handle;
  SS.pFrom := PChar(SrcDir + #0);
  ss.pTo := PChar(DesDir + #0);
  ss.wFunc := FO_COPY;
  ss.fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
  Result := SHFileOperation(ss) = 0;
end;



procedure TForm1.N15Click(Sender: TObject);
begin
  aboutform.ShowModal;
end;

procedure TForm1.N14Click(Sender: TObject);
var
  path: string;
begin
  path := ExtractFilePath(Application.ExeName) + '\manual.pdf';
  ShellExecute(Handle, 'Open', PChar(path), nil, nil, SW_SHOW);


  //  ShellExecute(Handle,'Open','1.pdf',nil,'D:\',SW_NORMAL);

 //   ShellExecute(Handle, 'Open', PChar(path), nil, nil, SW_SHOW);

{  if AdobeReaderInstalled = True then
  begin
    path := ExtractFilePath(Application.ExeName) + '\1.doc';
    ShellExecute(Handle, 'Open', PChar(path), nil, nil, SW_SHOW);
  end
  else
    bsSkinMessage1.MessageDlg('请安装PDF浏览器！', (mtInformation), [mbOK], 0);

        }
end;


procedure TForm1.Label30Click(Sender: TObject);
var
  i: integer;
begin
  bsSkinPanel2.Visible := false; //功能单位
  Label4.Caption := Label1.Caption + '4';
  for i := 0 to 12 do
    PageControl1.Pages[i].TabVisible := false;

  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := true;
  bsSkinPanel2.Visible := false;

  ListBox4.Clear;
  ListBox4.Items.Add('全部');
  ListBox4.Items.Add('资源类');
  ListBox4.Items.Add('能源类');
  ListBox4.Items.Add('大气污染物');
  ListBox4.Items.Add('水体污染物');
  ListBox4.Items.Add('固体废弃物');



  ADOQuery1.Close;
  ListBox4Click(Sender);
end;

procedure TForm1.N22Click(Sender: TObject);
var
  B: TBitmap;
  str: string;
begin
  B := TBitmap.Create;
  PaintCtrlToBmp(dxFlowChart2, B);
  if savedialog1.Execute then
  begin
    savedialog1.DefaultExt := 'bmp';
    str := savedialog1.FileName + '.bmp';
  end;
  try
    if trim(str) <> '' then
      B.SaveToFile(str);
  finally
    B.Free;
  end;
end;

procedure Tform1.PaintCtrlToBmp(Ctrl: TWinControl; Bmp: TBitmap);
var
  cCanvas: TControlCanvas;
begin
  Bmp.Width := Ctrl.Width;
  Bmp.Height := Ctrl.Height;
  cCanvas := TControlCanvas.Create;
  try
    cCanvas.Control := Ctrl;
    Bmp.Canvas.CopyRect(Rect(0, 0, Bmp.Width, Bmp.Height), cCanvas, Ctrl.ClientRect);
  finally
    cCanvas.Free;
  end;
end;

procedure TForm1.RadioButton16Click(Sender: TObject);
begin
  if RadioButton16.Checked then
  begin
    ComboBox4.Enabled := false;
    bsSkinCheckListBox2.Clear;

    bsSkinCheckListBox2.Items.Add('生命周期成本');

    CheckBox5.Checked := true;
    CheckBox5Click(Sender);

  end;
  Button2Click(Sender); //LCA结果查询
end;

procedure TForm1.N39Click(Sender: TObject);
begin
  close;
end;

procedure TForm1.LCA1Click(Sender: TObject);
begin
  newprojectForm.ListBox1Renew;
  newprojectForm.ShowModal;
end;


function TForm1.ReturnName(Value: pchar): string; //返回文件夹名
begin
  result := AnsiStrRScan(Value, '\');
  Delete(result, 1, 1);
end;

procedure TForm1.LCA2Click(Sender: TObject);
var
  Dir, s: string;
begin
  s := ExtractFilePath(Application.ExeName) + 'projects\';
  newprojectForm.ListBox1Renew;
  if SelectDirectory('选择项目文件夹', '', Dir) then
    if newprojectForm.ListBox1.Items.IndexOf(ReturnName(pchar(dir))) <> -1 then
    begin
      if form1.bsSkinMessage1.MessageDlg(ReturnName(pchar(dir)) + '文件夹已存在，是否覆盖?', mtConfirmation,
        [mbYes, mbNo], 0) = mrYes then
      begin
        CopyDir(dir, s);
        bsSkinMessage1.MessageDlg('导入完成！', (mtInformation), [mbOK], 0);
      end
      else
        exit;
    end;

end;

procedure TForm1.ComboBox7Change(Sender: TObject);
var
  mysql1, s: string;
  Aqr1: TAdoQuery;
  i: integer;
begin

  Aqr1 := TAdoQuery.Create(nil);


  try
    if PageControl1.TabIndex = 9 then
    begin

      Aqr1.Close;
      Aqr1.Connection := form1.ADOConnection1;

      Aqr1.SQL.Clear;
      mysql1 := 'select * from lcifactor ';
      if ComboBox7.ItemIndex > 0 then
        mysql1 := mysql1 + 'where inf_type = ' + #39 + ComboBox7.Text + #39;
      if ComboBox7.ItemIndex = 6 then
        mysql1 := 'select distinct impact_type from char_coef';

      if ComboBox7.ItemIndex = -1 then
        mysql1 := 'select * from lcifactor where see_level= ' + #39 + '1' + #39;


      Aqr1.SQL.Add(mysql1);
      Aqr1.Open;

      bsSkinCheckListBox55.Clear;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if ComboBox7.ItemIndex <> 6 then
        begin
          if bsSkinCheckListBox55.Items.IndexOf(Aqr1.FieldByName('inf_name').AsString) = -1 then
            bsSkinCheckListBox55.Items.Add(Aqr1.FieldByName('inf_name').AsString);
        end
        else
        begin
          if bsSkinCheckListBox55.Items.IndexOf(Aqr1.FieldByName('impact_type').AsString) = -1 then
            bsSkinCheckListBox55.Items.Add(Aqr1.FieldByName('impact_type').AsString);


        end;

      end;
    end; //if   PageControl1.TabIndex=9


    if PageControl1.TabIndex = 10 then
    begin
      bsSkinCheckListBox55.Clear;

      if ComboBox7.ItemIndex = 3 then
      begin
        if checkedcount(bsSkinCheckListBox8) = 0 then ADOQuery10.Close;
        if checkedcount(bsSkinCheckListBox8) = 0 then exit;

        for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
        begin
          if bsSkinCheckListBox8.Checked[i] = true then
          begin
            s := bsSkinCheckListBox8.Items[i];
            break;
          end;
        end;



        ADOConnection5.Close;
        ADOConnection5.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
        ADOConnection5.Open;
        AQr1.Connection := form1.ADOConnection5;

        bsSkinCheckListBox55.Clear;
        Aqr1.Close;
        Aqr1.SQL.Clear;
        mysql1 := 'select distinct pd_name from data_allocation order by pd_name asc';
        Aqr1.SQL.Add(mysql1);
        Aqr1.Open;

        for i := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := i;
          if bsSkinCheckListBox55.Items.IndexOf(Aqr1.FieldByName('pd_name').AsString) = -1 then
          begin
            bsSkinCheckListBox55.Items.Add(Aqr1.FieldByName('pd_name').AsString);

          end;

        end;

      end; //   if ComboBox7.ItemIndex = 3 then
      Button3Click(Sender);
    end;

  finally
    Aqr1.Close;
    Aqr1.Free;
  end;


end;

procedure TForm1.bsSkinButton31Click(Sender: TObject);
var
  mysql1, s: string;
  i, j: integer;
  Aqr1: TAdoQuery;
begin
  bsSkinEdit1.Text := '';
  if checkedcount(bsSkinCheckListBox8) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个项目', (mtInformation), [mbOK], 0);
    exit;
  end;

  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;



  dxFlowChart2.Visible := true;
  chart1.Visible := false;
  CheckBox6.Visible := false;

  if FileExists(ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.flc') then
  begin
    dxFlowChart2.Clear;
    dxFlowChart2.LoadFromFile(ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.flc');
    dxFlowChart2.Zoom := 0;
  end
  else
  begin
    bsSkinMessage1.MessageDlg('未找到流程图文件', (mtInformation), [mbOK], 0);
    exit;
  end;


  Aqr1 := TAdoQuery.Create(nil);
  ADOConnection6.Close;
  ADOConnection6.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection6.Open;



  Aqr1.Connection := form1.ADOConnection6; //这里有问题。连接有问题

  try

    for i := 0 to dxFlowChart2.ConnectionCount - 1 do
    begin
      if dxFlowChart2.Connections[i].ObjectSource = nil then continue;
      mysql1 := 'select * from data_allocation where process = ' + #39 + dxFlowChart2.Connections[i].ObjectSource.Text + #39 + ' and (material_type = ' + #39 + '产品' + #39 + ' or material_type =' + #39 + '副产品或固体废弃物' + #39 + ')';

      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add(mysql1);
      Aqr1.Open;
      Listbox1.Clear;
      if Aqr1.RecordCount > 0 then
        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          if listbox1.Items.IndexOf(Aqr1.FieldbyName('pd_name').AsString) = -1 then
            listbox1.Items.Add(Aqr1.FieldbyName('pd_name').AsString);
        end;
      if listbox1.Items.Count = 0 then continue;

      if dxFlowChart2.Connections[i].ObjectDest = nil then continue;
      mysql1 := 'select * from data_allocation where process  = ' + #39 + dxFlowChart2.Connections[i].ObjectDest.Text + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )';
      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add(mysql1);
      Aqr1.Open;
      if Aqr1.RecordCount > 0 then
        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          if listbox1.Items.IndexOf(Aqr1.FieldbyName('pd_name').AsString) <> -1 then
            if Aqr1.FieldbyName('co_num').AsString <> '' then
         //       dxFlowChart1.Connections[i].Text :=  Aqr1.FieldbyName('name').AsString + '  ' + FormatFloat('#,##0.000', Aqr1.FieldbyName('xishu').AsFloat)
        //          + ' ' + Aqr1.FieldbyName('xishu_unit').AsString;

              dxFlowChart2.Connections[i].Text := FormatFloat('#,##0.000', Aqr1.FieldbyName('co_num').AsFloat);
          dxFlowChart2.Connections[i].Font.Color := clBlue;
        end;
    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;





end;

procedure TForm1.bsSkinButton32Click(Sender: TObject);
var
  mysql1, s, product, factor: string;
  i: integer;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := true;
  chart1.Series[1].Active := false;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;
  mysql1 := 'select * from LCA where product= ' + #39 + product + #39 + ' and lcifactor=' + #39 + factor + #39;
  ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add(mysql1);
  ADOQuery7.Open;

  chart1.Series[0].Clear;
  chart1.Title.Text.Clear;
  chart1.Title.Text.Add(ADOQuery7.FieldByName('product').AsString + '生命周期' + ADOQuery7.FieldByName('LCIfactor').AsString + '负荷构成    （' + ADOQuery7.FieldByName('unit').AsString + '）');
  for i := 5 to 8 do
    chart1.Series[0].Add(ADOQuery7.Fields[i].AsFloat, ADOQuery7.Fields[i].FieldName, clteecolor);
  for i := 10 to 12 do
    chart1.Series[0].Add(ADOQuery7.Fields[i].AsFloat, ADOQuery7.Fields[i].FieldName, clteecolor);


  bsSkinEdit1.Text := '';
end;

procedure TForm1.bsSkinCheckListBox8ClickCheck(Sender: TObject);
var
  mysql1, s: string;
  i, j: integer;
  Aqr1: TAdoQuery;
begin
  if checkedcount(bsSkinCheckListBox8) = 0 then exit;

  if PageControl1.TabIndex = 10 then //只能单选
  begin
    j := bsSkinCheckListBox8.ItemIndex;
    if bsSkinCheckListBox8.Selected[j] then
      for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
        bsSkinCheckListBox8.Checked[i] := i = j;
  end;

  if PageControl1.TabIndex = 10 then
    CheckBox7.Visible := false
  else
    CheckBox7.Visible := true;



  Aqr1 := TAdoQuery.Create(nil);


  try
    bsSkinCheckListBox5.Clear;
    for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
    begin
      if bsSkinCheckListBox8.Checked[i] = true then
      begin
        s := bsSkinCheckListBox8.Items[i];
        ADOConnection3.Close;
        ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
        ADOConnection3.Open;
        Aqr1.Connection := form1.ADOConnection3;
        mysql1 := 'select distinct product from LCA ';
        Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add(mysql1);
        Aqr1.Open;
        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          if bsSkinCheckListBox5.Items.IndexOf(Aqr1.FieldByName('product').AsString) = -1 then
            bsSkinCheckListBox5.Items.Add(Aqr1.FieldByName('product').AsString);

        end;



      end;





    end;




  finally
    Aqr1.Close;
    Aqr1.Free;
  end;


end;

procedure TForm1.bsSkinButton33Click(Sender: TObject);
var
  mysql1, s, product, factor: string;
  i: integer;

begin
  bsSkinEdit1.Text := '';
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := true;
  chart1.Series[1].Active := false;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;
  mysql1 := 'select * from LCA where product= ' + #39 + product + #39 + ' and lcifactor=' + #39 + factor + #39;
  ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add(mysql1);
  ADOQuery7.Open;

  chart1.Series[0].Clear;
  chart1.Title.Text.Clear;
  chart1.Title.Text.Add(ADOQuery7.FieldByName('product').AsString + '生命周期' + ADOQuery7.FieldByName('LCIfactor').AsString + '负荷构成    （' + ADOQuery7.FieldByName('unit').AsString + '）');
  chart1.Series[0].Add(ADOQuery7.Fields[9].AsFloat, '企业内部负荷', clteecolor);
  chart1.Series[0].Add(ADOQuery7.Fields[10].AsFloat, '外部运输负荷', clteecolor);
  chart1.Series[0].Add(ADOQuery7.Fields[11].AsFloat, '上游负荷', clteecolor);
  chart1.Series[0].Add(ADOQuery7.Fields[12].AsFloat, '外卖副产品收益', clteecolor);

end;

procedure TForm1.CheckBox6Click(Sender: TObject);
begin
  if chart1.Series[0].Active = true then
  begin
    if CheckBox6.Checked = true then
      chart1.Series[0].Marks.style := smsPercent
    else
      chart1.Series[0].Marks.style := smsValue;
  end;

  if chart1.Series[1].Active = true then
  begin
    if CheckBox6.Checked = true then
      chart1.Series[1].Marks.style := smsLabelPercent
    else
      chart1.Series[1].Marks.style := smsLabelValue;
  end;

end;

procedure TForm1.MenuItem1Click(Sender: TObject);

begin
  form1.SaveDialog1.Filter := '*.bmp';
  if form1.SaveDialog1.Execute then
    chart1.SaveToBitmapFile(form1.SaveDialog1.FileName + '.bmp');

end;

procedure TForm1.MenuItem2Click(Sender: TObject);
begin
  if DBGrideh10.DataSource.DataSet.Active = false then exit;
 //   SaveDialog1.FileName    :=    Chart1.Series[0].Title       ;

  DBgridToExcel(DBGrideh10);
  DBGridEh10.DataSource.DataSet.First;
end;

procedure TForm1.bsSkinButton42Click(Sender: TObject); //g1
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  a, a1, a2: double;
  s, product1, process, factor: string;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;
  FlatProgressBar1.Position := 0;
  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := false;
  chart1.Series[1].Active := true;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection1;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;

  try

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;

    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select id, pd_name, lcifactor  from comparison where lcifactor= ' + #39 + factor + #39);
    Aqr3.Open;

    a := 0;

    if factor <> '(a)二氧化碳' then
      if factor <> '(e)总一次能源' then
        for j := 1 to Aqr3.RecordCount do //  黄河水 除盐水        (r)水耗
        begin
          Aqr3.RecNo := j;

          if Aqr4.Active then Aqr4.Close;
          Aqr4.SQL.Clear;
          Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39 + ' and pd_name=' + #39 + Aqr3.FieldByName('pd_name').AsString + #39);
          Aqr4.Open;
          if Aqr4.RecordCount > 0 then
            for k := 1 to Aqr4.RecordCount do
            begin
              Aqr4.RecNo := k;
              if Aqr4.FieldByName('material_type').AsString = '产品' then
                a := (-1) * Aqr4.FieldByName('co_num').AsFloat;
              if Aqr4.FieldByName('material_type').AsString = '副产品或固体废弃物' then
                a := (-1) * Aqr4.FieldByName('co_num').AsFloat;
              if Aqr4.FieldByName('material_type').AsString = '能源及能源介质' then
                a := Aqr4.FieldByName('co_num').AsFloat;
              if Aqr4.FieldByName('material_type').AsString = '原材料' then
                a := Aqr4.FieldByName('co_num').AsFloat;
              if Aqr4.FieldByName('material_type').AsString = '辅助材料' then
                a := Aqr4.FieldByName('co_num').AsFloat;
              if Aqr4.FieldByName('material_type').AsString = '大气排放' then
                a := (-1) * Aqr4.FieldByName('co_num').AsFloat;
              if Aqr4.FieldByName('material_type').AsString = '水体排放' then
                a := (-1) * Aqr4.FieldByName('co_num').AsFloat;

              ADOQuery7.Append;
              ADOQuery7.FieldByName('lcifactor').AsString := factor;
              ADOQuery7.FieldByName('process').AsString := process;
              ADOQuery7.FieldByName('pd_name').AsString := Aqr3.FieldByName('pd_name').AsString;
              ADOQuery7.FieldByName('AF').AsFloat := a;
              ADOQuery7.FieldByName('EF').AsFloat := 1;
              ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
              ADOQuery7.Post;
            end;

        end;


    if factor = '(a)二氧化碳' then
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from shuxing where shuxing = ' + #39 + 'CO2排放系数' + #39);
      Aqr1.Open;

      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, shuxing from comparison where shuxing is not null ');
      Aqr4.Open;



      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
      Aqr2.Open;



      for j := 1 to Aqr2.RecordCount do
      begin
        Aqr2.RecNo := j;
        if Aqr4.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
        begin
          if Aqr1.Locate('name', Aqr4.FieldByName('shuxing').AsString, []) = true then
          begin
            if Aqr2.FieldByName('material_type').AsString = '产品' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '副产品或固体废弃物' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '能源及能源介质' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '原材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '辅助材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '大气排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '水体排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;

            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr2.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := a;
            ADOQuery7.FieldByName('EF').AsFloat := Aqr1.FieldByName('shuzhi').AsFloat * 1000;
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;

          end;
        end;
      end;
    end;


    if factor = '(e)总一次能源' then
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from shuxing where shuxing = ' + #39 + '热值' + #39);
      Aqr1.Open;

      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, shuxing from comparison where shuxing is not null ');
      Aqr4.Open;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
      Aqr2.Open;


      for j := 1 to Aqr2.RecordCount do
      begin
        Aqr2.RecNo := j;
        if Aqr4.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
        begin
          if Aqr1.Locate('name', Aqr4.FieldByName('shuxing').AsString, []) = true then
          begin
            if Aqr2.FieldByName('material_type').AsString = '产品' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '副产品或固体废弃物' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '能源及能源介质' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '原材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '辅助材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '大气排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '水体排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;

            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr2.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := a;
            ADOQuery7.FieldByName('EF').AsFloat := Aqr1.FieldByName('shuzhi').AsFloat / 1000;
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;
    end;




    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;
    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;




    chart1.Series[1].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;
      if a1 - a2 > 0.0000000001 then
        chart1.Series[1].Add(a1 - a2, '其它', clteecolor);

    end; //  >8

    chart1.Title.Text.Clear;
    chart1.Title.Text.Add(product1 + '生命周期' + factor + '负荷构成 （工序直接负荷）');

    CheckBox6.Checked := true;
    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);
  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;



end;

procedure TForm1.bsSkinButton60Click(Sender: TObject); //g2
var
  i, j: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := false;
  chart1.Series[1].Active := true;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;

  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;

  try

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;



    if product_type = '能源公辅工序' then
    begin
      chart1.Series[1].Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PB2.NrOfColumns := Aqr1.RecordCount;
      PB2.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PB2.Elem[i, j] := 0; //  elem[列，行]       PB2 行 工序  列 产品

      FlatProgressBar1.Position := 20;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      for i := 1 to Aqr1.RecordCount do //产品
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do //    '能源公辅工序'
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PB2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;


      FlatProgressBar1.Position := 30;

      I_PB2.NrOfColumns := PB2.NrOfColumns;
      I_PB2.NrOfRows := PB2.NrOfRows;



      for i := 1 to PB2.NrOfColumns do // I-PB2
        for j := 1 to PB2.NrOfRows do
          if i = j then
            I_PB2.Elem[i, j] := 1 - PB2.Elem[i, j]
          else
            I_PB2.Elem[i, j] := -PB2.Elem[i, j];


      I_PB2.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆



      I_PB2_I.NrOfColumns := I_PB2.NrOfColumns;
      I_PB2_I.NrOfRows := I_PB2.NrOfRows;

      for i := 1 to I_PB2.NrOfColumns do // (I-PB2)-1-I
        for j := 1 to I_PB2.NrOfRows do
          if i = j then
            I_PB2_I.Elem[i, j] := I_PB2.Elem[i, j] - 1
          else
            I_PB2_I.Elem[i, j] := I_PB2.Elem[i, j];


      FlatProgressBar1.Position := 40;

      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCIfactor ');
      Aqr3.Open;

      g1_e.NrOfColumns := 1; //53 选1
      g1_e.NrOfRows := Aqr1.RecordCount;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr3.Active then Aqr3.Close;
        Aqr3.SQL.Clear;
        Aqr3.SQL.Add('select * from LCA where  product = ' + #39 + Aqr1.FieldByName('pd_name').AsString + #39);
        Aqr3.Open;
        for j := 1 to Aqr3.RecordCount do
        begin
          Aqr3.RecNo := j;
          if Aqr3.FieldByName('LCIfactor').AsString = factor then
            g1_e.Elem[1, i] := Aqr3.FieldByName('g1').AsFloat; //  elem[列，行]
        end;
      end;
      FlatProgressBar1.Position := 50;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to Aqr1.RecordCount do
          begin
            Aqr1.RecNo := j;
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr1.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := I_PB2_I.Elem[i, j];
            ADOQuery7.FieldByName('EF').AsFloat := g1_e.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;
    end; //nengyuan



    if product_type = '主生产工序' then
    begin
      chart1.Series[1].Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      FlatProgressBar1.Position := 20;

      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;

      FlatProgressBar1.Position := 30;

      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;
      for i := 1 to I_PA1.NrOfRows do // I-PA1
        for j := 1 to I_PA1.NrOfColumns do
          I_PA1.Elem[j, i] := 0;

      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


      FlatProgressBar1.Position := 40;


      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      PA2.NrOfColumns := Aqr1.RecordCount; // 主产品
      PA2.NrOfRows := Aqr4.RecordCount; //  能源


      for i := 1 to PA2.NrOfColumns do
        for j := 1 to PA2.NrOfRows do
          PA2.Elem[i, j] := 0;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 50;

      PA2_WQ.NrOfColumns := PA2.NrOfColumns;
      PA2_WQ.NrOfRows := PA2.NrOfRows;

      PA2.Multiply(I_PA1, PA2_WQ);
      PA2_WQ.Transpose;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      PE2.NrOfRows := Aqr4.RecordCount;
      PE2.NrOfColumns := 1;


      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39);
        Aqr1.Open;

        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          if Aqr1.FieldByName('LCIfactor').AsString = factor then
            PE2.Elem[1, i] := Aqr1.FieldByName('g_site').AsFloat;
        end;
      end;

      FlatProgressBar1.Position := 60;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

//          PA2.NrOfColumns := Aqr1.RecordCount; // 主产品
//      PA2.NrOfRows := Aqr4.RecordCount; //  能源
      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        if Aqr1.FieldByName('process').AsString = process then
          for i := 1 to Aqr4.RecordCount do
          begin
            Aqr4.RecNo := i;
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr4.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := PA2_WQ.Elem[i, j];
            ADOQuery7.FieldByName('EF').AsFloat := PE2.Elem[1, i];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
      end;
    end;


    FlatProgressBar1.Position := 70;

    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;


    chart1.Series[1].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    FlatProgressBar1.Position := 80;
    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[1].Add(a1 - a2, '其它', clteecolor);

    end; //  >8

    chart1.Title.Text.Clear;
    chart1.Title.Text.Add(product1 + '生命周期' + factor + '负荷构成  （自产能源带入负荷）');
    CheckBox6.Checked := true;
    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);
  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;


  StatusBar1.Panels[3].Text := '  ';
  StatusBar1.Refresh;


end;

procedure TForm1.bsSkinButton61Click(Sender: TObject); //g3
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := false;
  chart1.Series[1].Active := true;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;

  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;
  FlatProgressBar1.Position := 10;
  try
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;

    if product_type = '能源公辅工序' then
    begin
      chart1.Series[1].Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;
      RB2.NrOfRows := Aqr1.RecordCount;


      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '原材料' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '副产品或固体废弃物' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      listbox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = true then
          if listbox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
            listbox1.Items.Add(Aqr4.FieldByName('pd_name').AsString);
      end;

      RB2.NrOfColumns := listbox1.Items.Count;
      FlatProgressBar1.Position := 20;

      for i := 1 to RB2.NrOfColumns do
        for j := 1 to RB2.NrOfRows do
          RB2.Elem[i, j] := 0; //  elem[列，行]

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 0 to listbox1.Items.Count - 1 do
        begin
          if Aqr4.Active then Aqr4.Close;
          Aqr4.SQL.Clear;
          Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);

          Aqr4.Open;
          RB2.Elem[j + 1, i] := Aqr4.FieldByName('co_num').AsFloat;
        end;
      end; //RB2

      FlatProgressBar1.Position := 30;
      gBCL.NrOfRows := listbox1.Items.Count;
      gBCL.NrOfColumns := 1;

      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null ');
      Aqr4.Open;

      for i := 0 to listbox1.Items.Count - 1 do
      begin
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr4.FieldByName('pd_name').AsString = listbox1.Items[i] then
          begin
            if Aqr1.Active then Aqr1.Close;
            Aqr1.SQL.Clear;
            Aqr1.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39);

            Aqr1.Open;

            for k := 1 to Aqr1.RecordCount do
            begin
              Aqr1.RecNo := k;
              if Aqr1.FieldByName('LCIfactor').AsString = factor then
                gBCL.Elem[1, i + 1] := Aqr1.FieldByName('sum').AsFloat;
            end;
          end;
        end;
      end;
      FlatProgressBar1.Position := 40;


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to listbox1.Items.Count do
          begin
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := listbox1.Items[j - 1];
            ADOQuery7.FieldByName('AF').AsFloat := RB2.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := gBCL.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;

      FlatProgressBar1.Position := 70;

    end;



    if product_type = '主生产工序' then
    begin
      chart1.Series[1].Clear;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;


      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCIfactor ');
      Aqr3.Open;

      g1.NrOfColumns := 1; //53 选1
      g1.NrOfRows := Aqr1.RecordCount;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr3.Active then Aqr3.Close;
        Aqr3.SQL.Clear;
        Aqr3.SQL.Add('select * from LCA where  product = ' + #39 + Aqr1.FieldByName('pd_name').AsString + #39);
        Aqr3.Open;
        for j := 1 to Aqr3.RecordCount do
        begin
          Aqr3.RecNo := j;
          if Aqr3.FieldByName('LCIfactor').AsString = factor then
            g1.Elem[1, i] := Aqr3.FieldByName('g1').AsFloat; //  elem[列，行]
        end;
      end; //g1
      FlatProgressBar1.Position := 20;



      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;

      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;
      FlatProgressBar1.Position := 30;


      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆



      I_PA1_I.NrOfColumns := I_PA1.NrOfColumns;
      I_PA1_I.NrOfRows := I_PA1.NrOfRows;

      for i := 1 to I_PA1.NrOfColumns do // (I-PA1)-1-I
        for j := 1 to I_PA1.NrOfRows do
          if i = j then
            I_PA1_I.Elem[i, j] := I_PA1.Elem[i, j] - 1
          else
            I_PA1_I.Elem[i, j] := I_PA1.Elem[i, j];

   //g1.NrOfColumns := 1;            //53 选1
  //  g1.NrOfRows := Aqr1.RecordCount;
      FlatProgressBar1.Position := 40;
      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        if Aqr1.FieldByName('process').AsString = process then
          for i := 1 to Aqr4.RecordCount do
          begin
            Aqr4.RecNo := i;
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr4.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := I_PA1_I.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := g1.Elem[1, i];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
      end;

      FlatProgressBar1.Position := 70;


    end;



    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;

    FlatProgressBar1.Position := 80;
    chart1.Series[1].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[1].Add(a1 - a2, '其它', clteecolor);
    end; //  >8

    FlatProgressBar1.Position := 90;

    chart1.Title.Text.Clear;
    chart1.Title.Text.Add(product1 + '生命周期' + factor + '负荷构成（中间产品带入负荷）');
    CheckBox6.Checked := true;

    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;

    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);
  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;

end;

procedure TForm1.bsSkinButton62Click(Sender: TObject); //g4
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4, Aqr5: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := false;
  chart1.Series[1].Active := true;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;
  Aqr5 := TAdoQuery.Create(nil);
  Aqr5.Connection := ADOConnection3;


  FlatProgressBar1.Position := 10;
  try
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;

    if product_type = '能源公辅工序' then
    begin
      chart1.Series[1].Clear;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PB2.NrOfColumns := Aqr1.RecordCount;
      PB2.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PB2.Elem[i, j] := 0; //  elem[列，行]       PB2 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      for i := 1 to Aqr1.RecordCount do //产品
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do //    '能源公辅工序'
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PB2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;

      FlatProgressBar1.Position := 20;


      I_PB2.NrOfColumns := PB2.NrOfColumns;
      I_PB2.NrOfRows := PB2.NrOfRows;



      for i := 1 to PB2.NrOfColumns do // I-PB2
        for j := 1 to PB2.NrOfRows do
          if i = j then
            I_PB2.Elem[i, j] := 1 - PB2.Elem[i, j]
          else
            I_PB2.Elem[i, j] := -PB2.Elem[i, j];


      I_PB2.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;
      RB.NrOfColumns := Aqr1.RecordCount;


      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
      Aqr4.Open;

      listbox1.Clear; //分配表没有设置工序类别，所以这里辨识一下。
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
          if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
            listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
      end;
      FlatProgressBar1.Position := 30;

    //判断副产品是内部使用的   某工序产出的副产品，在其它工序作为输入；     名称不一致怎么办？

      for i := listbox1.Items.Count - 1 downto 0 do
      begin
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + listbox1.Items[i] + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
        Aqr4.Open;
        if Aqr4.RecordCount = 0 then
          listbox1.Items.Delete(i);
      end;


      RB.NrOfRows := listbox1.Items.Count;

      for i := 1 to RB.NrOfColumns do
        for j := 1 to RB.NrOfRows do
          RB.Elem[i, j] := 0; //  elem[列，行]


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 0 to listbox1.Items.Count - 1 do
        begin
          if Aqr4.Active then Aqr4.Close;
          Aqr4.SQL.Clear;
          Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
          Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
          Aqr4.Open;
          RB.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;
        end;
      end; //RB
      FlatProgressBar1.Position := 40;

 // RA*(I-PB2)-1

      RB_WQ.NrOfColumns := RB.NrOfColumns;
      RB_WQ.NrOfRows := RB.NrOfRows;

      RB.Multiply(I_PB2, RB_WQ);

      RB_WQ.Transpose; //

      RG2.NrOfRows := listbox1.Items.Count;
      RG2.NrOfColumns := 1;

      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
      Aqr4.Open;

      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open; //     if Aqr1.Locate('LCIfactor',Aqr2.FieldByName('inf_name').AsString ,[]) then

      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
        Aqr1.Open;

        if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then // ccc
        begin
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39);
          Aqr2.Open;
          for k := 1 to Aqr2.RecordCount do
          begin
            Aqr2.RecNo := k;
            if Aqr2.FieldByName('LCIfactor').AsString = factor then
              RG2.Elem[1, j + 1] := Aqr2.FieldByName('sum').AsFloat;
          end;
        end;
      end;

      FlatProgressBar1.Position := 50;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to listbox1.Items.Count do
          begin
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := listbox1.Items[j - 1];
            ADOQuery7.FieldByName('AF').AsFloat := Rb_WQ.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := RG2.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;
      FlatProgressBar1.Position := 70;
    end;



    if product_type = '主生产工序' then
    begin
      chart1.Series[1].Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 20;
      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;



      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆




      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      RA.NrOfColumns := Aqr1.RecordCount;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
      Aqr4.Open;

      listbox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
        begin
          if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
            listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
          if listbox1.Items.IndexOf(Aqr4.FieldByName('alafter').AsString) = -1 then
            listbox1.Items.Add(Aqr4.FieldByName('alafter').AsString);
        end;
      end;

        //判断副产品是内部使用的   某工序产出的副产品，在其它工序作为输入；     名称不一致怎么办？

      for i := listbox1.Items.Count - 1 downto 0 do
      begin
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + listbox1.Items[i] + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
        Aqr4.Open;
        if Aqr4.RecordCount = 0 then
          listbox1.Items.Delete(i);
      end;


      for i := listbox1.Items.Count - 1 downto 0 do
      begin
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
        Aqr4.Open;
        if Aqr4.Locate('alafter', listbox1.Items[i], []) then
          if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) <> -1 then
            if Aqr4.FieldByName('albefore').AsString <> Aqr4.FieldByName('alafter').AsString then
              Listbox1.Items.Delete(i);
      end;




      RA.NrOfRows := listbox1.Items.Count;

      for i := 1 to RA.NrOfColumns do
        for j := 1 to RA.NrOfRows do
          RA.Elem[i, j] := 0; //  elem[列，行]


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 0 to listbox1.Items.Count - 1 do
        begin
          if Aqr4.Active then Aqr4.Close;
          Aqr4.SQL.Clear;
          Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
          Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
          Aqr4.Open;
          if Aqr4.RecordCount = 0 then //分配后的名称在其它工序使用的情况
          begin
            Aqr2.Close;
            Aqr2.SQL.Clear;
            Aqr2.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and alafter=' + #39 + listbox1.Items[j] + #39);
            Aqr2.SQL.Add(' and process = ' + #39 + Aqr1.FieldByName('process').AsString + #39);
            Aqr2.Open;
            if Aqr2.RecordCount > 0 then
            begin
              Aqr4.Close;
              Aqr4.SQL.Clear;
              Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr2.FieldByName('albefore').AsString + #39);
              Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);

              Aqr4.Open;
            end;
          end;


          RA.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;

        end;
      end; //RA
      FlatProgressBar1.Position := 30;



      RA_WQ.NrOfColumns := RA.NrOfColumns;
      RA_WQ.NrOfRows := RA.NrOfRows;

      RA.Multiply(I_PA1, RA_WQ);

      RA_WQ.Transpose;

      RG.NrOfRows := listbox1.Items.Count;
      RG.NrOfColumns := 1;

      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
      Aqr4.Open;

      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
        Aqr1.Open;

        if Aqr1.RecordCount = 0 then
        begin
          Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and alafter=' + #39 + listbox1.Items[j] + #39);
          Aqr1.Open;
        end;

        if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then // ccc
        begin
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39);
          Aqr2.Open;
          for k := 1 to Aqr2.RecordCount do
          begin
            Aqr2.RecNo := k;
            if Aqr2.FieldByName('LCIfactor').AsString = factor then
              RG.Elem[1, j + 1] := Aqr2.FieldByName('sum').AsFloat;
          end;
        end;

      end;



      FlatProgressBar1.Position := 40;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to listbox1.Items.Count do
          begin
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := listbox1.Items[j - 1];
            ADOQuery7.FieldByName('AF').AsFloat := RA_WQ.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := RG.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;


      FlatProgressBar1.Position := 70;




    end;



    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;
    FlatProgressBar1.Position := 80;

    chart1.Series[1].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[1].Add(a1 - a2, '其它', clteecolor);

    end; //  >8

    chart1.Title.Text.Clear;
    chart1.Title.Text.Add(product1 + '生命周期' + factor + '负荷构成  （副产品内部收益）');
    CheckBox6.Checked := true;
    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);

  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;

end;

procedure TForm1.bsSkinButton63Click(Sender: TObject); //g5
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;
  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := false;
  chart1.Series[1].Active := true;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;
  FlatProgressBar1.Position := 10;
  try
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;

    if product_type = '能源公辅工序' then
    begin
      chart1.Series[1].Clear;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PB2.NrOfColumns := Aqr1.RecordCount;
      PB2.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PB2.Elem[i, j] := 0; //  elem[列，行]       PB2 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      for i := 1 to Aqr1.RecordCount do //产品
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do //    '能源公辅工序'
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PB2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 20;



      I_PB2.NrOfColumns := PB2.NrOfColumns;
      I_PB2.NrOfRows := PB2.NrOfRows;



      for i := 1 to PB2.NrOfColumns do // I-PB2
        for j := 1 to PB2.NrOfRows do
          if i = j then
            I_PB2.Elem[i, j] := 1 - PB2.Elem[i, j]
          else
            I_PB2.Elem[i, j] := -PB2.Elem[i, j];


      I_PB2.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open; //所有能源工序

      Aqr4.Connection := ADOConnection3;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '能源公辅工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;
      juli_e.NrOfColumns := 4;
      juli_e.NrOfRows := Aqr4.RecordCount;
      for i := 1 to juli_e.NrOfColumns do
        for j := 1 to juli_e.NrOfRows do
          juli_e.Elem[i, j] := 0;
      FlatProgressBar1.Position := 30;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from transport ');
      Aqr2.Open;

      Aqr3.Connection := ADOConnection1;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from comparison where transport is not null ');
      Aqr3.open;

      ListBox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr3.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = true then
        begin
          if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
            ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString); //在对照表中能找到的数据，存入listbox
          if Aqr2.Locate('product', Aqr3.FieldByName('transport').AsString, []) then
            for j := 1 to 4 do
              juli_e.Elem[j, i] := Aqr2.Fields[j + 2].AsFloat;
        end;
      end; // m* 4
      FlatProgressBar1.Position := 40;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from transport_base ');
      Aqr2.Open;

      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCIfactor ');
      Aqr3.Open;

      T_base.NrOfRows := 4;
      T_base.NrOfColumns := Aqr3.RecordCount; //4种运输方式的LCI清单

      for i := 1 to T_base.NrOfColumns do
        for j := 1 to T_base.NrOfRows do
          T_base.Elem[i, j] := 0;



      for i := 1 to Aqr3.RecordCount do
      begin
        Aqr3.RecNo := i;
        if Aqr2.Locate('shuxing', Aqr3.FieldByName('inf_name').AsString, []) then
          for j := 1 to 4 do
            T_base.Elem[i, j] := Aqr2.Fields[j + 2].AsFloat;
      end; // 4* 53

      T_km_e.NrOfColumns := T_base.NrOfColumns;
      T_km_e.NrOfRows := juli_e.NrOfRows;

      juli_e.Multiply(T_base, T_km_e); // T_km  外购产品距离 带入的负荷 还需乘以运输量     m*53

     // 对有运输数据的产品的单耗    OF
      FlatProgressBar1.Position := 50;

      OF_2.NrOfColumns := Aqr1.RecordCount; //能源工序数
      OF_2.NrOfRows := juli_e.NrOfRows;
      for i := 1 to OF_2.NrOfColumns do
        for j := 1 to OF_2.NrOfRows do
          OF_2.Elem[i, j] := 0;

      for i := 1 to Aqr1.RecordCount do //每个工序
      begin
        Aqr1.RecNo := i;
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where processtype = ' + #39 + '能源公辅工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
        Aqr4.Open;

        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr4.FieldByName('process').AsString = Aqr1.FieldByName('process').AsString then
            OF_2.Elem[i, j] := Aqr4.FieldByName('co_num').AsFloat;
        end;
      end;

    //   I_PB2.Transpose;

      OF2_WQ.NrOfColumns := I_PB2.NrOfColumns;
      OF2_WQ.NrOfRows := OF_2.NrOfRows;
      OF_2.Multiply(I_PB2, OF2_WQ); //  OF2_WQ  6  x 1

      OF_2.Transpose;

      T_e.NrOfRows := OF_2.NrOfRows;
      T_e.NrOfColumns := T_km_e.NrOfColumns;
      OF_2.Multiply(T_km_e, T_e); // T 各工序直接运输负荷


      FlatProgressBar1.Position := 60;

      {
    I_PB2.Multiply(T_e, g5_e);}

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '能源公辅工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to Aqr4.RecordCount do //  6
          begin
            Aqr4.RecNo := j;
            if Aqr4.FieldByName('process').AsString = process then
            begin
              ADOQuery7.Append;
              ADOQuery7.FieldByName('lcifactor').AsString := factor;
              ADOQuery7.FieldByName('process').AsString := process;
              ADOQuery7.FieldByName('pd_name').AsString := Aqr4.FieldByName('pd_name').AsString;

              ADOQuery7.FieldByName('AF').AsFloat := OF_2.Elem[j, i];
              for k := 1 to Aqr3.RecordCount do
              begin
                Aqr3.RecNo := k;
                if Aqr3.FieldByName('inf_name').AsString = factor then
                  ADOQuery7.FieldByName('EF').AsFloat := T_km_e.Elem[k, j];
              end;

              ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat / 1000;
              ADOQuery7.Post;
            end;
          end;
        end;
      end;
      FlatProgressBar1.Position := 70;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to Aqr1.RecordCount do
          begin
            Aqr1.RecNo := j;
            if i <> j then
            begin
              ADOQuery7.Append;
              ADOQuery7.FieldByName('lcifactor').AsString := factor;
              ADOQuery7.FieldByName('process').AsString := process;
              ADOQuery7.FieldByName('pd_name').AsString := Aqr1.FieldByName('pd_name').AsString;
              ADOQuery7.FieldByName('AF').AsFloat := I_PB2.Elem[i, j];
              for k := 1 to Aqr3.RecordCount do
              begin
                Aqr3.RecNo := k;
                if Aqr3.FieldByName('inf_name').AsString = factor then
                  ADOQuery7.FieldByName('EF').AsFloat := T_e.Elem[k, j];
              end;

              ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat / 1000;
              ADOQuery7.Post;
            end;
          end;
        end;
      end;
    end;




    if product_type = '主生产工序' then
    begin
      chart1.Series[1].Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 20;
      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;



      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open; //所有主工序

      Aqr4.Connection := ADOConnection3;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;
      juli.NrOfColumns := 4;
      juli.NrOfRows := Aqr4.RecordCount;
      for i := 1 to juli.NrOfColumns do
        for j := 1 to juli.NrOfRows do
          juli.Elem[i, j] := 0;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from transport ');
      Aqr2.Open;


      Aqr3.Connection := ADOConnection1;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from comparison where transport is not null ');
      Aqr3.open;

      ListBox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr3.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = true then
        begin
          if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
            ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString); //在对照表中能找到的数据，存入listbox
          if Aqr2.Locate('product', Aqr3.FieldByName('transport').AsString, []) then
            for j := 1 to 4 do
              juli.Elem[j, i] := Aqr2.Fields[j + 2].AsFloat;
        end;
      end; // m* 4
      FlatProgressBar1.Position := 30;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from transport_base ');
      Aqr2.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCIfactor ');
      Aqr3.Open;

      T_base.NrOfRows := 4;
      T_base.NrOfColumns := Aqr3.RecordCount; //4种运输方式的LCI清单

      for i := 1 to T_base.NrOfColumns do
        for j := 1 to T_base.NrOfRows do
          T_base.Elem[i, j] := 0;

      for i := 1 to Aqr3.RecordCount do
      begin
        Aqr3.RecNo := i;
        if Aqr2.Locate('shuxing', Aqr3.FieldByName('inf_name').AsString, []) then
          for j := 1 to 4 do
            T_base.Elem[i, j] := Aqr2.Fields[j + 2].AsFloat;
      end; // 4* 53

      T_km.NrOfColumns := T_base.NrOfColumns;
      T_km.NrOfRows := juli.NrOfRows;

      juli.Multiply(T_base, T_km); // T_km  外购产品距离 带入的负荷 还需乘以运输量     m*53
               //        checkmat(T_base)  ;  exit;
       // 对有运输数据的产品的单耗    OF

      OF_1.NrOfColumns := Aqr1.RecordCount;
      OF_1.NrOfRows := juli.NrOfRows;
      for i := 1 to OF_1.NrOfColumns do
        for j := 1 to OF_1.NrOfRows do
          OF_1.Elem[i, j] := 0;
      FlatProgressBar1.Position := 40;

      for i := 1 to Aqr1.RecordCount do //每个工序
      begin
        Aqr1.RecNo := i;
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from data_allocation where processtype = ' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
        Aqr4.Open;

        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr4.FieldByName('process').AsString = Aqr1.FieldByName('process').AsString then
            OF_1.Elem[i, j] := Aqr4.FieldByName('co_num').AsFloat;
        end;
      end;





      OF_1_WQ.NrOfColumns := I_PB2.NrOfColumns;
      OF_1_WQ.NrOfRows := OF_1.NrOfRows;
      I_PA1.Transpose;
      OF_1.Multiply(I_PA1, OF_1_WQ); //  OF2_WQ

      OF_1.Transpose;


      T.NrOfRows := OF_1.NrOfRows;
      T.NrOfColumns := T_km.NrOfColumns;
      for i := 1 to t.NrOfColumns do
        for j := 1 to t.NrOfRows do
        begin
          t.Elem[i, j] := 0;
        end;



      OF_1.Multiply(T_km, T); // T 各工序直接运输负荷


      FlatProgressBar1.Position := 50;


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr2.Open;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to Aqr4.RecordCount do //  6
          begin
            Aqr4.RecNo := j;
            if Aqr4.FieldByName('process').AsString = process then
            begin
              ADOQuery7.Append;
              ADOQuery7.FieldByName('lcifactor').AsString := factor;
              ADOQuery7.FieldByName('process').AsString := process;
              ADOQuery7.FieldByName('pd_name').AsString := Aqr4.FieldByName('pd_name').AsString;

              ADOQuery7.FieldByName('AF').AsFloat := OF_1.Elem[j, i];
              for k := 1 to Aqr3.RecordCount do
              begin
                Aqr3.RecNo := k;
                if Aqr3.FieldByName('inf_name').AsString = factor then
                  ADOQuery7.FieldByName('EF').AsFloat := T_km.Elem[k, j];
              end;

              ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat / 1000;
              ADOQuery7.Post;
            end;

          end;


        end;
      end;
      FlatProgressBar1.Position := 60;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to Aqr1.RecordCount do
          begin
            Aqr1.RecNo := j;
            if i <> j then
            begin
              ADOQuery7.Append;
              ADOQuery7.FieldByName('lcifactor').AsString := factor;
              ADOQuery7.FieldByName('process').AsString := process;
              ADOQuery7.FieldByName('pd_name').AsString := Aqr1.FieldByName('pd_name').AsString;
              ADOQuery7.FieldByName('AF').AsFloat := I_PA1.Elem[j, i];
              for k := 1 to Aqr3.RecordCount do
              begin
                Aqr3.RecNo := k;
                if Aqr3.FieldByName('inf_name').AsString = factor then
                  ADOQuery7.FieldByName('EF').AsFloat := T.Elem[k, j];
              end;

              ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat / 1000;
              ADOQuery7.Post;
            end;
          end;
        end;
      end;


      FlatProgressBar1.Position := 70;
      //能源使用 带入的运输负荷
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;
      PF2.NrOfRows := Aqr4.RecordCount;
      PF2.NrOfColumns := 1; //能源运输LCI结果


      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39);
        Aqr1.Open;

        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          if Aqr1.FieldByName('LCIfactor').AsString = factor then
            PF2.Elem[1, i] := Aqr1.FieldByName('g5').AsFloat;
        end;

      end;



      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      PA2.NrOfColumns := Aqr1.RecordCount; // 主产品
      PA2.NrOfRows := Aqr4.RecordCount; //  能源

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;


      PA2_WQ.NrOfColumns := PA2.NrOfColumns;
      PA2_WQ.NrOfRows := PA2.NrOfRows;
      I_PA1.Transpose;
      PA2.Multiply(I_PA1, PA2_WQ);
      PA2_WQ.Transpose;

      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;

      Aqr2.Connection := ADOConnection3;
      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr2.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to Aqr2.RecordCount do //  6
          begin
            Aqr2.RecNo := j;
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr2.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := PA2_WQ.Elem[j, i];

            ADOQuery7.FieldByName('EF').AsFloat := PF2.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;


        end;
      end;




    end;

    FlatProgressBar1.Position := 80;


    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;


    chart1.Series[1].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;
    FlatProgressBar1.Position := 90;
    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[1].Add(a1 - a2, '其它', clteecolor);

    end; //  >8

    chart1.Title.Text.Clear;
    chart1.Title.Text.Add(product1 + '生命周期' + factor + '负荷构成  （运输过程）');
    CheckBox6.Checked := true;
    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);

  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;



end;

procedure TForm1.bsSkinButton64Click(Sender: TObject); //g6
var
  i, j: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;
  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := false;
  chart1.Series[1].Active := true;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;
  FlatProgressBar1.Position := 10;
  try
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;

    if product_type = '能源公辅工序' then
    begin
      chart1.Series[1].Clear;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PB2.NrOfColumns := Aqr1.RecordCount;
      PB2.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PB2.Elem[i, j] := 0; //  elem[列，行]       PB2 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      for i := 1 to Aqr1.RecordCount do //产品
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do //    '能源公辅工序'
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PB2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 20;



      I_PB2.NrOfColumns := PB2.NrOfColumns;
      I_PB2.NrOfRows := PB2.NrOfRows;



      for i := 1 to PB2.NrOfColumns do // I-PB2
        for j := 1 to PB2.NrOfRows do
          if i = j then
            I_PB2.Elem[i, j] := 1 - PB2.Elem[i, j]
          else
            I_PB2.Elem[i, j] := -PB2.Elem[i, j];


      I_PB2.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆

      FlatProgressBar1.Position := 30;

      Aqr4.Connection := ADOConnection3;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '能源公辅工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' or material_type=' + #39 + '副产品或固体废弃物' + #39);
      Aqr1.Open;

      ListBox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
          if Aqr1.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = false then //非产品、副产品，为外购产品
            ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString);
      end;

      FlatProgressBar1.Position := 40;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '能源公辅工序' + #39 + ' and material_type = ' + #39 + '产品' + #39);
      Aqr4.Open;


      OB.NrOfColumns := Aqr4.RecordCount;
      OB.NrOfRows := ListBox1.Items.Count;



      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        for i := 0 to ListBox1.Items.Count - 1 do
        begin
          if Aqr1.Active then Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr4.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + ListBox1.Items[i] + #39);
          Aqr1.Open;
          if Aqr1.RecordCount = 0 then
            OB.Elem[j, i + 1] := 0
          else
            OB.Elem[j, i + 1] := Aqr1.FieldByName('co_num').AsFloat;
        end;

      end;


   //  OB(I-PB2)-1            L


      OB_WQ.NrOfColumns := OB.NrOfColumns;
      OB_WQ.NrOfRows := OB.NrOfRows;

      OB.Multiply(I_PB2, OB_WQ);
      OB_WQ.Transpose;
      FlatProgressBar1.Position := 50;

      OG.NrOfRows := listbox1.Items.Count;
      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null ');
      Aqr4.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open; //     if Aqr1.Locate('LCIfactor',Aqr2.FieldByName('inf_name').AsString ,[]) then
      OG.NrOfColumns := 1;


      for i := 1 to OG.NrOfColumns do
        for j := 1 to OG.NrOfRows do
          OG.Elem[i, j] := 0;

      for i := 0 to listbox1.Items.Count - 1 do
      begin
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr4.FieldByName('pd_name').AsString = listbox1.Items[i] then
          begin
            if Aqr1.Active then Aqr1.Close;
            Aqr1.SQL.Clear;
            Aqr1.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39);
            Aqr1.Open;
            if Aqr1.Locate('LCIfactor', factor, []) then
              OG.Elem[1, i + 1] := Aqr1.FieldByName('sum').AsFloat;
          end;
        end;
      end;

      FlatProgressBar1.Position := 60;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to listbox1.Items.Count do
          begin
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := listbox1.Items[j - 1];
            ADOQuery7.FieldByName('AF').AsFloat := OB_WQ.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := OG.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;
      FlatProgressBar1.Position := 70;
    end;





    if product_type = '主生产工序' then
    begin
      chart1.Series[1].Clear;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 20;
      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;



      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PA1)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


      Aqr4.Connection := ADOConnection3;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
      Aqr4.Open;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;

 //     Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' or material_type=' + #39 + '副产品或固体废弃物' + #39);
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39);
      Aqr1.Open;

      ListBox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if ListBox1.Items.IndexOf(Aqr4.FieldByName('pd_name').AsString) = -1 then
          if Aqr1.Locate('pd_name', Aqr4.FieldByName('pd_name').AsString, []) = false then //非产品、副产品，为外购产品
            ListBox1.Items.Add(Aqr4.FieldByName('pd_name').AsString);
      end;

      FlatProgressBar1.Position := 30;


      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where processtype=' + #39 + '主生产工序' + #39 + ' and material_type = ' + #39 + '产品' + #39);
      Aqr4.Open;

      OA.NrOfColumns := Aqr4.RecordCount;
      OA.NrOfRows := ListBox1.Items.Count;

      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        for i := 0 to ListBox1.Items.Count - 1 do
        begin
          if Aqr1.Active then Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr4.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + ListBox1.Items[i] + #39);
          Aqr1.Open;
          if Aqr1.RecordCount = 0 then
            OA.Elem[j, i + 1] := 0
          else
            OA.Elem[j, i + 1] := Aqr1.FieldByName('co_num').AsFloat;
        end;

      end;


        //                   OA(I-PA1)-1            L


      OA_WQ.NrOfColumns := OA.NrOfColumns;
      OA_WQ.NrOfRows := OA.NrOfRows;

      OA.Multiply(I_PA1, OA_WQ);
      OA_WQ.Transpose;
      FlatProgressBar1.Position := 40;
      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCIfactor ');
      Aqr3.Open;
      OG.NrOfColumns := 1;

      OG.NrOfRows := listbox1.Items.Count;
      for i := 1 to OG.NrOfColumns do
        for j := 1 to OG.NrOfRows do
          OG.Elem[i, j] := 0;


      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null ');
      Aqr4.Open;

      for i := 0 to listbox1.Items.Count - 1 do
      begin
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr4.FieldByName('pd_name').AsString = listbox1.Items[i] then
          begin
            if Aqr1.Active then Aqr1.Close;
            Aqr1.SQL.Clear;
            Aqr1.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39);

            Aqr1.Open;
            if Aqr1.Locate('LCIfactor', factor, []) = true then
              OG.Elem[1, i + 1] := Aqr1.FieldByName('sum').AsFloat; //上游lci
          end;

        end;
      end;
    end;
    FlatProgressBar1.Position := 50;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;


    Aqr3.Connection := ADOConnection3;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from lcifactor ');
    Aqr3.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('process').AsString = process then
      begin
        for j := 1 to listbox1.Items.Count do
        begin
          ADOQuery7.Append;
          ADOQuery7.FieldByName('lcifactor').AsString := factor;
          ADOQuery7.FieldByName('process').AsString := process;
          ADOQuery7.FieldByName('pd_name').AsString := listbox1.Items[j - 1];
          ADOQuery7.FieldByName('AF').AsFloat := OA_WQ.Elem[j, i];
          ADOQuery7.FieldByName('EF').AsFloat := OG.Elem[1, j];
          ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
          ADOQuery7.Post;
        end;
      end;
    end;





     //能源使用 带入的上游负荷
    Aqr4.Connection := ADOConnection2;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;



    PG2.NrOfRows := Aqr4.RecordCount;
    PG2.NrOfColumns := 1; //能源上游LCI结果

    for i := 1 to PG2.NrOfColumns do
      for j := 1 to PG2.NrOfRows do
        PG2.Elem[i, j] := 0;

    for i := 1 to Aqr4.RecordCount do
    begin
      Aqr4.RecNo := i;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39);
      Aqr1.Open;
      if Aqr1.Locate('LCIfactor', factor, []) = true then
        PG2.Elem[1, i] := Aqr1.FieldByName('g6').AsFloat;
    end;


    FlatProgressBar1.Position := 60;
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;

    PA2.NrOfColumns := Aqr1.RecordCount; // 主产品
    PA2.NrOfRows := Aqr4.RecordCount; //  能源
    for i := 1 to PA2.NrOfColumns do
      for j := 1 to PA2.NrOfRows do
        PA2.Elem[i, j] := 0;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
        Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
        Aqr2.Open;
        if Aqr2.RecordCount > 0 then
        begin
          Aqr2.First;
          PA2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
        end;
      end;
    end;


    PA2_WQ.NrOfColumns := PA2.NrOfColumns;
    PA2_WQ.NrOfRows := PA2.NrOfRows;

    PA2.Multiply(I_PA1, PA2_WQ);
    PA2_WQ.Transpose;



    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
    Aqr4.Open;


    Aqr3.Connection := ADOConnection3;
    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from lcifactor ');
    Aqr3.Open;

    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      if Aqr1.FieldByName('process').AsString = process then
      begin
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          ADOQuery7.Append;
          ADOQuery7.FieldByName('lcifactor').AsString := factor;
          ADOQuery7.FieldByName('process').AsString := process;
          ADOQuery7.FieldByName('pd_name').AsString := Aqr4.FieldByName('pd_name').AsString;
          ADOQuery7.FieldByName('AF').AsFloat := PA2_WQ.Elem[j, i];
          ADOQuery7.FieldByName('EF').AsFloat := PG2.Elem[1, j];
          ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
          ADOQuery7.Post;
        end;
      end;

    end;

    FlatProgressBar1.Position := 70;



    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;


    chart1.Series[1].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;
    FlatProgressBar1.Position := 80;
    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[1].Add(a1 - a2, '其它', clteecolor);

    end; //  >8

    chart1.Title.Text.Clear;

    chart1.Title.Text.Add(product1 + '生命周期' + factor + '负荷构成  （上游负荷）');
    CheckBox6.Checked := true;
    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);
  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;



end;

procedure TForm1.DBGridEh6TitleClick(Column: TColumnEh);
begin
  if flag then
    ADOQuery5.Sort := Column.FieldName + ' ASC'
  else
    ADOQuery5.Sort := Column.FieldName + ' DESC';

  flag := not flag;
end;

procedure TForm1.bsSkinButton65Click(Sender: TObject); //g7
var
  i, j: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;

  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := false;
  chart1.Series[1].Active := true;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;
  FlatProgressBar1.Position := 10;
  try
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;

    if product_type = '能源公辅工序' then
    begin
      chart1.Series[1].Clear;


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PB2.NrOfColumns := Aqr1.RecordCount;
      PB2.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PB2.Elem[i, j] := 0; //  elem[列，行]       PB2 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      for i := 1 to Aqr1.RecordCount do //产品
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do //    '能源公辅工序'
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PB2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;


      FlatProgressBar1.Position := 20;

      I_PB2.NrOfColumns := PB2.NrOfColumns;
      I_PB2.NrOfRows := PB2.NrOfRows;



      for i := 1 to PB2.NrOfColumns do // I-PB2
        for j := 1 to PB2.NrOfRows do
          if i = j then
            I_PB2.Elem[i, j] := 1 - PB2.Elem[i, j]
          else
            I_PB2.Elem[i, j] := -PB2.Elem[i, j];


      I_PB2.Invert; // (I-PB2)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆

      FlatProgressBar1.Position := 30;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;

      Aqr4.Connection := ADOConnection3;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
      Aqr4.Open;


      listbox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
        begin
          if Aqr3.Active then Aqr3.Close;
          Aqr3.SQL.Clear;
          Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('albefore').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
          Aqr3.Open;
          if Aqr3.RecordCount = 0 then
          begin
            if Aqr3.Active then Aqr3.Close;
            Aqr3.SQL.Clear;
            Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('alafter').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
            Aqr3.Open;
            if Aqr3.RecordCount = 0 then
              if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
                listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
          end;
        end;
      end;
        //  TRT电  -- 电网电     水渣－－水泥           listbox1.Items.Delete(i);


      FlatProgressBar1.Position := 40;

      RB_out.NrOfColumns := Aqr1.RecordCount;
      RB_out.NrOfRows := listbox1.Items.Count;


      for i := 1 to RA_out.NrOfColumns do
        for j := 1 to RA_out.NrOfRows do
          RB_out.Elem[i, j] := 0; //  elem[列，行]

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 0 to listbox1.Items.Count - 1 do
        begin
          if Aqr4.Active then Aqr4.Close;
          Aqr4.SQL.Clear;
          Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
          Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
          Aqr4.Open;
          RB_out.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;
        end;
      end; //RA_out

    // RA_out_WQ*(I-PA1)-1
      FlatProgressBar1.Position := 50;
      RB_out_WQ.NrOfColumns := RB_out.NrOfColumns;
      RB_out_WQ.NrOfRows := RB_out.NrOfRows;

      RB_out.Multiply(I_PB2, RB_out_WQ);

      RB_out_WQ.Transpose;

  //rg

      RG2_out.NrOfRows := listbox1.Items.Count;
      RG2_out.NrOfColumns := 1;
      for i := 1 to RG2_out.NrOfColumns do
        for j := 1 to RG2_out.NrOfRows do
          RG2_out.Elem[i, j] := 0;

      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
      Aqr4.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open; //     if Aqr1.Locate('LCIfactor',Aqr2.FieldByName('inf_name').AsString ,[]) then



      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
        Aqr1.Open;

        if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then // ccc
        begin
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39);
          Aqr2.Open;
          if Aqr2.Locate('LCIfactor', factor, []) then
            RG2_out.Elem[1, j + 1] := Aqr2.FieldByName('sum').AsFloat;
        end;
      end;
      FlatProgressBar1.Position := 60;


      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr1.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to listbox1.Items.Count do
          begin
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := listbox1.Items[j - 1];
            ADOQuery7.FieldByName('AF').AsFloat := RB_out_WQ.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := RG2_out.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;
      FlatProgressBar1.Position := 70;
    end;






    if product_type = '主生产工序' then
    begin
      chart1.Series[1].Clear;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 20;
      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;



      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PA1)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆


      FlatProgressBar1.Position := 30;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;


      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39);
      Aqr4.Open;

      listbox1.Clear;
      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Locate('process', Aqr4.FieldByName('process').AsString, []) = true then
        begin
          if Aqr3.Active then Aqr3.Close;
          Aqr3.SQL.Clear;
          Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('albefore').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
          Aqr3.Open;
          if Aqr3.RecordCount = 0 then
          begin
            if Aqr3.Active then Aqr3.Close;
            Aqr3.SQL.Clear;
            Aqr3.SQL.Add('select * from data_allocation where pd_name   = ' + #39 + Aqr4.FieldByName('alafter').AsString + #39 + '  and( material_type = ' + #39 + '原材料' + #39 + ' or material_type=' + #39 + '能源及能源介质' + #39 + ' or material_type=' + #39 + '辅助材料' + #39 + ' )');
            Aqr3.Open;
            if Aqr3.RecordCount = 0 then
              if listbox1.Items.IndexOf(Aqr4.FieldByName('albefore').AsString) = -1 then
                listbox1.Items.Add(Aqr4.FieldByName('albefore').AsString);
          end;
        end;
      end;
        //  TRT电  -- 电网电     水渣－－水泥           listbox1.Items.Delete(i);



      RA_out.NrOfColumns := Aqr1.RecordCount;
      RA_out.NrOfRows := listbox1.Items.Count;

      for i := 1 to RA_out.NrOfColumns do
        for j := 1 to RA_out.NrOfRows do
          RA_out.Elem[i, j] := 0; //  elem[列，行]

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 0 to listbox1.Items.Count - 1 do
        begin
          if Aqr4.Active then Aqr4.Close;
          Aqr4.SQL.Clear;
          Aqr4.SQL.Add('select * from data_allocation where process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + listbox1.Items[j] + #39);
          Aqr4.SQL.Add(' and material_type = ' + #39 + '副产品或固体废弃物' + #39);
          Aqr4.Open;
          RA_out.Elem[i, j + 1] := Aqr4.FieldByName('co_num').AsFloat;

        end;
      end; //RA_out

      FlatProgressBar1.Position := 40;

 // RA_out_WQ*(I-PA1)-1

      RA_out_WQ.NrOfColumns := RA_out.NrOfColumns;
      RA_out_WQ.NrOfRows := RA_out.NrOfRows;

      RA_out.Multiply(I_PA1, RA_out_WQ);
      RA_out_WQ.Transpose;

  //rg




      RG_out.NrOfRows := listbox1.Items.Count;
      RG_out.NrOfColumns := 1;


      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, upstream from comparison where upstream is not null '); //ccc---chuyanshui
      Aqr4.Open;

      Aqr3.Connection := ADOConnection2;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from LCIfactor ');
      Aqr3.Open;

      for j := 0 to listbox1.Items.Count - 1 do
      begin
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from allocation where allocation  = ' + #39 + '系统扩展法' + #39 + ' and albefore=' + #39 + listbox1.Items[j] + #39);
        Aqr1.Open;

        if Aqr4.Locate('pd_name', Aqr1.FieldByName('alafter').AsString, []) = true then // ccc
        begin
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from up where product = ' + #39 + Aqr4.FieldByName('upstream').AsString + #39);
          Aqr2.Open;
          if Aqr2.Locate('LCIfactor', factor, []) = true then
            RG_out.Elem[1, j + 1] := Aqr2.FieldByName('sum').AsFloat;

        end;
      end;

      FlatProgressBar1.Position := 50;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to listbox1.Items.Count do
          begin
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := listbox1.Items[j - 1];
            ADOQuery7.FieldByName('AF').AsFloat := RA_out_WQ.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := RG_out.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;


      FlatProgressBar1.Position := 60;

           //能源使用 带入副产品回用的收益

      Aqr4.Connection := ADOConnection3;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      PH2.NrOfRows := Aqr4.RecordCount;
      PH2.NrOfColumns := 1; //能源 g7   LCI结果
      for i := 1 to Aqr4.RecordCount do
        PH2.Elem[1, i] := 0;

      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39);
        Aqr1.Open;
        if Aqr1.Locate('lcifactor', factor, []) then
          PH2.Elem[1, i] := (-1) * Aqr1.FieldByName('g7').AsFloat;
      end;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;


      Aqr3.Connection := ADOConnection3;
      if Aqr3.Active then Aqr3.Close;
      Aqr3.SQL.Clear;
      Aqr3.SQL.Add('select * from lcifactor ');
      Aqr3.Open;

      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = process then
        begin
          for j := 1 to Aqr4.RecordCount do
          begin
            Aqr4.RecNo := j;
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr4.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := PA2_WQ.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := PH2.Elem[1, j];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
        end;
      end;

      FlatProgressBar1.Position := 70;
    end;





    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[1].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;

    FlatProgressBar1.Position := 80;
    chart1.Series[1].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[1].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[1].Add(a1 - a2, '其它', clteecolor);

    end; //  >8

    chart1.Title.Text.Clear;

    chart1.Title.Text.Add(product1 + '生命周期' + factor + '负荷构成  （副产品外部收益）');
    CheckBox6.Checked := true;

    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);

  finally
    FlatProgressBar1.Position := 100;
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;



end;

procedure TForm1.DBGridEh9TitleClick(Column: TColumnEh);
begin
  if flag then
    ADOQuery1.Sort := Column.FieldName + ' ASC'
  else
    ADOQuery1.Sort := Column.FieldName + ' DESC';

  flag := not flag;
end;

procedure TForm1.bsSkinButton69Click(Sender: TObject); //Lcc
var
  mysql1, s, s0, product: string;
  Aqr1: TAdoQuery;
  i, j: integer;
  a1, a2: double;
begin
  bsSkinEdit1.Text := '';
  if checkedcount(bsSkinCheckListBox8) = 0 then
  begin
    bsSkinMessage1.MessageDlg('未选择项目', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox5) = 0 then
  begin
    bsSkinMessage1.MessageDlg('请至少选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;



  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := false;
  chart1.Series[0].Active := true;
  chart1.Series[1].Active := false;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s0 := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s0 + '\' + s0 + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr1 := TAdoQuery.Create(nil);



  try

    for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
    begin
      if bsSkinCheckListBox8.Checked[i] = true then
      begin
        s := bsSkinCheckListBox8.Items[i];
        ADOConnection4.Close;
        ADOConnection4.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
        ADOConnection4.Open;
        Aqr1.Connection := ADOConnection4;

        for j := 0 to bsSkinCheckListBox5.Items.Count - 1 do
        begin
          if bsSkinCheckListBox5.Checked[j] = true then
          begin
            product := bsSkinCheckListBox5.Items[j];
            mysql1 := 'select * from LCA where product= ' + #39 + product + #39 + ' and lcifactor=' + #39 + '生命周期成本' + #39;
            Aqr1.Close;
            Aqr1.SQL.Clear;
            Aqr1.SQL.Add(mysql1);
            Aqr1.Open;

            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := '生命周期成本';
            ADOQuery7.FieldByName('process').AsString := s;
            ADOQuery7.FieldByName('pd_name').AsString := product;
            ADOQuery7.FieldByName('AF').AsFloat := 0;
            ADOQuery7.FieldByName('EF').AsFloat := 0;
            if ComboBox6.ItemIndex = -1 then
              ADOQuery7.FieldByName('AFxEF').AsFloat := Aqr1.FieldByName('g_route').AsFloat
            else
              ADOQuery7.FieldByName('AFxEF').AsFloat := Aqr1.Fields[ComboBox6.ItemIndex + 5].AsFloat;

            ADOQuery7.Post;


          end;
        end;
      end;
    end;

    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;


    chart1.Series[0].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('process').AsString + '-' + ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 < 0.0000000001 then exit;
      chart1.Series[0].Add(a1 - a2, '其它', clteecolor);
    end; //  >8

    chart1.Title.Text.Clear;


    if ComboBox6.ItemIndex = -1 then
      chart1.Title.Text.Add('生命周期成本比较(元/千克 或 元/立方米)')
    else
      chart1.Title.Text.Add(ComboBox6.Text + '-生命周期成本比较(元/千克 或 元/立方米)');


  finally
    Aqr1.Close;
    Aqr1.Free;
  end;
end;

procedure TForm1.bsSkinButton68Click(Sender: TObject);
var
  mysql1, s, s0, factor, product: string;
  Aqr1: TAdoQuery;
  i, j: integer;
  a1, a2: double;
begin
  bsSkinEdit1.Text := '';
  if checkedcount(bsSkinCheckListBox8) = 0 then
  begin
    bsSkinMessage1.MessageDlg('未选择项目', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox5) = 0 then
  begin
    bsSkinMessage1.MessageDlg('请至少选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := false;
  chart1.Series[0].Active := true;
  chart1.Series[1].Active := false;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s0 := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;


  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s0 + '\' + s0 + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr1 := TAdoQuery.Create(nil);


  try

    for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
    begin
      if bsSkinCheckListBox8.Checked[i] = true then
      begin
        s := bsSkinCheckListBox8.Items[i];
        ADOConnection4.Close;
        ADOConnection4.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
        ADOConnection4.Open;
        Aqr1.Connection := ADOConnection4;

        for j := 0 to bsSkinCheckListBox5.Items.Count - 1 do
        begin
          if bsSkinCheckListBox5.Checked[j] = true then
          begin
            product := bsSkinCheckListBox5.Items[j];
            mysql1 := 'select * from LCA where product= ' + #39 + product + #39 + ' and lcifactor=' + #39 + factor + #39;
            Aqr1.Close;
            Aqr1.SQL.Clear;
            Aqr1.SQL.Add(mysql1);
            Aqr1.Open;

            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := s;
            ADOQuery7.FieldByName('pd_name').AsString := product;
            ADOQuery7.FieldByName('AF').AsFloat := 0;
            ADOQuery7.FieldByName('EF').AsFloat := 0;
            if ComboBox6.ItemIndex = -1 then
              ADOQuery7.FieldByName('AFxEF').AsFloat := Aqr1.FieldByName('g_route').AsFloat
            else
              ADOQuery7.FieldByName('AFxEF').AsFloat := Aqr1.Fields[ComboBox6.ItemIndex + 5].AsFloat;

            ADOQuery7.Post;


          end;
        end;
      end;
    end;




    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;


    chart1.Series[0].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('process').AsString + '-' + ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[0].Add(a1 - a2, '其它', clteecolor);
    end; //  >8

    chart1.Title.Text.Clear;
    if ComboBox6.ItemIndex = -1 then
      chart1.Title.Text.Add('生命周期' + '-' + factor + '-比较 ')
    else
      chart1.Title.Text.Add(ComboBox6.Text + '-' + factor + '-比较 ');




  finally
    Aqr1.Close;
    Aqr1.Free;
  end;




end;

procedure TForm1.bsSkinButton66Click(Sender: TObject); //分布1
var
  i, j: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := true;
  chart1.Series[1].Active := false;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;
  FlatProgressBar1.Position := 10;
  try
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;

    if product_type = '能源公辅工序' then
    begin
      bsSkinMessage1.MessageDlg('此分布仅针对主工序产品', (mtInformation), [mbOK], 0);
      exit;
    end;

    if product_type = '主生产工序' then
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '能源公辅工序' + #39);
      Aqr4.Open;

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;

      PA2.NrOfColumns := Aqr1.RecordCount;
      PA2.NrOfRows := Aqr4.RecordCount;

      for i := 1 to PA2.NrOfColumns do
        for j := 1 to PA2.NrOfRows do
          PA2.Elem[i, j] := 0;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
     //       showmessage( Aqr2.SQL.Text)     ;
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA2.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end; // PA2     各主工序对自产能源的直接消耗系数矩阵       19x25

      checkmat(PA2);


      FlatProgressBar1.Position := 20;
      PE2_1.NrOfRows := 1;
      PE2_1.NrOfColumns := Aqr4.RecordCount;

      for i := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := i;
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from LCA where  product = ' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and LCIFactor=' + #39 + factor + #39);
        Aqr2.Open;
        Aqr2.First;
        PE2_1.Elem[i, 1] := Aqr2.FieldByName('g_site').AsFloat;

      end; // PE2_1        1x19

      g2_unit.NrOfColumns := PA2.NrOfColumns;
      g2_unit.NrOfRows := 1; //各单元过程的直接 能源负荷g2

      PE2_1.Multiply(PA2, g2_unit);

    //   checkmat(g2_unit);

      FlatProgressBar1.Position := 30;
      g1_unit.NrOfColumns := PA2.NrOfColumns;
      g1_unit.NrOfRows := 1;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from LCA where  product = ' + #39 + Aqr1.FieldByName('pd_name').AsString + #39 + ' and LCIFactor=' + #39 + factor + #39);
        Aqr2.Open;
        Aqr2.First;
        g1_unit.Elem[i, 1] := Aqr2.FieldByName('g1').AsFloat;
      end;


      FlatProgressBar1.Position := 40;

      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 50;
      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;

      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PA1)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆







      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        if Aqr1.FieldByName('process').AsString = process then
          for i := 1 to Aqr1.RecordCount do
          begin
            Aqr1.RecNo := i;
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr1.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := I_PA1.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := g1_unit.Elem[i, 1] + g2_unit.Elem[i, 1];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
      end;
    end;

    FlatProgressBar1.Position := 70;


    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;
    FlatProgressBar1.Position := 80;

    chart1.Series[0].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[0].Add(a1 - a2, '其它', clteecolor);
    end; //  >8

    chart1.Title.Text.Clear;
    chart1.Title.Text.Add(product1 + '  [' + factor + ']  在各工序的分布1');
    CheckBox6.Checked := true;
    CheckBox6Click(Sender);

    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);





  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;




end;

procedure TForm1.bsSkinButton67Click(Sender: TObject);
var
  i, j: integer;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  s, product1, process, product_type, factor: string;
  a1, a2: double;
begin
  if checkedcount(bsSkinCheckListBox5) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个产品', (mtInformation), [mbOK], 0);
    exit;
  end;

  if checkedcount(bsSkinCheckListBox55) <> 1 then
  begin
    bsSkinMessage1.MessageDlg('请选择一个环境指标', (mtInformation), [mbOK], 0);
    exit;
  end;
  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;

  dxFlowChart2.Visible := false;
  chart1.Visible := true;
  CheckBox6.Visible := true;
  chart1.Series[0].Active := true;
  chart1.Series[1].Active := false;


  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
  begin
    if bsSkinCheckListBox55.Checked[i] = true then
    begin
      factor := bsSkinCheckListBox55.Items[i];
      break;
    end;
  end;

  ADOConnection3.Close;
  ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection3.Open;
  ADOQuery7.Connection := form1.ADOConnection3;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('delete * from temp');
  ADOQuery7.ExecSQL;

  if ADOQuery7.Active then ADOQuery7.Close;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('select * from temp');
  ADOQuery7.Open;

  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection3;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection3;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection3;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection3;

  FlatProgressBar1.Position := 10;
  try
    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr4.Open;
    process := Aqr4.FieldByName('process').AsString;
    product_type := Aqr4.FieldByName('processtype').AsString;

    if product_type = '能源公辅工序' then
    begin
      bsSkinMessage1.MessageDlg('此分布仅针对主工序产品', (mtInformation), [mbOK], 0);
      exit;
    end;

    if product_type = '主生产工序' then
    begin

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr1.Open;
      if Aqr1.RecordCount = 0 then
      begin
        bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
        exit;
      end;

      g1_unit.NrOfColumns := Aqr1.RecordCount;
      g1_unit.NrOfRows := 1;
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr4.Active then Aqr4.Close;
        Aqr4.SQL.Clear;
        Aqr4.SQL.Add('select * from LCA where  product = ' + #39 + Aqr1.FieldByName('pd_name').AsString + #39 + ' and LCIFactor=' + #39 + factor + #39);
        Aqr4.Open;
        Aqr4.First;
        g1_unit.Elem[i, 1] := Aqr4.FieldByName('g1').AsFloat;
      end;


      FlatProgressBar1.Position := 30;


      PA1.NrOfColumns := Aqr1.RecordCount;
      PA1.NrOfRows := Aqr1.RecordCount;
      for i := 1 to Aqr1.RecordCount do
        for j := 1 to Aqr1.RecordCount do
          PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
      Aqr4.Open;


      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        for j := 1 to Aqr4.RecordCount do
        begin
          Aqr4.RecNo := j;
          if Aqr2.Active then Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
          Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
          Aqr2.Open;
          if Aqr2.RecordCount > 0 then
          begin
            Aqr2.First;
            PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
          end;
        end;
      end;
      FlatProgressBar1.Position := 40;
      I_PA1.NrOfColumns := PA1.NrOfColumns;
      I_PA1.NrOfRows := PA1.NrOfRows;

      for i := 1 to PA1.NrOfColumns do // I-PA1
        for j := 1 to PA1.NrOfRows do
          if i = j then
            I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
          else
            I_PA1.Elem[i, j] := -PA1.Elem[i, j];


      I_PA1.Invert; // (I-PA1)-1     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆

      FlatProgressBar1.Position := 50;


      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        if Aqr1.FieldByName('process').AsString = process then
          for i := 1 to Aqr1.RecordCount do
          begin
            Aqr1.RecNo := i;
            ADOQuery7.Append;
            ADOQuery7.FieldByName('lcifactor').AsString := factor;
            ADOQuery7.FieldByName('process').AsString := process;
            ADOQuery7.FieldByName('pd_name').AsString := Aqr1.FieldByName('pd_name').AsString;
            ADOQuery7.FieldByName('AF').AsFloat := I_PA1.Elem[j, i];
            ADOQuery7.FieldByName('EF').AsFloat := g1_unit.Elem[i, 1];
            ADOQuery7.FieldByName('AFxEF').AsFloat := ADOQuery7.FieldByName('AF').AsFloat * ADOQuery7.FieldByName('EF').AsFloat;
            ADOQuery7.Post;
          end;
      end;
    end;
    FlatProgressBar1.Position := 60;


    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add('select * from LCA where  product = ' + #39 + product1 + #39 + ' and lcifactor=' + #39 + factor + #39);
    Aqr3.Open;
    if Aqr3.RecordCount > 0 then
    begin
      ADOQuery7.Append;
      ADOQuery7.FieldByName('lcifactor').AsString := factor;
      ADOQuery7.FieldByName('process').AsString := process;
      ADOQuery7.FieldByName('pd_name').AsString := '能源系统';
      ADOQuery7.FieldByName('AF').AsFloat := 0; ;
      ADOQuery7.FieldByName('EF').AsFloat := 0;
      ADOQuery7.FieldByName('AFxEF').AsFloat := Aqr3.FieldByName('g2').AsFloat;
      ADOQuery7.Post;
    end;

    ADOQuery7.Close;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('select * from temp ');
    ADOQuery7.SQL.Add('order by AFxEF desc');
    ADOQuery7.Open;

    if ADOQuery7.RecordCount = 0 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('No data', (mtInformation), [mbOK], 0);
      exit;
    end;

    ADOQuery7.First;
    if ADOQuery7.Fields[6].AsFloat < 0.00000000000001 then
    begin
      chart1.Series[0].Clear;
      bsSkinMessage1.MessageDlg('无数据或数据全为零', (mtInformation), [mbOK], 0);
      exit;
    end;
    FlatProgressBar1.Position := 80;

    chart1.Series[0].Clear;
    if ADOQuery7.RecordCount <= 8 then
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

    a1 := 0;
    a2 := 0;
    if ADOQuery7.RecordCount > 8 then
    begin
      for i := 1 to ADOQuery7.RecordCount do
      begin
        ADOQuery7.RecNo := i;
        a1 := a1 + ADOQuery7.Fields[6].AsFloat;
      end;

      for i := 1 to 8 do
      begin
        ADOQuery7.RecNo := i;
        a2 := a2 + ADOQuery7.Fields[6].AsFloat;
        if ADOQuery7.Fields[6].AsFloat < 0.0000000001 then break;
        chart1.Series[0].Add(ADOQuery7.Fields[6].AsFloat, ADOQuery7.FieldByName('pd_name').AsString, clteecolor);
      end;

      if a1 - a2 > 0.0000000001 then
        chart1.Series[0].Add(a1 - a2, '其它', clteecolor);
    end; //  >8



    chart1.Title.Text.Clear;
    chart1.Title.Text.Add(product1 + '  [' + factor + ']   在各工序的分布2');
    CheckBox6.Checked := true;
    CheckBox6Click(Sender);

    a1 := 0;
    for i := 1 to ADOQuery7.RecordCount do
    begin
      ADOQuery7.RecNo := i;
      a1 := a1 + ADOQuery7.Fields[6].AsFloat;
    end;
    bsSkinEdit1.Text := '合计：' + formatfloat('###,##0.##0', a1);





  finally
    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '  ';
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;




end;

procedure TForm1.CheckBox7Click(Sender: TObject);
var
  i: integer;
begin
  if CheckBox7.Checked = true then
    for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
      bsSkinCheckListBox8.Checked[i] := true;
  if CheckBox7.Checked = false then
    for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
      bsSkinCheckListBox8.Checked[i] := false;
  bsSkinCheckListBox8ClickCheck(Sender);


end;

procedure TForm1.CheckBox8Click(Sender: TObject);
var
  i: integer;
begin
  if CheckBox8.Checked = true then
    for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
      bsSkinCheckListBox5.Checked[i] := true;
  if CheckBox8.Checked = false then
    for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
      bsSkinCheckListBox5.Checked[i] := false;

end;

procedure TForm1.N18Click(Sender: TObject);
begin
  bsSkinButton1Click(Sender);
end;

procedure TForm1.N19Click(Sender: TObject);
begin
  bsSkinButton48Click(Sender);
end;

procedure TForm1.N20Click(Sender: TObject);
begin
  bsSkinButton1Click(Sender);
end;

procedure TForm1.N21Click(Sender: TObject);
begin
  bsSkinButton1Click(Sender);
end;

procedure TForm1.LCA4Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 7;
  PageControl1Change(Sender);
end;

procedure TForm1.DBGridEh7TitleClick(Column: TColumnEh);
begin
  if flag then
    ADOQuery3.Sort := Column.FieldName + ' ASC'
  else
    ADOQuery3.Sort := Column.FieldName + ' DESC';

  flag := not flag;
end;

procedure TForm1.DBGridEh8TitleClick(Column: TColumnEh);
begin
  if flag then
    ADOQuery6.Sort := Column.FieldName + ' ASC'
  else
    ADOQuery6.Sort := Column.FieldName + ' DESC';

  flag := not flag;
end;

procedure TForm1.bsSkinButton25Click(Sender: TObject);
var
  i: integer;
begin
  if DBGridEh4.SelectedRows.Count = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('请选择需要删除的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if form1.bsSkinMessage1.MessageDlg('确定删除所选记录？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := 0 to DBGridEh4.SelectedRows.Count - 1 do
    begin
      DBGridEh4.DataSource.DataSet.GotoBookmark(pointer(DBGridEh4.SelectedRows[i]));
      DBGridEh4.DataSource.DataSet.Delete;
    end;
  end;

  ADOQuery8.UpdateBatch(arall);

end;

procedure TForm1.DBGridEh4GetCellParams(Sender: TObject; Column: TColumnEh;
  AFont: TFont; var Background: TColor; State: TGridDrawState);
begin
  if not (gdSelected in State) then
    if (Sender as TDBGridEh).SumList.RecNo mod 2 = 0 then
      Background := clInfobk;
end;

procedure TForm1.ADOQuery8AfterOpen(DataSet: TDataSet);
var
  i: integer;
begin
  ADOQuery8.FieldByName('ID').OnGetText := SetFdText8;
  for i := 1 to ADOQuery8.FieldCount - 1 do
    if ADOQuery8.Fields[i].DataType = ftFloat then
      (ADOQuery8.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0';
  ADOQuery8.First;
end;

procedure TForm1.bsSkinButton70Click(Sender: TObject);
var
  i: integer;
begin
  if DBGridEh2.SelectedRows.Count = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('请选择需要删除的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if form1.bsSkinMessage1.MessageDlg('确定删除所选记录？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := 0 to DBGridEh2.SelectedRows.Count - 1 do
    begin
      DBGridEh2.DataSource.DataSet.GotoBookmark(pointer(DBGridEh2.SelectedRows[i]));
      DBGridEh2.DataSource.DataSet.Delete;
    end;
  end;

  ADOQuery1.UpdateBatch(arall);

end;

procedure TForm1.Excel1Click(Sender: TObject);
var
  v1: Variant;
  i, j, k: integer;
  s: string;
 // a, b: double;
  Aqr1, Aqr2: TAdoQuery;
begin
  if form1.bsSkinMessage1.MessageDlg('载入数据将覆盖原有数据，确认载入？', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    exit;

  if OpenDialog1.Execute then
  begin

    try
      v1 := CreateOleObject('Excel.Application'); //创建ole对象
    except
      showmessage('未安装EXCEL');
    end;

    try
      v1.WorkBooks.open(OpenDialog1.FileName);
      s := v1.WorkBooks[1].WorkSheets[1].name; //下载
    finally
      v1.WorkBooks[1].Close;
      v1.quit;
      v1 := Unassigned;
    end;


    try

      ADOConExcel1.Close;
      ADOConExcel1.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + OpenDialog1.FileName + ';Extended Properties=excel 8.0;Persist Security Info=False ';
      ADOConExcel1.Open;

      Aqr1 := TAdoQuery.Create(nil);
      Aqr1.Connection := ADOConExcel1;
      Aqr2 := TAdoQuery.Create(nil);
      Aqr2.Connection := ADOConnection2; ;

      StatusBar1.Panels[3].Text := '正在导入';
      StatusBar1.Refresh;

      FlatProgressBar1.Position := 0;
      FlatProgressBar1.Max := ListBox2.Items.Count - 1;
      for i := 1 to ListBox2.Items.Count - 1 do
      begin
        FlatProgressBar1.Position := i;
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('SELECT * FROM [' + s + '$]' + ' where process=' + #39 + ListBox2.Items[i] + #39);
        Aqr1.Open;




        if Aqr1.RecordCount > 0 then
        begin
          Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('delete * from data where process=' + #39 + ListBox2.Items[i] + #39);
          Aqr2.ExecSQL;

          Aqr2.Close;
          Aqr2.SQL.Clear;
          Aqr2.SQL.Add('select * from data ');
          Aqr2.Open;

          for j := 1 to Aqr1.RecordCount do
          begin
            Aqr1.RecNo := j;
            Aqr2.Append;
            for k := 1 to Aqr2.FieldCount - 1 do
              Aqr2.Fields[k].AsString := Aqr1.Fields[k].AsString;
            Aqr2.Post;
          end;
        end;
      end;
      ProcessData_Format;

      form1.bsSkinMessage1.MessageDlg('载入完成！', (mtInformation), [mbOK], 0);
      StatusBar1.Panels[3].Text := ' ';
      FlatProgressBar1.Position := 0;
      StatusBar1.Refresh;

    finally
      ADOConExcel1.Close;
      Aqr1.Close;
      Aqr1.Free;
      Aqr2.Close;
      Aqr2.Free;
    end;

  end;

end;

procedure TForm1.N23Click(Sender: TObject);
var
  s: TStringList;

begin
  if form1.bsSkinMessage1.MessageDlg('载入数据将覆盖原有数据，确认载入？', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    exit;

  Form_loadproject.ListBox1.Clear;
  s := TStringList.Create;
  form1.ListDirs(ExtractFilePath(Application.ExeName) + '\projects\', s);
  Form_loadproject.ListBox1.Items.AddStrings(s);
  s.Free;

  Form_loadproject.ListBox1.Items.Delete(Form_loadproject.ListBox1.Items.IndexOf(label1.Caption)); //删除自身的项目


  Form_loadproject.Label2.Caption := '载入其它项目数据（全部工序）';
  Form_loadproject.ShowModal;

end;

procedure TForm1.bsSkinButton30Click(Sender: TObject);
var
  i: integer;
begin
  if DBGridEh5.SelectedRows.Count = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('请选择需要删除的记录！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if form1.bsSkinMessage1.MessageDlg('确定删除所选记录？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := DBGridEh5.SelectedRows.Count - 1 to 0 do
    begin
      DBGridEh5.DataSource.DataSet.GotoBookmark(pointer(DBGridEh5.SelectedRows[i]));
      DBGridEh5.DataSource.DataSet.Delete;
    end;
  end;
  ADOQuery2.Refresh;
  ListBox3Click(Sender);


end;

procedure TForm1.bsSkinButton71Click(Sender: TObject);
var
  WordApp, wordDoc, myTable: Variant;
  Aqr1, Aqr2: TAdoQuery;
  s, s1, s2, product: string;
  i, j, k, n: integer;
  s3, s4: widestring;
  SL, lcia: TStrings;
begin

  if label1.Caption = '项目名称' then exit;

  if TreeView1.Hint = '基础信息管理' then exit;


  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection2;

  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := form1.ADOConnection2;

  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('SELECT * FROM project_inf ');
    Aqr1.Open;
    Aqr1.Locate('project_item', '产品名称（最后工序）', [loCaseInsensitive]);
    product := Aqr1.FieldByName('item_inf').AsString;

    SaveDialog1.Filter := '*.doc';
    SaveDialog1.FileName := label1.Caption + 'LCA报告';
    SaveDialog1.Title := '保存';
    if SaveDialog1.Execute then
    begin
      try
        FlatProgressBar1.Position := 0;
        FlatProgressBar1.Max := 100;
        StatusBar1.Panels[3].Text := '正在生成LCA报告 ';
        StatusBar1.Refresh;

        wordApp := CreateOleObject('Word.Application'); //创建一个word对象

        wordDoc := WordApp.documents.open(ExtractFilePath(Application.ExeName) + '\LCAreport_Template.doc');
        wordapp.Documents.Item(1).Saveas(SaveDialog1.FileName);
        WordApp.Caption := product + 'LCA报告';


        WordApp.Documents.Item(1).Bookmarks.Item('projectname').Range.text := label1.Caption;
        Aqr1.Locate('project_item', 'LCA联系人信息', [loCaseInsensitive]);
        WordApp.Documents.Item(1).Bookmarks.Item('Contact').Range.text := Aqr1.FieldByName('item_inf').AsString;
        WordApp.Documents.Item(1).Bookmarks.Item('Reporttime').Range.text := datetimetostr(now); //报告生成时间

        Aqr1.Locate('project_item', '功能单位数据', [loCaseInsensitive]);
        s := Aqr1.FieldByName('item_inf').AsString + ' ';
        Aqr1.Locate('project_item', '功能单位', [loCaseInsensitive]);
        s := s + Aqr1.FieldByName('item_inf').AsString;
        WordApp.Documents.Item(1).Bookmarks.Item('FunctionUnit').Range.text := s + ' ' + product;
        WordApp.Documents.Item(1).Bookmarks.Item('functionUnit2').Range.text := s;
        WordApp.Documents.Item(1).Bookmarks.Item('functionUnit3').Range.text := s;

        FlatProgressBar1.Position := 10;

        WordApp.Documents.Item(1).Bookmarks.Item('product').Range.text := product;
        WordApp.Documents.Item(1).Bookmarks.Item('product1').Range.text := product;

        Aqr1.Locate('project_item', '产品系统描述', [loCaseInsensitive]);
        WordApp.Documents.Item(1).Bookmarks.Item('Systemdescribe').Range.text := Aqr1.FieldByName('item_inf').AsString;

        WordApp.Documents.Item(1).Bookmarks.Item('fig1').select;
        if FlatComboBox1.ItemIndex = 0 then //  从摇篮到大门  （Cradle to gate）
          WordApp.Selection.InlineShapes.AddPicture(ExtractFilePath(Application.ExeName) + '\system1.bmp', false, true, EmptyParam);
        if FlatComboBox1.ItemIndex = 1 then // 从摇篮到大门 ＋ 废钢循环  （Cradle to gate including recycling）
          WordApp.Selection.InlineShapes.AddPicture(ExtractFilePath(Application.ExeName) + '\system2.bmp', false, true, EmptyParam);
        if FlatComboBox1.ItemIndex = 2 then //  从摇篮到坟墓  （Cradle to grave）
          WordApp.Selection.InlineShapes.AddPicture(ExtractFilePath(Application.ExeName) + '\system3.bmp', false, true, EmptyParam);
        if FlatComboBox1.ItemIndex = 3 then //  从进厂门到出厂门（from gate to gate）
          WordApp.Selection.InlineShapes.AddPicture(ExtractFilePath(Application.ExeName) + '\system4.bmp', false, true, EmptyParam);
        if FlatComboBox1.ItemIndex = 4 then //  从摇篮到用户（from cradle to user）
          WordApp.Selection.InlineShapes.AddPicture(ExtractFilePath(Application.ExeName) + '\system5.bmp', false, true, EmptyParam);

        FlatProgressBar1.Position := 20;
        s := '';
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('SELECT ID, process  FROM data where processtype= ' + #39 + '主生产工序' + #39 + ' and id IN(select max(ID) from data group by process)' + ' order by ID ASC ');
        Aqr2.Open; // ,  ID, processtype

        for i := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := i;
          if i <> Aqr2.RecordCount then
            s := s + Aqr2.FieldbyName('process').asstring + '，'
          else
            s := s + Aqr2.FieldbyName('process').asstring + '。' + #13;
        end;
        s2 := Aqr2.FieldbyName('process').asstring;

        if Aqr2.RecordCount > 0 then
          s := '主生产工序：' + s;


        s1 := '';
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('SELECT ID, process  FROM data where processtype= ' + #39 + '能源公辅工序' + #39 + ' and id IN(select max(ID) from data group by process)' + ' order by ID ASC ');
        Aqr2.Open; // ,  ID, processtype

        for i := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := i;
          if i <> Aqr2.RecordCount then
            s1 := s1 + Aqr2.FieldbyName('process').asstring + '，'
          else
            s1 := s1 + Aqr2.FieldbyName('process').asstring + '。';
        end;
        if Aqr2.RecordCount > 0 then
          s1 := '能源公辅工序：' + s1;

        WordApp.Documents.Item(1).Bookmarks.Item('Process').Range.text := s + s1;

        WordApp.Documents.Item(1).Bookmarks.Item('fig2').select;
        WordApp.Selection.InlineShapes.AddPicture(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\' + form1.Label1.Caption + '.bmp', false, true, EmptyParam);

        WordApp.Documents.Item(1).Bookmarks.Item('product2').Range.text := product;

        WordApp.Documents.Item(1).Bookmarks.Item('LCAtype').Range.text := FlatComboBox1.Text;

    //  wRange:= wRange := wDoc.Range.goto(wdGoToBookmark, , , 'Table1');
        FlatProgressBar1.Position := 30;

        s := '';
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('SELECT *  FROM shuxing where shuxing=' + #39 + '电力构成' + #39);
        Aqr2.Open;
        for i := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := i;
          if Aqr2.FieldByName('shuzhi').AsFloat > 0 then
            s := s + Aqr2.FieldByName('name').AsString + '： ' + Aqr2.FieldByName('shuzhi').AsString + '%' + #13;
        end;

        WordApp.Documents.Item(1).Bookmarks.Item('Power').Range.text := s;


        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('SELECT *  FROM data where process=' + #39 + s2 + #39 + ' order by material_type asc');
        Aqr2.Open;


        WordApp.Documents.Item(1).Bookmarks.Item('Table1').select;
        myTable := WordApp.Selection.tables.add(WordApp.Selection.range, Aqr2.RecordCount + 1, 6, 2, 0);



        myTable.range.Font.Name := 'Times New Roman'; //针对整个表格
        myTable.range.Font.Size := 9;
        myTable.range.Font.Bold := false; //字体不加粗
        myTable.range.Font.Color := wdColorBlack; //字体颜色
        myTable.Rows.Alignment := wdAlignRowCenter; //表格居中

       //  myTable.Cell(1,1).Width :=30; //设置第一行第一列的宽度

        for i := 1 to 6 do
          myTable.cell(1, i).VerticalAlignment := wdCellAlignVerticalCenter;
        myTable.cell(1, 1).range.text := '序号';
        myTable.cell(1, 2).range.text := '工序名称';
        myTable.cell(1, 3).range.text := '数据类别';
        myTable.cell(1, 4).range.text := '数据名称';
        myTable.cell(1, 5).range.text := '单位';
        myTable.cell(1, 6).range.text := '数量';

        for i := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := i;
          myTable.cell(i + 1, 1).range.text := inttostr(i);
          myTable.cell(i + 1, 2).range.text := Aqr2.FieldByName('process').AsString;
          myTable.cell(i + 1, 3).range.text := Aqr2.FieldByName('material_type').AsString;
          myTable.cell(i + 1, 4).range.text := Aqr2.FieldByName('pd_name').AsString;
          myTable.cell(i + 1, 5).range.text := Aqr2.FieldByName('pd_unit').AsString;
          myTable.cell(i + 1, 6).range.text := FormatFloat('#,##0.000', Aqr2.FieldbyName('pd').AsFloat);
        end;


        myTable.AutoFitBehavior(1);
        FlatProgressBar1.Position := 40;

        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select ID, product,  sea_per,	train_per,	trunk_per,	river_per from transport ');
        Aqr2.Open;

        WordApp.Documents.Item(1).Bookmarks.Item('Table3').select;
        myTable := WordApp.Selection.tables.add(WordApp.Selection.range, Aqr2.RecordCount + 1, 6, 2, 0);


        myTable.range.Font.Name := 'Times New Roman'; //针对整个表格
        myTable.range.Font.Size := 9;
        myTable.range.Font.Bold := false; //字体不加粗
        myTable.range.Font.Color := wdColorBlack; //字体颜色
        myTable.Rows.Alignment := wdAlignRowCenter; //表格居中

       //  myTable.Cell(1,1).Width :=30; //设置第一行第一列的宽度

        for i := 1 to 6 do
          myTable.cell(1, i).VerticalAlignment := wdCellAlignVerticalCenter;
        myTable.cell(1, 1).range.text := '序号';
        myTable.cell(1, 2).range.text := '运输产品';
        myTable.cell(1, 3).range.text := '海运距离(km)';
        myTable.cell(1, 4).range.text := '火车距离(km)';
        myTable.cell(1, 5).range.text := '卡车距离(km)';
        myTable.cell(1, 6).range.text := '河运距离(km)';

        for i := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := i;
          myTable.cell(i + 1, 1).range.text := inttostr(i);
          myTable.cell(i + 1, 2).range.text := Aqr2.FieldByName('product').AsString;
          myTable.cell(i + 1, 3).range.text := FormatFloat('#,##0', Aqr2.FieldbyName('sea_per').AsFloat);
          myTable.cell(i + 1, 4).range.text := FormatFloat('#,##0', Aqr2.FieldbyName('train_per').AsFloat);
          myTable.cell(i + 1, 5).range.text := FormatFloat('#,##0', Aqr2.FieldbyName('trunk_per').AsFloat);
          myTable.cell(i + 1, 6).range.text := FormatFloat('#,##0', Aqr2.FieldbyName('river_per').AsFloat);
        end;


        myTable.AutoFitBehavior(1);
        FlatProgressBar1.Position := 50;

        Aqr1.Locate('project_item', '循环环境收益', [loCaseInsensitive]);
        WordApp.Documents.Item(1).Bookmarks.Item('Recycle').Range.text := Aqr1.FieldByName('item_inf').AsString;


        if FlatComboBox1.ItemIndex = 0 then //  从摇篮到大门  （Cradle to gate）
          WordApp.Documents.Item(1).Bookmarks.Item('data_type').Range.text := '（1）~（3）部分';
        if FlatComboBox1.ItemIndex = 1 then // 从摇篮到大门 ＋ 废钢循环  （Cradle to gate including recycling）
          WordApp.Documents.Item(1).Bookmarks.Item('data_type').Range.text := '（1）~（3），以及（5）';
        if FlatComboBox1.ItemIndex = 2 then //  从摇篮到坟墓  （Cradle to grave）
          WordApp.Documents.Item(1).Bookmarks.Item('data_type').Range.text := '全部数据类型';
        if FlatComboBox1.ItemIndex = 3 then //  从进厂门到出厂门（from gate to gate）
          WordApp.Documents.Item(1).Bookmarks.Item('data_type').Range.text := '企业内部生产制造过程的数据';
        if FlatComboBox1.ItemIndex = 4 then //  从摇篮到用户（from cradle to user）
          WordApp.Documents.Item(1).Bookmarks.Item('data_type').Range.text := '（1）~（4）部分';

        Aqr1.Locate('project_item', '数据取舍原则', [loCaseInsensitive]);
        WordApp.Documents.Item(1).Bookmarks.Item('Cutoff').Range.text := Aqr1.FieldByName('item_inf').AsString;

        Aqr1.Locate('project_item', '生产数据时间', [loCaseInsensitive]);
        WordApp.Documents.Item(1).Bookmarks.Item('Data_time').Range.text := Aqr1.FieldByName('item_inf').AsString;

        Aqr1.Locate('project_item', '数据分配', [loCaseInsensitive]);
        WordApp.Documents.Item(1).Bookmarks.Item('Data_allocation').Range.text := Aqr1.FieldByName('item_inf').AsString;

        Aqr1.Locate('project_item', '数据质量评价', [loCaseInsensitive]);
        WordApp.Documents.Item(1).Bookmarks.Item('Data_quality').Range.text := Aqr1.FieldByName('item_inf').AsString;

        WordApp.Documents.Item(1).Bookmarks.Item('product3').Range.text := product;
        WordApp.Documents.Item(1).Bookmarks.Item('product4').Range.text := product;
        FlatProgressBar1.Position := 60;

        s3 := trim(FlatEdit6.Text);
        s4 := stringreplace(s3, ';', #13, [rfReplaceAll]);
        SL := TstringList.Create;
        SL.text := s4;

        s := 'select ID, product,LCIfactor,unit,	g_site,g_outsite,g_route from LCA where LCItype <> ' + #39 + 'LCIA' + #39 + ' and (product=';
        for i := 0 to SL.count - 1 do
        begin
          if i <> SL.count - 1 then
            s := s + #39 + trim(SL.strings[i]) + #39 + ' or product='
          else
            s := s + #39 + trim(SL.strings[i]) + #39 + ')'
        end;

        s := s + ' and LCItype <> ' + #39 + 'LCC' + #39 + 'order by ID';
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add(s);
        Aqr2.Open;


        WordApp.Documents.Item(1).Bookmarks.Item('Table4').select;
        myTable := WordApp.Selection.tables.add(WordApp.Selection.range, Aqr2.RecordCount + 1, 7, 2, 0);


        myTable.range.Font.Name := 'Times New Roman'; //针对整个表格
        myTable.range.Font.Size := 9;
        myTable.range.Font.Bold := false; //字体不加粗
        myTable.range.Font.Color := wdColorBlack; //字体颜色
        myTable.Rows.Alignment := wdAlignRowCenter; //表格居中

        for i := 1 to 7 do
          myTable.cell(1, i).VerticalAlignment := wdCellAlignVerticalCenter;
        myTable.cell(1, 1).range.text := '序号';
        myTable.cell(1, 2).range.text := '产品名称';
        myTable.cell(1, 3).range.text := 'LCI指标';
        myTable.cell(1, 4).range.text := '单位';
        myTable.cell(1, 5).range.text := '包钢内部负荷';
        myTable.cell(1, 6).range.text := '包钢外部负荷';
        myTable.cell(1, 7).range.text := '合计';

        for i := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := i;
          myTable.cell(i + 1, 1).range.text := inttostr(i);
          myTable.cell(i + 1, 2).range.text := Aqr2.FieldByName('product').AsString;
          myTable.cell(i + 1, 3).range.text := Aqr2.FieldByName('LCIfactor').AsString;
          myTable.cell(i + 1, 4).range.text := Aqr2.FieldByName('unit').AsString;
          myTable.cell(i + 1, 5).range.text := FormatFloat('#,##0.##0', Aqr2.FieldbyName('g_site').AsFloat);
          myTable.cell(i + 1, 6).range.text := FormatFloat('#,##0.##0', Aqr2.FieldbyName('g_outsite').AsFloat);
          myTable.cell(i + 1, 7).range.text := FormatFloat('#,##0.##0', Aqr2.FieldbyName('g_route').AsFloat);
        end;
        myTable.AutoFitBehavior(1);
        FlatProgressBar1.Position := 70;


        s := 'select ID, product,LCIfactor,unit,	g_site,g_outsite,g_route from LCA where LCItype = ' + #39 + 'LCIA' + #39 + ' and (product=';
        for i := 0 to SL.count - 1 do
        begin
          if i <> SL.count - 1 then
            s := s + #39 + trim(SL.strings[i]) + #39 + ' or product='
          else
            s := s + #39 + trim(SL.strings[i]) + #39 + ')'
        end;

        s := s + ' and LCItype <> ' + #39 + 'LCC' + #39 + 'order by ID';
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add(s);
        Aqr2.Open;


        WordApp.Documents.Item(1).Bookmarks.Item('Table5').select;
        myTable := WordApp.Selection.tables.add(WordApp.Selection.range, Aqr2.RecordCount + 1, 7, 2, 0);


        myTable.range.Font.Name := 'Times New Roman'; //针对整个表格
        myTable.range.Font.Size := 9;
        myTable.range.Font.Bold := false; //字体不加粗
        myTable.range.Font.Color := wdColorBlack; //字体颜色
        myTable.Rows.Alignment := wdAlignRowCenter; //表格居中

        for i := 1 to 7 do
          myTable.cell(1, i).VerticalAlignment := wdCellAlignVerticalCenter;
        myTable.cell(1, 1).range.text := '序号';
        myTable.cell(1, 2).range.text := '产品名称';
        myTable.cell(1, 3).range.text := 'LCIA指标';
        myTable.cell(1, 4).range.text := '单位';
        myTable.cell(1, 5).range.text := '内部负荷';
        myTable.cell(1, 6).range.text := '外部负荷';
        myTable.cell(1, 7).range.text := '合计';

        for i := 1 to Aqr2.RecordCount do
        begin
          Aqr2.RecNo := i;
          myTable.cell(i + 1, 1).range.text := inttostr(i);
          myTable.cell(i + 1, 2).range.text := Aqr2.FieldByName('product').AsString;
          myTable.cell(i + 1, 3).range.text := Aqr2.FieldByName('LCIfactor').AsString;
          myTable.cell(i + 1, 4).range.text := Aqr2.FieldByName('unit').AsString;
          myTable.cell(i + 1, 5).range.text := FormatFloat('#,##0.##0', Aqr2.FieldbyName('g_site').AsFloat);
          myTable.cell(i + 1, 6).range.text := FormatFloat('#,##0.##0', Aqr2.FieldbyName('g_outsite').AsFloat);
          myTable.cell(i + 1, 7).range.text := FormatFloat('#,##0.##0', Aqr2.FieldbyName('g_route').AsFloat);
        end;
        myTable.AutoFitBehavior(1);
        FlatProgressBar1.Position := 80;

        lcia := TStringList.Create;
        lcia.Add('资源消耗');
        lcia.Add('能源消耗');
        lcia.Add('气候变化');
        lcia.Add('酸化');
        lcia.Add('富营养化');

        PageControl1.TabIndex := 9;
        PageControl1Change(Sender);

        for j := 0 to bsSkinCheckListBox8.Items.Count - 1 do //选中项目名称
        begin
          bsSkinCheckListBox8.Checked[j] := false;
          if bsSkinCheckListBox8.Items[j] = label1.Caption then
            bsSkinCheckListBox8.Checked[j] := true;
        end;
        bsSkinCheckListBox8ClickCheck(Sender);
        CheckBox6.Checked := true;

        for i := 0 to SL.count - 1 do //
        begin
          for j := 0 to bsSkinCheckListBox5.Items.Count - 1 do //选中产品名称
          begin
            bsSkinCheckListBox5.Checked[j] := false;
            if bsSkinCheckListBox5.Items[j] = trim(SL.strings[i]) then
              bsSkinCheckListBox5.Checked[j] := true;
          end;
          ComboBox7.ItemIndex := 6;
          ComboBox7Change(Sender);
          for k := 0 to 4 do
          begin
            for j := 0 to bsSkinCheckListBox55.Items.Count - 1 do //选中LCIA类别
            begin
              bsSkinCheckListBox55.Checked[j] := false;
              if bsSkinCheckListBox55.Items[j] = lcia.strings[k] then
                bsSkinCheckListBox55.Checked[j] := true;
            end;
            bsSkinButton33Click(Sender);
            chart1.CopyToClipboardBitmap;
            wordDoc.Range(wordDoc.Range.End - 1, wordDoc.Range.End - 1).paste;
            wordDoc.Range(wordDoc.Range.End - 1, wordDoc.Range.End - 1).Text := '图' + inttostr(4 + 5 * i + k) + ' ' + trim(SL.strings[i]) + lcia.Strings[k] + '生命周期阶段分布' + #13;
          end;
          n := 4 + 5 * i + k;
        end; // for SL


        wordDoc.Range(wordDoc.Range.End - 1, wordDoc.Range.End - 1).Text := '2、环境负荷在各工序的分布如图所示：' + #13;

        for i := 0 to SL.count - 1 do
        begin
          for j := 0 to bsSkinCheckListBox5.Items.Count - 1 do //选中产品名称
          begin
            bsSkinCheckListBox5.Checked[j] := false;
            if bsSkinCheckListBox5.Items[j] = trim(SL.strings[i]) then
              bsSkinCheckListBox5.Checked[j] := true;
          end;

          ComboBox7.ItemIndex := 6;
          ComboBox7Change(Sender);


          for k := 0 to 4 do
          begin
            for j := 0 to bsSkinCheckListBox55.Items.Count - 1 do //选中LCIA类别
            begin
              bsSkinCheckListBox55.Checked[j] := false;
              if bsSkinCheckListBox55.Items[j] = lcia.strings[k] then
                bsSkinCheckListBox55.Checked[j] := true;
            end;
            bsSkinButton67Click(Sender);
            chart1.CopyToClipboardBitmap;
            wordDoc.Range(wordDoc.Range.End - 1, wordDoc.Range.End - 1).paste;
            wordDoc.Range(wordDoc.Range.End - 1, wordDoc.Range.End - 1).Text := '图' + inttostr(n + 5 * i + k) + ' ' + trim(SL.strings[i]) + lcia.Strings[k] + '在各工序的分布' + #13;


          end;

        end; // for SL

        wordapp.Documents.Item(1).Save;
        WordApp.Documents.Item(1).Bookmarks.Item('top').select;

        wordApp.Visible := true;
        WordApp.Activate;
        FlatProgressBar1.Position := 100;

        StatusBar1.Panels[3].Text := '  ';
        StatusBar1.Refresh;
        FlatProgressBar1.Position := 0;
      except
        wordapp.Documents.close;
        form1.bsSkinMessage1.MessageDlg('运行 Microsoft Word 失败！', (mtInformation), [mbOK], 0);

      end;

    end;
  finally
    Aqr1.Close;
    Aqr1.Free;
    SL.Free;
    lcia.Free;
  end;

end;




procedure TForm1.ComboBox8Change(Sender: TObject);

begin
  if ComboBox8.ItemIndex = -1 then
    exit;

  DBGridEh6.DataSource.DataSet.Filtered := false;
  DBGridEh6.DataSource.DataSet.Filter := 'process=' + #39 + ComboBox8.Items[ComboBox8.ItemIndex] + #39;
  if ComboBox8.Itemindex > 0 then
    DBGridEh6.DataSource.DataSet.Filtered := true
  else
    DBGridEh6.DataSource.DataSet.Filtered := false;

end;

procedure TForm1.LCA5Click(Sender: TObject);
begin
  bsSkinButton71Click(Sender);
end;

procedure TForm1.N28Click(Sender: TObject);
begin
  bsSkinButton4Click(Sender); //上游数据管理
end;

procedure TForm1.N29Click(Sender: TObject);
begin
  bsSkinButton5Click(Sender);
end;

procedure TForm1.N4Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 8;
  PageControl1Change(Sender);
end;

procedure TForm1.N5Click(Sender: TObject);
begin
  bsSkinButton46Click(Sender);
end;

procedure TForm1.LCA6Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 0;
  PageControl1Change(Sender);
end;

procedure TForm1.N6Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 1;
  PageControl1Change(Sender);
end;

procedure TForm1.N7Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 3;
  PageControl1Change(Sender);
end;

procedure TForm1.N17Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 6;
  PageControl1Change(Sender);
end;

procedure TForm1.N2Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 2;
  PageControl1Change(Sender);
end;

procedure TForm1.N3Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 4;
  PageControl1Change(Sender);
end;

procedure TForm1.N16Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then
    exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 5;
  PageControl1Change(Sender);
end;

procedure TForm1.LCI1Click(Sender: TObject);
begin
  bsSkinButton11Click(Sender);
end;

procedure TForm1.LCIA1Click(Sender: TObject);
begin
  bsSkinButton12Click(Sender);
end;

procedure TForm1.bsSkinButton72Click(Sender: TObject);
var
  i, j: integer;
  Aqr1: TAdoQuery;
begin
  if ListBox6.ItemIndex = -1 then exit;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;

  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from scrapeLCA where inf_name=' + #39 + 'S' + #39 + ' and inf=' + #39 + ListBox6.Items[ListBox6.ItemIndex] + #39);
    Aqr1.Open;
    if Aqr1.RecordCount > 0 then
      Aqr1.Delete;


    ComboBox3.Items.Delete(ListBox6.ItemIndex);
    ListBox6.Items.Delete(ListBox6.ItemIndex);
    ComboBox3.Text := '';
  finally
    Aqr1.Close;
    Aqr1.Free;

  end;


end;

procedure TForm1.bsSkinButton73Click(Sender: TObject);
var
  i, j, k: integer;
  Aqr1, Aqr2, Aqr4: TAdoQuery;
  RR, y, s, heji: double;
  mysql, product: string;
begin

  if ListBox6.Items.Count = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('没有添加转炉工序', (mtInformation), [mbOK], 0);
    exit;
  end;

  if trim(bsSkinEdit2.text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请输入废钢循环率RR', (mtInformation), [mbOK], 0);
    exit;
  end;

  if trim(bsSkinEdit3.text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请输入废钢利用效率Y', (mtInformation), [mbOK], 0);
    exit;
  end;

  if RadioButton2.Checked = true then
    if ComboBox9.ItemIndex = -1 then
    begin
      form1.bsSkinMessage1.MessageDlg('请选择全废钢炼钢电炉工序名称', (mtInformation), [mbOK], 0);
      exit;
    end;


  FlatProgressBar1.Position := 0;
  FlatProgressBar1.Max := 100;
  StatusBar1.Panels[3].Text := '正在计算 ';
  StatusBar1.Refresh;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection2;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from scrapeLCA ');
    Aqr1.Open;

    if Aqr1.Locate('inf_name', 'RR', []) = false then
      Aqr1.Append
    else
      Aqr1.Edit;

    Aqr1.FieldbyName('inf_name').Asstring := 'RR';
    Aqr1.FieldbyName('inf').Asstring := '废钢回收率';
    Aqr1.FieldbyName('unit').Asstring := '%';
    Aqr1.FieldbyName('num').Asstring := bsSkinEdit2.text;
    Aqr1.post;

    if Aqr1.Locate('inf_name', 'Y', []) = false then
      Aqr1.Append
    else
      Aqr1.Edit;

    Aqr1.FieldbyName('inf_name').Asstring := 'Y';
    Aqr1.FieldbyName('inf').Asstring := '废钢利用效率';
    Aqr1.FieldbyName('unit').Asstring := 'kg';
    Aqr1.FieldbyName('num').Asstring := bsSkinEdit3.text;
    Aqr1.post;
    FlatProgressBar1.Position := 10;
    for i := 0 to ListBox6.Items.Count - 1 do
    begin
      if Aqr1.Locate('inf', ListBox6.Items[i], []) = false then
        Aqr1.Append
      else
        Aqr1.Edit;
      Aqr1.FieldbyName('inf_name').Asstring := 'S';
      Aqr1.FieldbyName('inf').Asstring := ListBox6.Items[i];
      Aqr1.FieldbyName('unit').Asstring := 'kg';
      Aqr1.FieldbyName('num').Asstring := ComboBox3.Items[i];
      Aqr1.post;
    end;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select *  from data_allocation where process= ' + #39 + ListBox6.Items[0] + #39 + ' and material_type=' + #39 + '产品' + #39);
    Aqr1.Open;
    product := Aqr1.FieldByName('pd_name').AsString;

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select * from LCA where product= ' + #39 + product + #39 + ' order by id ASC');
    Aqr2.Open;

    Scrape.NrOfColumns := ListBox6.Items.Count + 3; // X1  X2  +   Xre  +   X平均  +  Xpr
    Scrape.NrOfRows := Aqr2.RecordCount + 1; //留第一行做产量加权值


    FlatProgressBar1.Position := 20;



    for i := 1 to Scrape.NrOfRows do
      for j := 1 to Scrape.NrOfColumns do
        Scrape.Elem[j, i] := 0;

    mysql := 'select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and (process=';
    for i := 0 to ListBox6.Items.Count - 1 do
      if i <> ListBox6.Items.Count - 1 then
        mysql := mysql + #39 + ListBox6.Items[i] + #39 + ' or process='
      else
        mysql := mysql + #39 + ListBox6.Items[i] + #39 + ' )';

    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add(mysql);
    Aqr2.Open;


    heji := 0;
    for i := 0 to ListBox6.Items.Count - 1 do //   X1  X2  X3 ......
    begin
      if Aqr2.Locate('process', ListBox6.Items[i], []) = true then
      begin
        Scrape.Elem[i + 1, 1] := Aqr2.FieldByName('pd').AsFloat * unit_exchange(Aqr2.FieldByName('pd_unit').AsString); //产量
        heji := heji + Aqr2.FieldByName('pd').AsFloat * unit_exchange(Aqr2.FieldByName('pd_unit').AsString); // 产量求和

        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add('select * from LCA where product= ' + #39 + Aqr2.FieldByName('pd_name').AsString + #39 + ' order by id ASC');
        Aqr1.Open;
        for j := 1 to Aqr1.RecordCount do
        begin
          Aqr1.RecNo := j;
          Scrape.Elem[i + 1, j + 1] := Aqr1.FieldByName('g_route').AsFloat;
        end;
      end;

    end;
    FlatProgressBar1.Position := 30;
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from scrapeLCA where inf_name=' + #39 + 'Xre_source' + #39);
    Aqr1.Open;
    if Aqr1.RecordCount = 0 then
      Aqr1.Append
    else
      Aqr1.Edit;
    if RadioButton1.Checked = true then
      Aqr1.FieldByName('num').AsFloat := -1;
    if RadioButton2.Checked = true then
      Aqr1.FieldByName('num').AsFloat := ComboBox9.ItemIndex;
    Aqr1.FieldByName('inf_name').AsString := 'Xre_source';
    Aqr1.FieldByName('inf').AsString := '全废钢电炉炼钢数据来源';
    Aqr1.Post;




    if RadioButton2.Checked = true then
    begin
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select *  from data_allocation where process= ' + #39 + ComboBox9.Text + #39 + ' and material_type=' + #39 + '产品' + #39);
      Aqr1.Open;
      product := Aqr1.FieldByName('pd_name').AsString;
    end;


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    if RadioButton1.Checked = true then
      Aqr1.SQL.Add('select ID,inf_name,LCItype,LCIfactor,unit,num  from scrapeLCA where inf_name= ' + #39 + 'Xre' + #39 + ' order by id ASC');
    if RadioButton2.Checked = true then
      Aqr1.SQL.Add('select ID,product,LCItype,LCIfactor,unit,g_route  from LCA where product= ' + #39 + product + #39 + ' order by id ASC');
    Aqr1.Open;
    for j := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := j;
      Scrape.Elem[ListBox6.Items.Count + 1, j + 1] := Aqr1.Fields[5].AsFloat;
    end;

    FlatProgressBar1.Position := 40;
             // X1  X2  +   Xre  +   X平均  +  Xpr            X平均 =X1*A/SUM+X2*B/SUM+.....

    for j := 1 to Scrape.NrOfRows do
      Scrape.Elem[ListBox6.Items.Count + 2, j + 1] := 0; //赋初值

    for j := 1 to Scrape.NrOfRows do //   计算X 如果有多个转炉工序，按产量进行加权平均
    begin
      for i := 1 to ListBox6.Items.Count do
        Scrape.Elem[ListBox6.Items.Count + 2, j + 1] := Scrape.Elem[ListBox6.Items.Count + 2, j + 1] + Scrape.Elem[i, j + 1] * Scrape.Elem[i, 1] / heji;
    end;

    s := 0;
    RR := strtofloat(bsSkinEdit2.Text) / 100;
    y := strtofloat(bsSkinEdit3.Text);
    for i := 1 to ListBox6.Items.Count do
      s := s + strtofloat(ComboBox3.Items[i - 1]) * Scrape.Elem[i, 1] / heji;

    for j := 1 to Scrape.NrOfRows do //   计算Xpr
    begin
      Scrape.Elem[ListBox6.Items.Count + 3, j + 1] := (Scrape.Elem[ListBox6.Items.Count + 2, j + 1] - s * y * Scrape.Elem[ListBox6.Items.Count + 1, j + 1]) / (1 - s * y);
    end; //  Xpr=(X-S*Y*Xre )/(1-S*Y)


    FlatProgressBar1.Position := 50;

// checkmat(Scrape);


  //******************************************************判断转炉之后工序  : 根据完全消耗系数   >0
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr1.Open;

    PA1.NrOfColumns := Aqr1.RecordCount;
    PA1.NrOfRows := Aqr1.RecordCount;
    for i := 1 to Aqr1.RecordCount do
      for j := 1 to Aqr1.RecordCount do
        PA1.Elem[i, j] := 0; //  elem[列，行]        PA1 行 工序  列 产品

    if Aqr4.Active then Aqr4.Close;
    Aqr4.SQL.Clear;
    Aqr4.SQL.Add('select * from data_allocation where material_type = ' + #39 + '产品' + #39 + ' and processtype=' + #39 + '主生产工序' + #39);
    Aqr4.Open;


    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      for j := 1 to Aqr4.RecordCount do
      begin
        Aqr4.RecNo := j;
        if Aqr2.Active then Aqr2.Close;
        Aqr2.SQL.Clear;
        Aqr2.SQL.Add('select * from data_allocation where  process = ' + #39 + Aqr1.FieldByName('process').AsString + #39 + ' and pd_name=' + #39 + Aqr4.FieldByName('pd_name').AsString + #39 + ' and material_type <>' + #39 + '副产品或固体废弃物' + #39);
        Aqr2.SQL.Add(' and material_type <>' + #39 + '产品' + #39);
        Aqr2.Open;
        if Aqr2.RecordCount > 0 then
        begin
          Aqr2.First;
          PA1.Elem[i, j] := Aqr2.FieldByName('co_num').AsFloat;
        end;
      end;
    end;

    FlatProgressBar1.Position := 60;

    I_PA1.NrOfColumns := PA1.NrOfColumns;
    I_PA1.NrOfRows := PA1.NrOfRows;
    for i := 1 to I_PA1.NrOfRows do // I-PA1
      for j := 1 to I_PA1.NrOfColumns do
        I_PA1.Elem[j, i] := 0;

    for i := 1 to PA1.NrOfColumns do // I-PA1
      for j := 1 to PA1.NrOfRows do
        if i = j then
          I_PA1.Elem[i, j] := 1 - PA1.Elem[i, j]
        else
          I_PA1.Elem[i, j] := -PA1.Elem[i, j];


    I_PA1.Invert; //     求逆       求逆    求逆    求逆    求逆    求逆    求逆    求逆
    FlatProgressBar1.Position := 70;
    ListBox7.Clear;
    for j := 0 to ListBox6.Items.Count - 1 do
    begin
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        if Aqr1.FieldByName('process').AsString = ListBox6.Items[j] then
          for k := 1 to Aqr1.RecordCount do
            if I_PA1.Elem[k, i] > 0 then
            begin
              Aqr1.RecNo := k;
              if ListBox7.Items.IndexOf(Aqr1.FieldByName('pd_name').AsString) = -1 then
                ListBox7.Items.Add(Aqr1.FieldByName('pd_name').AsString);
            end;
      end;
    end;


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('delete *  from  scrapeLCA  where inf_name= ' + #39 + 'LCIincludingEOL' + #39);
    Aqr1.ExecSQL;

    FlatProgressBar1.Position := 80;

                               // X1  X2  +   Xre  +   X平均  +  Xpr            X平均 =X1*A/SUM+X2*B/SUM+.....                Scrape
    for i := 0 to ListBox7.Items.Count - 1 do
    begin

      product := ListBox7.Items[i];

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select *  from LCA where product= ' + #39 + product + #39 + ' order by ID ASC');
      Aqr1.Open;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select *  from scrapeLCA ');
      Aqr2.Open;

      for j := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := j;
        Aqr2.Append;
        Aqr2.FieldByName('inf_name').AsString := 'LCIincludingEOL';
        Aqr2.FieldByName('inf').AsString := Aqr1.FieldByName('product').AsString;
        Aqr2.FieldByName('num').AsFloat := Aqr1.FieldByName('g_route').AsFloat - (rr - s) * (Scrape.Elem[ListBox6.Items.Count + 3, j + 1] - Scrape.Elem[ListBox6.Items.Count + 1, j + 1]) * y;
          //LCIincludingEOL=  X-(rr-s)(Xpr-Xre)Y
        Aqr2.FieldByName('LCIfactor').AsString := Aqr1.FieldByName('LCIfactor').AsString;
        Aqr2.FieldByName('LCItype').AsString := Aqr1.FieldByName('LCItype').AsString;
        Aqr2.FieldByName('unit').AsString := Aqr1.FieldByName('unit').AsString;
        Aqr2.Post;
      end;

    end;
    FlatProgressBar1.Position := 90;
    ListBox7.ItemIndex := 0;
    ListBox7.Selected[0] := true;
    ListBox7Click(Sender);

    FlatProgressBar1.Position := 100;
    StatusBar1.Panels[3].Text := '';
    StatusBar1.Refresh;


//*********************************************************

  finally
    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr4.Close;
    Aqr4.Free;
    FlatProgressBar1.Position := 0;
  end;

end;

procedure TForm1.bsSkinButton14Click(Sender: TObject);
var
  i, j: integer;
  Aqr1, Aqr2: TAdoQuery;
begin
  if ComboBox1.ItemIndex = -1 then exit;
  if ComboBox1.Text = '' then exit;
  if ListBox6.Items.IndexOf(ComboBox1.Text) <> -1 then
  begin
    bsSkinMessage1.MessageDlg(ComboBox1.Text + '已存在！', (mtInformation), [mbOK], 0);
    exit;
  end;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;

  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation where process= ' + #39 + ComboBox1.Text + #39 + ' and material_type=' + #39 + '原材料' + #39 + ' and (pd_name=' + #39 + '废钢' + #39 + ' or pd_name=' + #39 + '处理后废钢' + #39 + ')');
    Aqr1.Open;
    if Aqr1.RecordCount = 0 then
    begin
      bsSkinMessage1.MessageDlg(ComboBox1.Text + '工序原材料中未查询到 废钢 数据，请检查数据', (mtInformation), [mbOK], 0);
      exit;
    end;
    ListBox6.Items.Add(ComboBox1.Text);
    ComboBox3.Items.Add(formatfloat('##0.###0', Aqr1.FieldByName('co_num').AsFloat));
    ComboBox3.ItemIndex := ComboBox3.Items.IndexOf(formatfloat('##0.###0', Aqr1.FieldByName('co_num').AsFloat));
    ListBox6.Selected[ComboBox3.ItemIndex] := true;
  finally
    Aqr1.Close;
    Aqr1.Free;

  end;

end;

procedure TForm1.ListBox6Click(Sender: TObject);
begin
  if ListBox6.ItemIndex <> -1 then
    ComboBox3.ItemIndex := ListBox6.ItemIndex;


end;

procedure TForm1.ComboBox3Change(Sender: TObject);
begin

  if ComboBox3.ItemIndex <> -1 then
  begin
    listbox6.ClearSelection;
    ListBox6.ItemIndex := ComboBox3.ItemIndex;
    ListBox6.Selected[ListBox6.ItemIndex] := true;
  end;




end;

procedure TForm1.ADOQuery9AfterOpen(DataSet: TDataSet);
var
  i: integer;
begin
  ADOQuery9.FieldByName('ID').OnGetText := SetFdText9;
  for i := 1 to ADOQuery9.FieldCount - 1 do
    if ADOQuery9.Fields[i].DataType = ftFloat then
      (ADOQuery9.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0';

  ADOQuery9.First;

end;

procedure TForm1.bsSkinEdit2KeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9', '.', #13, #9, #8]) then
    key := #0;
end;

procedure TForm1.bsSkinEdit3KeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9', '.', #13, #9, #8]) then
    key := #0;
end;

procedure TForm1.Label39Click(Sender: TObject);
var
  i: integer;
  product: string;
begin
  if RadioButton2.Checked = true then
    if ComboBox9.ItemIndex = -1 then
    begin
      form1.bsSkinMessage1.MessageDlg('请选择全废钢炼钢电炉工序名称', (mtInformation), [mbOK], 0);
      exit;
    end;

  ADOQuery9.Connection := ADOConnection2;
  if RadioButton2.Checked = true then
  begin
    if ADOQuery9.Active then ADOQuery9.Close;
    ADOQuery9.SQL.Clear;
    ADOQuery9.SQL.Add('select *  from data_allocation where process= ' + #39 + ComboBox9.Text + #39 + ' and material_type=' + #39 + '产品' + #39);
    ADOQuery9.Open;
    product := ADOQuery9.FieldByName('pd_name').AsString;
  end;

  if ADOQuery9.Active then ADOQuery9.Close;
  ADOQuery9.SQL.Clear;
  if RadioButton1.Checked = true then
    ADOQuery9.SQL.Add('select ID,inf_name,LCItype,LCIfactor,unit,num  from scrapeLCA where inf_name= ' + #39 + 'Xre' + #39);
  if RadioButton2.Checked = true then
    ADOQuery9.SQL.Add('select ID,product,LCItype,LCIfactor,unit,g_route  from LCA where product= ' + #39 + product + #39);

  ADOQuery9.Open;

  for i := 0 to DBGridEh11.Columns.Count - 1 do
  begin
    DBGridEh11.Columns[i].Width := 80;
    DBGridEh11.Columns[i].Title.Alignment := taCenter;
    DBGridEh11.Columns[i].Alignment := taCenter;
  end;

  DBGridEh11.Columns[5].Alignment := taRightJustify; //数据
  DBGridEh11.Columns[2].Alignment := taLeftJustify;
  DBGridEh11.Columns[3].Alignment := taLeftJustify;

  DBGridEh11.Columns[0].Title.caption := '序号';
  DBGridEh11.Columns[1].Title.caption := '产品';
  DBGridEh11.Columns[2].Title.caption := '类别';
  DBGridEh11.Columns[3].Title.caption := 'LCI因子';
  DBGridEh11.Columns[4].Title.caption := '单位';
  DBGridEh11.Columns[5].Title.caption := '数量';

  DBGridEh11.Columns[0].Width := 40;
  DBGridEh11.Columns[1].Width := 40;
  DBGridEh11.Columns[2].Width := 80;
  DBGridEh11.Columns[3].Width := 120;
  DBGridEh11.Columns[4].Width := 50;

end;

procedure TForm1.RadioButton2Click(Sender: TObject);
begin
  if RadioButton2.Checked = true then
    ComboBox9.Visible := true;
end;

procedure TForm1.RadioButton1Click(Sender: TObject);
begin
  if RadioButton1.Checked = true then
    ComboBox9.Visible := false;
end;

procedure TForm1.ListBox7Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;

begin
  if ListBox7.Items.Count = 0 then exit;
  if ListBox7.ItemIndex = -1 then exit;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;

  ADOQuery9.Connection := ADOConnection2;
  try


    if ADOQuery9.Active then ADOQuery9.Close;
    ADOQuery9.SQL.Clear;
    ADOQuery9.SQL.Add('select ID,inf,LCItype,LCIfactor,unit,num from scrapeLCA where inf_name=' + #39 + 'LCIincludingEOL' + #39 + ' and inf=' + #39 + ListBox7.Items[ListBox7.ItemIndex] + #39 + ' order by id');
    ADOQuery9.Open;

    for i := 0 to DBGridEh11.Columns.Count - 1 do
    begin
      DBGridEh11.Columns[i].Width := 80;
      DBGridEh11.Columns[i].Title.Alignment := taCenter;
      DBGridEh11.Columns[i].Alignment := taCenter;
    end;

    DBGridEh11.Columns[5].Alignment := taRightJustify; //数据
    DBGridEh11.Columns[2].Alignment := taLeftJustify;
    DBGridEh11.Columns[3].Alignment := taLeftJustify;

    DBGridEh11.Columns[0].Title.caption := '序号';
    DBGridEh11.Columns[1].Title.caption := '产品';
    DBGridEh11.Columns[2].Title.caption := '类别';
    DBGridEh11.Columns[3].Title.caption := 'LCI因子';
    DBGridEh11.Columns[4].Title.caption := '单位';
    DBGridEh11.Columns[5].Title.caption := '数量';

    DBGridEh11.Columns[0].Width := 40;
    DBGridEh11.Columns[1].Width := 80;
    DBGridEh11.Columns[2].Width := 80;
    DBGridEh11.Columns[3].Width := 100;
    DBGridEh11.Columns[4].Width := 50;


  finally
    Aqr1.Close;
    Aqr1.Free;


  end;





end;

procedure TForm1.bsSkinButton78Click(Sender: TObject);
begin
  if ListBox7.ItemIndex = -1 then exit;
  ADOQuery9.Connection := ADOConnection2;
  if ADOQuery9.Active then ADOQuery9.Close;
  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('select LCA.ID,LCA.LCIfactor,LCA.g_route,  scrapeLCA.num,  LCA.g_route-scrapeLCA.num from LCA, scrapeLCA where LCA.product=' + #39 + ListBox7.Items[ListBox7.ItemIndex] + #39 + ' and scrapeLCA.inf=' + #39 + ListBox7.Items[ListBox7.ItemIndex] + #39 + ' and LCA.LCIfactor=scrapeLCA.LCIfactor');
  ADOQuery9.Open;
  DBGridEh11.Columns[0].Title.caption := '序号';
  DBGridEh11.Columns[1].Title.caption := 'LCI因子';
  DBGridEh11.Columns[2].Title.caption := '废钢循环前';
  DBGridEh11.Columns[3].Title.caption := '废钢循环后';
  DBGridEh11.Columns[4].Title.caption := '前后差值';
  DBGridEh11.Columns[0].Width := 40;
  DBGridEh11.Columns[1].Width := 120;
  DBGridEh11.Columns[2].Width := 90;
  DBGridEh11.Columns[3].Width := 90;
  DBGridEh11.Columns[4].Width := 90;

//  ADOQuery9.SQL.Add('select LCA.ID,LCA.LCIfactor,LCA.g_route,  scrapeLCA.num, case when LCA.g_route<>0 then 1-scrapeLCA.num/LCA.g_route else 0 end as bili  from LCA, scrapeLCA where LCA.product=' + #39 + ListBox7.Items[ListBox7.ItemIndex] + #39 + ' and scrapeLCA.inf=' + #39 + ListBox7.Items[ListBox7.ItemIndex] + #39 + ' and LCA.LCIfactor=scrapeLCA.LCIfactor');


end;

procedure TForm1.bsSkinCheckListBox5ClickCheck(Sender: TObject);
var
  i, j: integer;
begin
  if PageControl1.TabIndex = 10 then //只能单选
  begin
    j := bsSkinCheckListBox5.ItemIndex;
    if bsSkinCheckListBox5.Selected[j] then
      for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
        bsSkinCheckListBox5.Checked[i] := i = j;

  end;


  if PageControl1.TabIndex = 10 then
    CheckBox7.Visible := false
  else
    CheckBox7.Visible := true;

  if PageControl1.TabIndex = 10 then
    Button3Click(Sender);


end;

procedure TForm1.Button3Click(Sender: TObject);
var
  i, j: integer;
  mysql, s, s1, product1, process: string;
  Aqr1, Aqr2, Aqr3, Aqr4: TAdoQuery;
  e, a: double;
begin
  if checkedcount(bsSkinCheckListBox8) = 0 then exit;

  if ComboBox7.ItemIndex > 1 then
    bsSkinButton76.Visible := false
  else
    bsSkinButton76.Visible := true;

  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;

  ADOConnection5.Close;
  ADOConnection5.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
  ADOConnection5.Open;
  ADOQuery10.Connection := form1.ADOConnection5;

  if ADOQuery10.Active then ADOQuery10.Close;
  ADOQuery10.SQL.Clear;
  ADOQuery10.SQL.Add('delete * from temp');
  ADOQuery10.ExecSQL;

  if ADOQuery10.Active then ADOQuery10.Close;
  ADOQuery10.SQL.Clear;
  ADOQuery10.SQL.Add('select * from temp');
  ADOQuery10.Open;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection1; //连接对照表
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection5;
  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection1; //连接historydata
  Aqr4 := TAdoQuery.Create(nil);
  Aqr4.Connection := ADOConnection5;

  if Aqr4.Active then Aqr4.Close;
  Aqr4.SQL.Clear;
  Aqr4.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
  Aqr4.Open;
  process := Aqr4.FieldByName('process').AsString;
  s1 := Aqr4.FieldByName('co_unit').AsString;
  s1 := copy(s1, pos('/', s1) + 1, length(s1) - pos('/', s1));

  try
    if ComboBox7.ItemIndex = 0 then //工序能耗
    begin
      Aqr1.Connection := ADOConnection5;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from shuxing where shuxing = ' + #39 + '折标煤系数' + #39);
      Aqr1.Open;

      Aqr4.Connection := ADOConnection1; //连接对照表
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, shuxing from comparison where shuxing is not null ');
      Aqr4.Open;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
      Aqr2.Open;

      e := 0; a := 0;
      for j := 1 to Aqr2.RecordCount do
      begin
        Aqr2.RecNo := j;
        if Aqr4.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
        begin
          if Aqr1.Locate('name', Aqr4.FieldByName('shuxing').AsString, []) = true then
          begin
            if Aqr2.FieldByName('material_type').AsString = '产品' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '副产品或固体废弃物' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '能源及能源介质' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '原材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '辅助材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '大气排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '水体排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;


            ADOQuery10.Append;
            ADOQuery10.FieldByName('lcifactor').AsString := '工序能耗';
            ADOQuery10.FieldByName('process').AsString := process;
            ADOQuery10.FieldByName('pd_name').AsString := Aqr2.FieldByName('pd_name').AsString;
            if s1 = 'kg' then
              a := 1000 * a; //化成 t 为单位
            ADOQuery10.FieldByName('AF').AsFloat := a;
            ADOQuery10.FieldByName('EF').AsFloat := Aqr1.FieldByName('shuzhi').AsFloat;
            ADOQuery10.FieldByName('AFxEF').AsFloat := a * Aqr1.FieldByName('shuzhi').AsFloat;
            ADOQuery10.Post;
            e := e + a * Aqr1.FieldByName('shuzhi').AsFloat;
          end;
        end;
      end;

   //   bsSkinEdit4.Text := FormatFloat('#,##0.000', e);

      if ADOQuery10.Active then ADOQuery10.Close;
      ADOQuery10.SQL.Clear;
      ADOQuery10.SQL.Add('select ID,pd_name,AF,EF,AFxEF from temp order by AFxEF ASC');
      ADOQuery10.Open;


      for i := 0 to DBGridEh12.Columns.Count - 1 do
      begin
        DBGridEh12.Columns[i].Width := 60;
        DBGridEh12.Columns[i].Title.Alignment := taCenter;
        DBGridEh12.Columns[i].Alignment := taRightJustify;
      end;

      DBGridEh12.Columns[0].Alignment := taCenter;
      DBGridEh12.Columns[1].Alignment := taLeftJustify;

      DBGridEh12.Columns[0].Title.caption := '序号';
      DBGridEh12.Columns[1].Title.caption := '能源及能源介质';
      DBGridEh12.Columns[2].Title.caption := '单耗';
      DBGridEh12.Columns[3].Title.caption := '折标煤系数';
      DBGridEh12.Columns[4].Title.caption := '折标煤';

      DBGridEh12.Columns[0].Width := 40;
      DBGridEh12.Columns[1].Width := 100;


      DBGridEh12.SumList.Active := True;

      DBGridEh12.Columns[2].Footer.ValueType := fvtStaticText;
      DBGridEh12.Columns[2].Footer.Value := '工序能耗';

      DBGridEh12.Columns[3].Footer.ValueType := fvtStaticText;
      DBGridEh12.Columns[3].Footer.Value := 'kgce/' + s1;
      if s1 = 'kg' then
        DBGridEh12.Columns[3].Footer.Value := 'kgce/t';
      DBGridEh12.Columns[4].Footer.ValueType := fvtSum;


      Label35.Caption := DBGridEh12.Columns[3].Footer.Value;
      Label40.Caption := DBGridEh12.GetFooterValue(0, DBGridEh12.Columns[4]);





      if ADOQuery11.Active then ADOQuery11.Close;
      ADOQuery11.SQL.Clear;
      ADOQuery11.SQL.Add('select  ID,	source,	process, unit,	num,	remark from historydata where inf=' + #39 + '工序能耗' + #39 + ' order by source asc');
      ADOQuery11.Open;

      for i := 0 to DBGridEh13.Columns.Count - 1 do
      begin
        DBGridEh13.Columns[i].Width := 60;
        DBGridEh13.Columns[i].Title.Alignment := taCenter;
        DBGridEh13.Columns[i].Alignment := taRightJustify;
      end;

      DBGridEh13.Columns[0].Alignment := taCenter;
      DBGridEh13.Columns[1].Alignment := taLeftJustify;

      DBGridEh13.Columns[0].Title.caption := '序号';
      DBGridEh13.Columns[1].Title.caption := '来源';
      DBGridEh13.Columns[2].Title.caption := '工序';
      DBGridEh13.Columns[3].Title.caption := '单位';
      DBGridEh13.Columns[4].Title.caption := '工序能耗';
      DBGridEh13.Columns[5].Title.caption := '备注';
      DBGridEh13.Columns[0].Width := 40;
      DBGridEh13.Columns[1].Width := 180;
      DBGridEh13.Columns[2].Width := 120;


    end; //工序能耗


    if ComboBox7.ItemIndex = 1 then //碳排放
    begin
      Aqr1.Connection := ADOConnection5;
      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from shuxing where shuxing = ' + #39 + 'CO2排放系数' + #39);
      Aqr1.Open;

      Aqr4.Connection := ADOConnection1;
      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, shuxing from comparison where shuxing is not null ');
      Aqr4.Open;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
      Aqr2.Open;

      e := 0; a := 0;
      for j := 1 to Aqr2.RecordCount do
      begin
        Aqr2.RecNo := j;
        if Aqr4.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
        begin
          if Aqr1.Locate('name', Aqr4.FieldByName('shuxing').AsString, []) = true then
          begin
            if Aqr2.FieldByName('material_type').AsString = '产品' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '副产品或固体废弃物' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '能源及能源介质' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '原材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '辅助材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '大气排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '水体排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;


            ADOQuery10.Append;
            ADOQuery10.FieldByName('lcifactor').AsString := 'CO2排放';
            ADOQuery10.FieldByName('process').AsString := process;
            ADOQuery10.FieldByName('pd_name').AsString := Aqr2.FieldByName('pd_name').AsString;
  //        if s1='kg' then
     //       a:=1000*a;       //化成 t 为单位
            ADOQuery10.FieldByName('AF').AsFloat := a;
            ADOQuery10.FieldByName('EF').AsFloat := Aqr1.FieldByName('shuzhi').AsFloat;
            ADOQuery10.FieldByName('AFxEF').AsFloat := a * Aqr1.FieldByName('shuzhi').AsFloat;
            ADOQuery10.Post;
            e := e + a * Aqr1.FieldByName('shuzhi').AsFloat;
          end;
        end;
      end;

   //   bsSkinEdit4.Text := FormatFloat('#,##0.000', e);

      if ADOQuery10.Active then ADOQuery10.Close;
      ADOQuery10.SQL.Clear;
      ADOQuery10.SQL.Add('select ID,pd_name,AF,EF,AFxEF from temp order by AFxEF ASC');
      ADOQuery10.Open;


      for i := 0 to DBGridEh12.Columns.Count - 1 do
      begin
        DBGridEh12.Columns[i].Width := 60;
        DBGridEh12.Columns[i].Title.Alignment := taCenter;
        DBGridEh12.Columns[i].Alignment := taRightJustify;
      end;

      DBGridEh12.Columns[0].Alignment := taCenter;
      DBGridEh12.Columns[1].Alignment := taLeftJustify;

      DBGridEh12.Columns[0].Title.caption := '序号';
      DBGridEh12.Columns[1].Title.caption := '能源及能源介质';
      DBGridEh12.Columns[2].Title.caption := '单耗';
      DBGridEh12.Columns[3].Title.caption := 'CO2排放系数';
      DBGridEh12.Columns[4].Title.caption := 'CO2排放量';


      DBGridEh12.Columns[0].Width := 40;

      DBGridEh12.SumList.Active := True;

      DBGridEh12.Columns[2].Footer.ValueType := fvtStaticText;
      DBGridEh12.Columns[2].Footer.Value := 'CO2排放量';

      DBGridEh12.Columns[3].Footer.ValueType := fvtStaticText;
      DBGridEh12.Columns[3].Footer.Value := 'kgCO2/' + s1;

      DBGridEh12.Columns[4].Footer.ValueType := fvtSum; //


      Label35.Caption := DBGridEh12.Columns[3].Footer.Value;
      Label40.Caption := DBGridEh12.GetFooterValue(0, DBGridEh12.Columns[4]);


      if ADOQuery11.Active then ADOQuery11.Close;
      ADOQuery11.SQL.Clear;
      ADOQuery11.SQL.Add('select  ID,	source,	process, unit,	num,	remark from historydata where inf=' + #39 + '碳排放' + #39);
      ADOQuery11.Open;

      for i := 0 to DBGridEh13.Columns.Count - 1 do
      begin
        DBGridEh13.Columns[i].Width := 60;
        DBGridEh13.Columns[i].Title.Alignment := taCenter;
        DBGridEh13.Columns[i].Alignment := taRightJustify;
      end;

      DBGridEh13.Columns[0].Alignment := taCenter;
      DBGridEh13.Columns[1].Alignment := taLeftJustify;
      DBGridEh13.Columns[2].Alignment := taLeftJustify;


      DBGridEh13.Columns[0].Title.caption := '序号';
      DBGridEh13.Columns[1].Title.caption := '来源';
      DBGridEh13.Columns[2].Title.caption := '工序';
      DBGridEh13.Columns[3].Title.caption := '单位';
      DBGridEh13.Columns[4].Title.caption := '碳排放';
      DBGridEh13.Columns[5].Title.caption := '备注';
      DBGridEh13.Columns[0].Width := 40;

      DBGridEh13.Columns[1].Width := 180;
      DBGridEh13.Columns[2].Width := 120;




    end;


    if ComboBox7.ItemIndex = 2 then //铁平衡
    begin

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from shuxing where shuxing = ' + #39 + '铁含量' + #39);
      Aqr1.Open;

      if Aqr4.Active then Aqr4.Close;
      Aqr4.SQL.Clear;
      Aqr4.SQL.Add('select id, pd_name, shuxing from comparison where shuxing is not null ');
      Aqr4.Open;

      if Aqr2.Active then Aqr2.Close;
      Aqr2.SQL.Clear;
      Aqr2.SQL.Add('select * from data_allocation where process = ' + #39 + process + #39);
      Aqr2.Open;

      e := 0; a := 0;
      for j := 1 to Aqr2.RecordCount do
      begin
        Aqr2.RecNo := j;
        if Aqr4.Locate('pd_name', Aqr2.FieldByName('pd_name').AsString, []) = true then
        begin
          if Aqr1.Locate('name', Aqr4.FieldByName('shuxing').AsString, []) = true then
          begin

            if Aqr2.FieldByName('material_type').AsString = '产品' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '副产品或固体废弃物' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '能源及能源介质' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '原材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '辅助材料' then
              a := Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '大气排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;
            if Aqr2.FieldByName('material_type').AsString = '水体排放' then
              a := (-1) * Aqr2.FieldByName('co_num').AsFloat;

            ADOQuery10.Append;
            ADOQuery10.FieldByName('lcifactor').AsString := '工序能耗';
            ADOQuery10.FieldByName('process').AsString := process;
            ADOQuery10.FieldByName('pd_name').AsString := Aqr2.FieldByName('pd_name').AsString;
            ADOQuery10.FieldByName('AF').AsFloat := a;
            ADOQuery10.FieldByName('EF').AsFloat := Aqr1.FieldByName('shuzhi').AsFloat;
            ADOQuery10.FieldByName('AFxEF').AsFloat := a * Aqr1.FieldByName('shuzhi').AsFloat;
            ADOQuery10.Post;
            e := e + a * Aqr1.FieldByName('shuzhi').AsFloat;
          end;
        end;
      end;

      if ADOQuery10.Active then ADOQuery10.Close;
      ADOQuery10.SQL.Clear;
      ADOQuery10.SQL.Add('select ID,pd_name,AF,EF,AFxEF from temp order by AFxEF ASC');
      ADOQuery10.Open;


      for i := 0 to DBGridEh12.Columns.Count - 1 do
      begin
        DBGridEh12.Columns[i].Width := 60;
        DBGridEh12.Columns[i].Title.Alignment := taCenter;
        DBGridEh12.Columns[i].Alignment := taRightJustify;
      end;

      DBGridEh12.Columns[0].Alignment := taCenter;
      DBGridEh12.Columns[1].Alignment := taLeftJustify;

      DBGridEh12.Columns[0].Title.caption := '序号';
      DBGridEh12.Columns[1].Title.caption := '含铁物料';
      DBGridEh12.Columns[2].Title.caption := '单耗';
      DBGridEh12.Columns[3].Title.caption := '铁含量(%)';
      DBGridEh12.Columns[4].Title.caption := '铁含量(kg)';


      DBGridEh12.Columns[0].Width := 40;

      DBGridEh12.SumList.Active := True;

      DBGridEh12.Columns[2].Footer.ValueType := fvtStaticText;
      DBGridEh12.Columns[2].Footer.Value := '铁平衡误差';

      DBGridEh12.Columns[3].Footer.ValueType := fvtStaticText;
      DBGridEh12.Columns[3].Footer.Value := 'kg/kg';

      DBGridEh12.Columns[4].Footer.ValueType := fvtSum; //

    end;


    if ComboBox7.ItemIndex = 3 then //物料平衡
      if checkedcount(bsSkinCheckListBox55) > 0 then
      begin

        if ADOQuery10.Active then ADOQuery10.Close;
        ADOQuery10.SQL.Clear;
        mysql := 'select id, process, material_type, pd_name ,pd_unit,pd from data_allocation where (pd_name=  ';

        for i := 0 to bsSkinCheckListBox55.Items.Count - 1 do
        begin
          if bsSkinCheckListBox55.Checked[i] = true then
          begin
            mysql := mysql + #39 + bsSkinCheckListBox55.Items.Strings[i] + #39;
            break;
          end;
        end;

        for j := i + 1 to bsSkinCheckListBox55.Items.Count - 1 do
          if bsSkinCheckListBox55.Checked[j] then
            mysql := mysql + ' or  pd_name= ' + #39 + bsSkinCheckListBox55.Items.Strings[j] + #39;

        mysql := mysql + ')';

        ADOQuery10.SQL.Add(mysql);
        ADOQuery10.Open;

        for i := 0 to DBGridEh12.Columns.Count - 1 do
        begin
          DBGridEh12.Columns[i].Width := 60;
          DBGridEh12.Columns[i].Title.Alignment := taCenter;
          DBGridEh12.Columns[i].Alignment := taRightJustify;
        end;

        DBGridEh12.Columns[0].Alignment := taCenter;
        DBGridEh12.Columns[1].Alignment := taLeftJustify;

        DBGridEh12.Columns[0].Title.caption := '序号';
        DBGridEh12.Columns[1].Title.caption := '工序';
        DBGridEh12.Columns[2].Title.caption := '类别';
        DBGridEh12.Columns[3].Title.caption := '名称';
        DBGridEh12.Columns[4].Title.caption := '单位';
        DBGridEh12.Columns[5].Title.caption := '数量';

        DBGridEh12.Columns[0].Width := 30;





      end;

  finally


    Aqr1.Close;
    Aqr1.Free;
    Aqr2.Close;
    Aqr2.Free;
    Aqr3.Close;
    Aqr3.Free;
    Aqr4.Close;
    Aqr4.Free;

  end;

end;

procedure TForm1.ADOQuery10AfterOpen(DataSet: TDataSet);
var
  i: integer;
begin
  ADOQuery10.FieldByName('ID').OnGetText := SetFdText10;
  for i := 1 to ADOQuery10.FieldCount - 1 do
    if ADOQuery10.Fields[i].DataType = ftFloat then
      (ADOQuery10.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0';

  ADOQuery10.First;


end;

procedure TForm1.bsSkinButton76Click(Sender: TObject);
var
  i, j: integer;
  Aqr1: TAdoQuery;
  s, product1, process, str1, str2: string;
begin

  if ADOQuery10.Active = false then exit;
  if ADOQuery10.RecordCount = 0 then exit;

  Aqr1 := TAdoQuery.Create(nil);
  try
    for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
    begin
      if bsSkinCheckListBox8.Checked[i] = true then
      begin
        s := bsSkinCheckListBox8.Items[i];
        break;
      end;
    end;


    for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
    begin
      if bsSkinCheckListBox5.Checked[i] = true then
      begin
        product1 := bsSkinCheckListBox5.Items[i];
        break;
      end;
    end;

  //  ADOConnection5.Close;
  //  ADOConnection5.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + s + '\' + s + '.mdb ; Persist Security Info=False';
 //   ADOConnection5.Open;

    Aqr1.Connection := ADOConnection5;
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from data_allocation  where  material_type= ' + #39 + '产品' + #39 + ' and pd_name=' + #39 + product1 + #39);
    Aqr1.Open;
    process := Aqr1.FieldByName('process').AsString;

    Aqr1.Close;
    Aqr1.Connection := ADOConnection1; //连接对标数据库

    if ComboBox7.ItemIndex = 0 then //工序能耗
    begin

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from historydata where inf=' + #39 + '工序能耗' + #39);
      Aqr1.Open;

      if Aqr1.Locate('source; process ', VarArrayOf([s, process]), []) = false then
      begin
        Aqr1.Append;
        Aqr1.FieldByName('source').AsString := s;
        Aqr1.FieldByName('process').AsString := process;
        Aqr1.FieldByName('inf').AsString := '工序能耗';
        Aqr1.FieldByName('unit').AsString := Label35.Caption;
        str1 := Label40.Caption;
        str2 := ',';
        while pos(str2, str1) <> 0 do
          delete(str1, pos(str2, str1), length(str2));


        Aqr1.FieldByName('num').AsString := str1;
        Aqr1.Post;
      end
      else
      begin
        if form1.bsSkinMessage1.MessageDlg('该数据已存在，是否更新数据？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          Aqr1.edit;
          Aqr1.FieldByName('source').AsString := s;
          Aqr1.FieldByName('process').AsString := process;
          Aqr1.FieldByName('inf').AsString := '工序能耗';
          Aqr1.FieldByName('unit').AsString := Label35.Caption;

          str1 := Label40.Caption;
          str2 := ',';
          while pos(str2, str1) <> 0 do
            delete(str1, pos(str2, str1), length(str2));


          Aqr1.FieldByName('num').AsString := str1;



          Aqr1.Post;
        end;
      end;

    end; //工序能耗


    if ComboBox7.ItemIndex = 1 then //碳排放
    begin

      if Aqr1.Active then Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add('select * from historydata where inf=' + #39 + '碳排放' + #39);
      Aqr1.Open;

      if Aqr1.Locate('source; process ', VarArrayOf([s, process]), []) = false then
      begin
        Aqr1.Append;
        Aqr1.FieldByName('source').AsString := s;
        Aqr1.FieldByName('process').AsString := process;
        Aqr1.FieldByName('inf').AsString := '碳排放';
        Aqr1.FieldByName('unit').AsString := Label35.Caption;
        str1 := Label40.Caption;
        str2 := ',';
        while pos(str2, str1) <> 0 do
          delete(str1, pos(str2, str1), length(str2));


        Aqr1.FieldByName('num').AsString := str1;
        Aqr1.Post;
      end
      else
      begin
        if form1.bsSkinMessage1.MessageDlg('该数据已存在，是否更新数据？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          Aqr1.edit;
          Aqr1.FieldByName('source').AsString := s;
          Aqr1.FieldByName('process').AsString := process;
          Aqr1.FieldByName('inf').AsString := '碳排放';
          Aqr1.FieldByName('unit').AsString := Label35.Caption;
          str1 := Label40.Caption;
          str2 := ',';
          while pos(str2, str1) <> 0 do
            delete(str1, pos(str2, str1), length(str2));


          Aqr1.FieldByName('num').AsString := str1;
          Aqr1.Post;
        end;
      end;

    end; //碳排放




    ADOQuery11.Close;
    ADOQuery11.Open;

    for i := 0 to DBGridEh13.Columns.Count - 1 do
    begin
      DBGridEh13.Columns[i].Width := 60;
      DBGridEh13.Columns[i].Title.Alignment := taCenter;
      DBGridEh13.Columns[i].Alignment := taRightJustify;
    end;

    DBGridEh13.Columns[0].Alignment := taCenter;
    DBGridEh13.Columns[1].Alignment := taLeftJustify;

    DBGridEh13.Columns[0].Title.caption := '序号';
    DBGridEh13.Columns[1].Title.caption := '来源';
    DBGridEh13.Columns[2].Title.caption := '工序';
    DBGridEh13.Columns[3].Title.caption := '单位';
    if ComboBox7.ItemIndex = 0 then
      DBGridEh13.Columns[4].Title.caption := '工序能耗';
    if ComboBox7.ItemIndex = 1 then //碳排放
      DBGridEh13.Columns[4].Title.caption := '碳排放';
    DBGridEh13.Columns[5].Title.caption := '备注';
    DBGridEh13.Columns[0].Width := 40;
    DBGridEh13.Columns[1].Width := 180;
    DBGridEh13.Columns[2].Width := 120;


    ADOQuery11.Locate('source; process ', VarArrayOf([s, process]), []);
    DBGridEh13.SelectedRows.CurrentRowSelected := true;

  finally
    Aqr1.Close;
    Aqr1.Free;
  end;

end;

procedure TForm1.ADOQuery11AfterOpen(DataSet: TDataSet);
var
  i: integer;
begin
  ADOQuery11.FieldByName('ID').OnGetText := SetFdText11;
  for i := 1 to ADOQuery11.FieldCount - 1 do
    if ADOQuery11.Fields[i].DataType = ftFloat then
      (ADOQuery11.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0';

  ADOQuery11.First;

end;

procedure TForm1.bsSkinButton77Click(Sender: TObject);
var
  i: integer;
  rootnode: ttreenode;
begin
  Panel7.Visible := false;
  TreeView_SetItemHeight(TreeView1.Handle, 30);
  TreeView1.Hint := '基础信息管理';
  TreeView1.Items.Clear;
  PageControl1.Visible := true;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := false;
  PageControl1.ActivePageIndex := 12;
  bsSkinButton24.Visible := false;

  rootnode := TreeView1.items.add(nil, '基础信息管理');
  TreeView1.Items.Addchild(rootnode, '外部LCA数据管理');
  TreeView1.Items.Addchild(rootnode, '运输LCA基础数据管理');
  TreeView1.Items.Addchild(rootnode, 'LCI环境指标管理');

  TreeView1.Items.Addchild(rootnode, '环境影响类型管理');

  rootnode := TreeView1.items.add(nil, '对照表管理');
  TreeView1.Items.Addchild(rootnode, '物料属性名称对照表');
  TreeView1.Items.Addchild(rootnode, '运输名称对照表');
  TreeView1.Items.Addchild(rootnode, '上游数据名称对照表');
  TreeView1.Items.Addchild(rootnode, 'LCI因子名称对照表');
  rootnode := TreeView1.items.add(nil, '工序对标数据管理');

  TreeView1.FullExpand;
  TreeView1.Items[10].Selected := true;
  TreeView1.SetFocus;
  StatusBar1.Panels[1].Text := '基础信息管理： ' + TreeView1.Selected.Text;

  Image1.Visible := false;
  bsSkinPanel2.Visible := false;

  label4.Caption := '工序对标数据管理';

  ListBox4.Clear;

  if ADOQuery11.Active then ADOQuery11.Close;
  ADOQuery11.SQL.Clear;
  ADOQuery11.SQL.Add('select * from historydata');
  ADOQuery11.Open;
  for i := 1 to ADOQuery11.RecordCount do
  begin
    ADOQuery11.RecNo := i;
    if ListBox4.Items.IndexOf(ADOQuery11.FieldByName('source').AsString) = -1 then
      ListBox4.Items.Add(ADOQuery11.FieldByName('source').AsString);
  end;


  AdoQuery1.Close;
  ListBox4Click(Sender);

end;

procedure TForm1.bsSkinCheckListBox55ClickCheck(Sender: TObject);
begin
  if checkedcount(bsSkinCheckListBox55) = 0 then exit;
  Button3Click(Sender);
end;

procedure TForm1.bsSkinButton74Click(Sender: TObject);
var
  i, j: integer;
  s, s1, product1, process: string;
begin

  if DBGridEh13.DataSource.DataSet.RecordCount = 0 then exit;

  for i := 0 to bsSkinCheckListBox8.Items.Count - 1 do
  begin
    if bsSkinCheckListBox8.Checked[i] = true then
    begin
      s := bsSkinCheckListBox8.Items[i];
      break;
    end;
  end;

  for i := 0 to bsSkinCheckListBox5.Items.Count - 1 do
  begin
    if bsSkinCheckListBox5.Checked[i] = true then
    begin
      product1 := bsSkinCheckListBox5.Items[i];
      break;
    end;
  end;


 //     ADOQuery11.SQL.Add('select  ID,	source,	process, unit,	num,	remark from historydata where inf=' + #39 + '碳排放' + #39);





  chart2.Title.Text.Clear;
  chart2.Title.Text.Add('[  ' + product1 + '  ]  ' + DBGridEh13.Columns[4].Title.caption + '对标(' + Label35.Caption + ')');
  if chart2.Series[0].Count = 0 then
  begin
    chart2.Series[0].Add(strtofloat(Label40.Caption), s, clteecolor);
    chart2.Series[0].Add(DBGridEh13.DataSource.DataSet.FieldByName('num').AsFloat, DBGridEh13.DataSource.DataSet.FieldByName('source').AsString + '/' + DBGridEh13.DataSource.DataSet.FieldByName('process').AsString, clteecolor);
  end
  else
  begin
    chart2.Series[0].Add(DBGridEh13.DataSource.DataSet.FieldByName('num').AsFloat, DBGridEh13.DataSource.DataSet.FieldByName('source').AsString + '/' + DBGridEh13.DataSource.DataSet.FieldByName('process').AsString, clteecolor);
  end;






end;

procedure TForm1.bsSkinButton56Click(Sender: TObject);
begin
  if chart2.Series[0].Count > 0 then
    chart2.Series[0].Delete(chart2.Series[0].Count - 1);
end;

procedure TForm1.N12Click(Sender: TObject);
begin
  if TreeView1.Hint = '基础信息管理' then
    exit;


  TreeView1Click(Sender);
  PageControl1.TabIndex := 10;
  PageControl1Change(Sender);
end;

procedure TForm1.bsSkinButton2Click(Sender: TObject);
begin
  if PageControl1.Visible = false then exit;
  PageControl1.ActivePageIndex := 9;
  PageControl1Change(Sender);
end;

procedure TForm1.N8Click(Sender: TObject);
begin
  bsSkinButton77Click(Sender);
end;

procedure TForm1.N11Click(Sender: TObject);
var
  Aqr1: TAdoQuery;
  i: integer;
begin
  Aqr1 := TAdoQuery.Create(nil);
  try
    Aqr1.Connection := form1.ADOConnection1;
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from alloy order by alloy');
    Aqr1.Open;

    bsSkinCheckListBox1.items.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      eco.bsSkinCheckListBox1.Items.Add(Aqr1.Fields[1].AsString);
    end;

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select distinct cases from alloydesign ');
    Aqr1.Open;

    eco.bsSkinCheckComboBox1.Items.Clear;
    eco.ComboBox1.Items.Clear;
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      eco.bsSkinCheckComboBox1.Items.Add(Aqr1.Fields[0].AsString);
      eco.ComboBox1.Items.Add(Aqr1.Fields[0].AsString);
    end;


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from lcifactor ');
    Aqr1.Open;
    eco.ComboBox2.Items.Clear;
    eco.ComboBox2.Items.Add('');
    for i := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := i;
      eco.ComboBox2.Items.Add(Aqr1.FieldByName('inf_name').AsString);
    end;







  finally
    Aqr1.Close;
    Aqr1.Free;

  end;
  eco.PageControl1.ActivePageIndex := 0;
  eco.ShowModal;

end;

procedure TForm1.RadioButton3Click(Sender: TObject);
begin
  if RadioButton3.Checked then
  begin
    ComboBox4.Enabled := false;
    bsSkinCheckListBox2.Clear;

    bsSkinCheckListBox2.Items.Add('basedonLCI');

    CheckBox5.Checked := true;
    CheckBox5Click(Sender);

  end;
  Button2Click(Sender); //LCA结果查询
end;

procedure TForm1.bsSkinButton55Click(Sender: TObject);
var
  i, j: integer;
  mysql: string;
  Aqr1, Aqr2, Aqr3: TAdoQuery;
  a, b: double;
begin
  Aqr3 := TAdoQuery.Create(nil);
  Aqr3.Connection := ADOConnection2;
  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := ADOConnection2;
  Aqr2 := TAdoQuery.Create(nil);
  Aqr2.Connection := ADOConnection2;



  if RadioButton11.checked = true then
  begin

    mysql := 'select * from  CO2emission_quanchang ';
    if AdoQuery12.Active then AdoQuery12.Close;
    AdoQuery12.SQL.Clear;
    AdoQuery12.SQL.Add(mysql);
    ADOQuery12.Open;
    for i := 1 to ADOQuery12.RecordCount do
    begin
      ADOQuery12.RecNo := i;
      ADOQuery12.Edit;
      if ADOQuery12.FieldByName('beizhu').AsString <> '合计' then
        ADOQuery12.FieldByName('paifangliang').AsFloat := ADOQuery12.FieldByName('shiyongliang').AsFloat * ADOQuery12.FieldByName('paifangyinzi').AsFloat;
      ADOQuery12.Post;
    end;

   // ADOQuery12.Refresh;


    for i := 1 to ADOQuery12.RecordCount do
    begin
      ADOQuery12.RecNo := i;
      if ADOQuery12.FieldByName('beizhu').AsString = '合计' then
      begin
        ADOQuery12.Edit;
        ADOQuery12.FieldByName('paifangliang').AsFloat := 0;

        if Aqr3.Active then Aqr3.Close;
        Aqr3.SQL.Clear;
        Aqr3.SQL.Add(mysql + ' where (processunit=' + #39 + ADOQuery12.FieldByName('processunit').AsString + #39 + ')');
        Aqr3.Open; //  + ' and beizhu <>' + #39 + '合计' + #39


        for j := 1 to Aqr3.RecordCount do
        begin
          Aqr3.RecNo := j;
          if Aqr3.FieldByName('beizhu').AsString <> '合计' then
            ADOQuery12.FieldByName('paifangliang').AsFloat := ADOQuery12.FieldByName('paifangliang').AsFloat + Aqr3.FieldByName('paifangliang').AsFloat;
        end;

        ADOQuery12.Post;
      end;

    end;



    if Aqr3.Active then Aqr3.Close;
    Aqr3.SQL.Clear;
    Aqr3.SQL.Add(mysql + ' where beizhu=' + #39 + '合计' + #39);
    Aqr3.Open;
    ADOQuery12.Locate('processunit', '排放量汇总', [loCaseInsensitive]);
    ADOQuery12.Edit;
    ADOQuery12.FieldByName('paifangliang').AsFloat := 0;
    for j := 1 to Aqr3.RecordCount do
    begin
      Aqr3.RecNo := j;
      ADOQuery12.FieldByName('paifangliang').AsFloat := ADOQuery12.FieldByName('paifangliang').AsFloat + Aqr3.FieldByName('paifangliang').AsFloat;
    end;
    ADOQuery12.post;

    ADOQuery12.Locate('processunit', '排放量汇总', [loCaseInsensitive]);
    FlatEdit19.Text := formatfloat('##0', ADOQuery12.FieldByName('paifangliang').AsFloat);


  end;



  if Aqr3.Active then Aqr3.Close;
  Aqr3.SQL.Clear;
  Aqr3.SQL.Add('select * from  CO2emission_quanchang ');
  Aqr3.Open;
  Aqr3.Locate('processunit', '排放量汇总', [loCaseInsensitive]);
  FlatEdit19.Text := formatfloat('##0', Aqr3.FieldByName('paifangliang').AsFloat);



  if Aqr3.Active then Aqr3.Close;
  Aqr3.SQL.Clear;
  Aqr3.SQL.Add('delete * from  CO2emission where beizhu = ' + #39 + 'allocation' + #39);
  Aqr3.ExecSQL;

  if Aqr3.Active then Aqr3.Close;
  Aqr3.SQL.Clear;
  Aqr3.SQL.Add('select * from  CO2emission ');
  Aqr3.Open;

  if Aqr1.Active then Aqr1.Close;
  Aqr1.SQL.Clear;
  Aqr1.SQL.Add('select * from  CO2_daleichanpin');
  Aqr1.Open;

  for i := 1 to Aqr1.RecordCount do
  begin
    Aqr1.RecNo := i;
    Aqr3.Locate('product', Aqr1.FieldByName('xilei').AsString, [loCaseInsensitive]);
    Aqr3.Edit;
    Aqr3.FieldByName('beizhu').AsString := Aqr1.FieldByName('dalei').AsString;
    Aqr3.Post;
  end;

  if Aqr1.Active then Aqr1.Close;
  Aqr1.SQL.Clear;
  Aqr1.SQL.Add('select distinct dalei from  CO2_daleichanpin');
  Aqr1.Open;

  for i := 1 to Aqr1.RecordCount do
  begin
    Aqr1.RecNo := i;
    if Aqr2.Active then Aqr2.Close;
    Aqr2.SQL.Clear;
    Aqr2.SQL.Add('select sum(CO2Total) as CO2_Total,sum(waimailiang) as waimailiang1  from CO2emission where beizhu=' + #39 + Aqr1.FieldByName('dalei').AsString + #39);
    Aqr2.Open;
    Aqr3.Append;
    Aqr3.FieldByName('product').AsString := Aqr1.FieldByName('dalei').AsString;
    Aqr3.FieldByName('waimailiang').AsFloat := Aqr2.FieldByName('waimailiang1').AsFloat / 1000;
    Aqr3.FieldByName('CO2Total').AsFloat := Aqr2.FieldByName('CO2_Total').AsFloat;
    Aqr3.FieldByName('beizhu').AsString := 'allocation';
    Aqr3.Post;
  end;


                                                               //    where beizhu <> ' + #39 + ' allocation' + #39

  if Aqr3.Active then Aqr3.Close;
  Aqr3.SQL.Clear;
  Aqr3.SQL.Add('select sum(CO2Total) as CO2_Total from CO2emission ');
  Aqr3.Open;
  if Aqr2.Active then Aqr2.Close;
  Aqr2.SQL.Clear;
  Aqr2.SQL.Add('select sum(CO2Total) as CO2_Total from CO2emission where beizhu=' + #39 + 'allocation' + #39);
  Aqr2.Open;


  if Aqr3.RecordCount > 0 then
    FlatEdit21.Text := formatfloat('##0', Aqr3.FieldByName('CO2_Total').AsFloat - Aqr2.FieldByName('CO2_Total').AsFloat)
  else
    FlatEdit21.Text := '0';




       // 差额 * 比例  + CO2total
       //差额 = 全部CO2求和 -  大类CO2求和
       //比例=   大类CO2 / 大类CO2求和

  if Aqr2.Active then Aqr2.Close;
  Aqr2.SQL.Clear;
  Aqr2.SQL.Add('select sum(CO2Total) as CO2_Total from CO2emission where beizhu=' + #39 + 'allocation' + #39);
  Aqr2.Open;
  a := Aqr2.FieldByName('CO2_Total').AsFloat; // 大类CO2求和
  b := strtofloat(FlatEdit21.Text) - a; //差额
 // showmessage(floattostr(a) + '  bbbb  ' + floattostr(b));

  if Aqr3.Active then Aqr3.Close;
  Aqr3.SQL.Clear;
  Aqr3.SQL.Add('select * from  CO2emission where beizhu=' + #39 + 'allocation' + #39);
  Aqr3.Open;
  for i := 1 to Aqr3.RecordCount do
  begin
    Aqr3.RecNo := i;
    Aqr3.Edit;
    try
      Aqr3.FieldByName('CO2Total_allocation').AsFloat := Aqr3.FieldByName('CO2Total').AsFloat * (1 + b / a);
      Aqr3.FieldByName('LCI_allocation').AsFloat := Aqr3.FieldByName('CO2Total_allocation').AsFloat / Aqr3.FieldByName('waimailiang').AsFloat;
    except
    end;
    Aqr3.Post;
  end;

  try
    FlatEdit22.Text := formatfloat('#0.#0', (strtofloat(FlatEdit21.Text) / strtofloat(FlatEdit19.Text) - 1) * 100) + '  %'

  except

  end;

 {
  for i := 0 to DBGridEh14.Columns.Count - 1 do
    DBGridEh14.Columns[i].Width := 120;
  DBGridEh14.Columns[0].Width := 30;
  DBGridEh14.Columns[1].Width := 150;  }

  RadioButton11.Checked := true;
  RadioButton11Click(Sender);


  Aqr3.Close;
  Aqr3.Free;
  Aqr2.Close;
  Aqr2.Free;

  Aqr1.Close;
  Aqr1.Free;
end;

procedure TForm1.ADOQuery12AfterOpen(DataSet: TDataSet);
var
  i: integer;
begin
  ADOQuery12.FieldByName('ID').OnGetText := SetFdText12;
  for i := 1 to ADOQuery12.FieldCount - 1 do
    if ADOQuery12.Fields[i].DataType = ftFloat then
      (ADOQuery12.Fields[i] as Tfloatfield).DisplayFormat := '###,###,##0.##0';

  ADOQuery12.First;

end;

procedure TForm1.DBGridEh14TitleClick(Column: TColumnEh);
begin
  if flag then
    ADOQuery12.Sort := Column.FieldName + ' ASC'
  else
    ADOQuery12.Sort := Column.FieldName + ' DESC';

  flag := not flag;
end;

procedure TForm1.DBGridEh14GetCellParams(Sender: TObject;
  Column: TColumnEh; AFont: TFont; var Background: TColor;
  State: TGridDrawState);
begin
  if not (gdSelected in State) then
    if (Sender as TDBGridEh).SumList.RecNo mod 2 = 0 then
      Background := clInfobk;
end;

procedure TForm1.RadioButton11Click(Sender: TObject);
var
  i: integer;
begin

  if AdoQuery12.Active then AdoQuery12.Close;
  AdoQuery12.SQL.Clear;
  AdoQuery12.SQL.Add('select * from  CO2emission_quanchang ');
  ADOQuery12.Open;
  for i := 0 to DBGridEh14.Columns.Count - 1 do
    DBGridEh14.Columns[i].Width := 120;
  DBGridEh14.Columns[0].Width := 30;
  DBGridEh14.Columns[1].Width := 150;


  DBGridEh14.Columns[1].Title.caption := '单元';
  DBGridEh14.Columns[2].Title.caption := '数据项';
  DBGridEh14.Columns[3].Title.caption := '使用量';
  DBGridEh14.Columns[4].Title.caption := '排放因子';
  DBGridEh14.Columns[5].Title.caption := '排放量';
  DBGridEh14.Columns[6].Title.caption := '备注';
end;

procedure TForm1.RadioButton12Click(Sender: TObject);
var
  i: integer;
begin
  if AdoQuery12.Active then AdoQuery12.Close;
  AdoQuery12.SQL.Clear;
  AdoQuery12.SQL.Add('select ID, product,waimailiang,LCI_allocation, CO2Total_allocation from  CO2emission where beizhu=' + #39 + 'allocation' + #39);
  ADOQuery12.Open;



  for i := 0 to DBGridEh14.Columns.Count - 1 do
    DBGridEh14.Columns[i].Width := 100;
  DBGridEh14.Columns[0].Width := 30;
  DBGridEh14.Columns[1].Width := 150;


  DBGridEh14.Columns[1].Title.caption := '产品';
  DBGridEh14.Columns[2].Title.caption := '外卖量(t)';
  DBGridEh14.Columns[3].Title.caption := 'LCI排放系数(t/t)';
  DBGridEh14.Columns[4].Title.caption := 'CO2排放合计(t)';

 //  DBGridEh14.Columns[2].Title.caption := '产量';
//  DBGridEh14.Columns[3].Title.caption := '耗量';
//  DBGridEh14.Columns[6].Title.caption := '内部排放合计';
 // DBGridEh14.Columns[7].Title.caption := '外购电使用量';
//  DBGridEh14.Columns[8].Title.caption := '外购电CO2排放量';



end;

end.

