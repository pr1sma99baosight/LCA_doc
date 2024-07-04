unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, comobj;

type
  TForm1 = class(TForm)
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    ADOConExcel1: TADOConnection;
    ADOConExcel2: TADOConnection;
    Edit2: TEdit;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    Edit1: TEdit;
    Edit3: TEdit;
    ADOQuery3: TADOQuery;
    procedure Button1Click(Sender: TObject);
    procedure Edit2Change(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure Edit3Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  v, v1, v2, sheet0: Variant;
  i, j: integer;
  s1, s2: string;

begin

  try
    v1 := CreateOleObject('Excel.Application'); //创建ole对象
    v2 := CreateOleObject('Excel.Application'); //创建ole对象

  except
    showmessage('未安装EXCEL');
  end;

  try
    v1.WorkBooks.open(edit2.Text);
    s1 := v1.WorkBooks[1].WorkSheets[1].name; //下载
  finally
    v1.WorkBooks[1].Close;
    v1.quit;
    v1 := Unassigned;
  end;

  try
    v2.WorkBooks.open(edit1.Text);
    s2 := v2.WorkBooks[1].WorkSheets[1].name; //下载
  finally
    v2.WorkBooks[1].Close;
    v2.quit;
    v2 := Unassigned;
  end;


  try

    ADOConExcel1.Close;
    ADOConExcel1.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + edit2.Text + ';Extended Properties=excel 8.0;Persist Security Info=False ';
    ADOConExcel1.Open;

    if ADOQuery1.Active then ADOQuery1.Close;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT * FROM [' + s1 + '$]'); //连下载文件
    ADOQuery1.open; //up

    ADOConExcel2.Close;
    ADOConExcel2.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + edit1.Text + ';Extended Properties=excel 8.0;Persist Security Info=False ';
    ADOConExcel2.Open;

    if ADOQuery2.Active then ADOQuery2.Close;
    ADOQuery2.SQL.Clear;
    ADOQuery2.SQL.Add('SELECT * FROM [' + s2 + '$]');
    ADOQuery2.open; //db



    for j := 5 to ADOQuery1.FieldCount - 1 do
    begin
      for i := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := i;
        ADOQuery2.Append;
        ADOQuery2.Fields[0].AsInteger := j;
        ADOQuery2.Fields[1].AsString := ADOQuery1.Fields[j].FieldName;
        ADOQuery2.Fields[2].AsString := ADOQuery1.Fields[0].AsString;
        ADOQuery2.Fields[3].AsString := ADOQuery1.Fields[1].AsString;

        ADOQuery2.Fields[4].AsString := ADOQuery1.Fields[j].AsString;

        ADOQuery2.Fields[5].AsString := ADOQuery1.Fields[2].AsString;
        ADOQuery2.Fields[6].AsString := ADOQuery1.Fields[3].AsString;
        ADOQuery2.Fields[7].AsString := ADOQuery1.Fields[4].AsString;

         ADOQuery2.Post;



    //   product	LCIfactor	unit	sum	function_unit	LCItype	see_level
       //       1      2       3    4     5               6    7

   //   lci	Units	function_unit	LCItype	see_level	煤

   //     0    1    2              3     4         5


      end;

    end





  finally
    ADOConExcel1.Close;
    ADOConExcel2.Close;


    v1.quit;
    v1 := Unassigned;
    v2.quit;
    v2 := Unassigned;


  end;






end;

procedure TForm1.Edit2Change(Sender: TObject);
begin
  if OpenDialog1.Execute then
    edit2.Text := OpenDialog1.FileName
  else exit;

end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
  if OpenDialog1.Execute then
    edit1.Text := OpenDialog1.FileName
  else exit;
end;

procedure TForm1.Edit3Change(Sender: TObject);
begin
  if OpenDialog1.Execute then
    edit3.Text := OpenDialog1.FileName
  else exit;
end;

end.

