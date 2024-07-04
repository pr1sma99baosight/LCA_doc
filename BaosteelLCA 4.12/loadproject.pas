unit loadproject;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, bsSkinCtrls, StdCtrls, flowchart, DB, ADODB;

type
  TForm_loadproject = class(TForm)
    Label1: TLabel;
    bsSkinPanel4: TbsSkinPanel;
    ListBox1: TListBox;
    bsSkinButton9: TbsSkinButton;
    bsSkinButton1: TbsSkinButton;
    Label2: TLabel;
    ADOQuery7: TADOQuery;
    ADOConnection3: TADOConnection;
    procedure bsSkinButton1Click(Sender: TObject);
    procedure bsSkinButton9Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form_loadproject: TForm_loadproject;

implementation

uses Unit1;

{$R *.dfm}

procedure TForm_loadproject.bsSkinButton1Click(Sender: TObject);
begin
  close;
end;

procedure TForm_loadproject.bsSkinButton9Click(Sender: TObject);
var
  i, j, k: integer;
begin
  if ListBox1.ItemIndex = -1 then
  begin
    form1.bsSkinMessage1.MessageDlg('请选择项目', (mtInformation), [mbOK], 0);
    exit;
  end;

  if Label2.Caption = '载入流程图' then
  begin
    form2.dxFlowChart1.LoadFromFile(ExtractFilePath(Application.ExeName) + 'projects\' + listbox1.Items[listbox1.ItemIndex] + '\' + listbox1.Items[listbox1.ItemIndex] + '.flc');
    form2.ListBox1.Clear;

    for i := 0 to form2.dxFlowChart1.ObjectCount - 1 do
      if form2.dxFlowChart1.Objects[i].Text <> '主生产工序' then
        if trim(form2.dxFlowChart1.Objects[i].Text) <> '能源公辅工序' then
          if form2.dxFlowChart1.Objects[i].Text <> '至各工序' then
            form2.ListBox1.Items.Add(form2.dxFlowChart1.Objects[i].Text);
    close;
  end;

  if Label2.Caption = '载入目录树' then
  begin
    form2.TreeView1.LoadFromFile(ExtractFilePath(Application.ExeName) + 'projects\' + listbox1.Items[listbox1.ItemIndex] + '\' + 'tree.txt');
    close;
  end;

  if Label2.Caption = '载入其它项目数据（全部工序）' then
  begin
    ADOConnection3.Close;
    ADOConnection3.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + ExtractFilePath(Application.ExeName) + 'projects\' + listbox1.Items[listbox1.ItemIndex] + '\' + listbox1.Items[listbox1.ItemIndex] + '.mdb ; Persist Security Info=False';
    ADOConnection3.Open;
    ADOQuery7.Connection := ADOConnection3;

    for i := 1 to form1.ListBox2.Items.Count - 1 do
    begin
      ADOQuery7.Close;
      ADOQuery7.SQL.Clear;
      ADOQuery7.SQL.Add('select * from data where process=' + #39 + form1.ListBox2.Items[i] + #39);
      ADOQuery7.Open;


      if ADOQuery7.RecordCount > 0 then
      begin
        form1.ADOQuery1.Close;
        form1.ADOQuery1.SQL.Clear;
        form1.ADOQuery1.SQL.Add('delete * from data where process=' + #39 + form1.ListBox2.Items[i] + #39);
        form1.ADOQuery1.ExecSQL;

        form1.ADOQuery1.Close;
        form1.ADOQuery1.SQL.Clear;
        form1.ADOQuery1.SQL.Add('select * from data ');
        form1.ADOQuery1.Open;

        for j := 1 to ADOQuery7.RecordCount do
        begin
          ADOQuery7.RecNo := j;
          form1.ADOQuery1.Append;
          for k := 1 to ADOQuery7.FieldCount - 1 do
            form1.ADOQuery1.Fields[k].AsString := ADOQuery7.Fields[k].AsString;
          form1.ADOQuery1.Post;

        end;



      end;
    end;




    form1.ProcessData_Format;
    close;
  end;
end;

end.

