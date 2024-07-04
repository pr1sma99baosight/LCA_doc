unit ecodesign;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, GridsEh, DBGridEh, bsSkinCtrls, bsSkinBoxCtrls, DB,
  ADODB, ComCtrls, bsSkinTabs, bsSkinGrids, Grids, matrix, ExtCtrls;

type
  Teco = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Edit1: TEdit;
    Label3: TLabel;
    bsSkinCheckListBox1: TbsSkinCheckListBox;
    bsSkinButton5: TbsSkinButton;
    bsSkinButton6: TbsSkinButton;
    bsSkinButton4: TbsSkinButton;
    DBGridEh2: TDBGridEh;
    TabSheet2: TTabSheet;
    Label1: TLabel;
    bsSkinButton14: TbsSkinButton;
    bsSkinButton72: TbsSkinButton;
    Label2: TLabel;
    Label4: TLabel;
    bsSkinButton2: TbsSkinButton;
    bsSkinButton3: TbsSkinButton;
    bsSkinCheckComboBox1: TbsSkinCheckComboBox;
    Label5: TLabel;
    Label6: TLabel;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    DataSource2: TDataSource;
    ComboBox1: TComboBox;
    bsSkinButton8: TbsSkinButton;
    bsSkinButton1: TbsSkinButton;
    StringGrid1: TStringGrid;
    CheckBox1: TCheckBox;
    bsSkinButton7: TbsSkinButton;
    ComboBox2: TComboBox;
    Matrix1: TMatrix;
    Label7: TLabel;
    ComboBox3: TComboBox;
    bsSkinCheckListBox2: TbsSkinCheckListBox;
    bsSkinButton9: TbsSkinButton;
    bsSkinButton10: TbsSkinButton;
    bsSkinButton11: TbsSkinButton;
    Panel1: TPanel;
    bsSkinButton12: TbsSkinButton;
    procedure bsSkinButton5Click(Sender: TObject);
    procedure bsSkinButton6Click(Sender: TObject);
    procedure bsSkinButton4Click(Sender: TObject);
    procedure bsSkinButton14Click(Sender: TObject);
    procedure DBGridEh2GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure DBGridEh1TitleClick(Column: TColumnEh);
    procedure ADOQuery2AfterOpen(DataSet: TDataSet);
    procedure SetADOText2(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure SetADOText1(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure alloydesign_Format;
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox1KeyPress(Sender: TObject; var Key: Char);
    procedure bsSkinButton72Click(Sender: TObject);
    procedure bsSkinButton8Click(Sender: TObject);
    procedure bsSkinButton1Click(Sender: TObject);
    procedure bsSkinButton2Click(Sender: TObject);
    procedure bsSkinButton7Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure bsSkinButton3Click(Sender: TObject);
    procedure bsSkinButton11Click(Sender: TObject);
    procedure bsSkinButton12Click(Sender: TObject); //合金组合表格格式化
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  eco: Teco;
  flag: boolean;

implementation
uses unit1;


{$R *.dfm}

procedure Teco.bsSkinButton5Click(Sender: TObject);
var
  Aqr1: TAdoQuery;

begin
  if trim(Edit1.Text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请输入名称！', (mtInformation), [mbOK], 0);
    edit1.SetFocus;
    exit;
  end;


  if bsSkinCheckListBox1.Items.IndexOf(trim(edit1.Text)) <> -1 then
  begin
    showmessage('该合金已存在，请修改名称');
    exit;
  end;


  Aqr1 := TAdoQuery.Create(nil);
  try
    Aqr1.Connection := form1.ADOConnection1;
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from alloy');
    Aqr1.Open;


    Aqr1.Append;
    Aqr1.Fields[1].AsString := trim(edit1.Text);
    Aqr1.Post;


    bsSkinCheckListBox1.Items.Add(trim(edit1.Text));

  finally
    Aqr1.Close;
    Aqr1.Free;

  end;



end;


procedure Teco.bsSkinButton6Click(Sender: TObject);
var
  Aqr1: TAdoQuery;
  i: integer;
begin

  if Form1.checkedcount(bsSkinCheckListBox1) = 0 then exit;

  if Form1.bsSkinMessage1.MessageDlg('确定删除所选' + inttostr(Form1.checkedcount(bsSkinCheckListBox1)) + '项合金？', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    exit;

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection1;
  try

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from alloy');
    Aqr1.Open;


    for i := bsSkinCheckListBox1.Items.Count - 1 downto 0 do
      if bsSkinCheckListBox1.Checked[i] = true then
      begin
        if Aqr1.Locate('alloy', bsSkinCheckListBox1.Items.Strings[i], []) then
          Aqr1.Delete;
        bsSkinCheckListBox1.Items.Delete(i);
      end;







  finally

    Aqr1.Close;
    Aqr1.Free;

  end;



end;

procedure Teco.bsSkinButton4Click(Sender: TObject);
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

  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection1;
  try
    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from alloy where alloy like ' + #39 + '%' + trim(Edit1.Text) + '%' + #39);
    Aqr1.Open;
    if Aqr1.RecordCount = 0 then
    begin
      form1.bsSkinMessage1.MessageDlg('未查询到相关产品！', (mtInformation), [mbOK], 0);
      edit1.SetFocus;
      exit;
    end;

    bsSkinCheckListBox1.Refresh;

    for i := bsSkinCheckListBox1.Items.Count - 1 downto 0 do
      bsSkinCheckListBox1.Selected[i] := false;

    if Aqr1.RecordCount > 0 then
      for i := 1 to Aqr1.RecordCount do
      begin
        Aqr1.RecNo := i;
        bsSkinCheckListBox1.Selected[bsSkinCheckListBox1.Items.IndexOf(Aqr1.FieldByName('alloy').AsString)] := true;
      end;
  finally
    Aqr1.Close;
    Aqr1.Free;
  end;




end;

procedure Teco.bsSkinButton14Click(Sender: TObject); // →
var
  i: integer;
begin
  if trim(ComboBox1.Text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请先输入方案名称！', (mtInformation), [mbOK], 0);
    ComboBox1.SetFocus;
    exit;
  end;

  if form1.checkedcount(bsSkinCheckListBox1) = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('未选择合金！', (mtInformation), [mbOK], 0);
    exit;
  end;


  if ADOQuery2.Active then ADOQuery2.Close;
  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('select ID, alloyname, alloynum, cases from alloydesign where cases= ' + #39 + trim(ComboBox1.Text) + #39);
  ADOQuery2.Open;

  for i := bsSkinCheckListBox1.Items.Count - 1 downto 0 do
    if bsSkinCheckListBox1.Checked[i] = true then
      if ADOQuery2.Locate('alloyname', bsSkinCheckListBox1.Items.Strings[i], []) = false then
      begin
        ADOQuery2.Append;
        ADOQuery2.FieldByName('cases').AsString := trim(ComboBox1.Text);
        ADOQuery2.FieldByName('alloyname').AsString := bsSkinCheckListBox1.Items.Strings[i];
        ADOQuery2.Post;
        ADOQuery2.Refresh;
      end;

  alloydesign_Format;

  if ComboBox1.Items.IndexOf(trim(ComboBox1.Text)) = -1 then
    ComboBox1.Items.Add(trim(ComboBox1.Text));


end;


procedure Teco.SetADOText1(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  if ADOQuery1.Active then
  begin
    Text := inttostr(ADOQuery1.RecNo);
    if ADOQuery1.RecordCount = 0 then
      Text := '0';

  end;
end;


procedure Teco.SetADOText2(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  if ADOQuery2.Active then
  begin
    Text := inttostr(ADOQuery2.RecNo);
    if ADOQuery2.RecordCount = 0 then
      Text := '0';

  end;
end;




procedure Teco.DBGridEh2GetCellParams(Sender: TObject; Column: TColumnEh;
  AFont: TFont; var Background: TColor; State: TGridDrawState);
begin
  if not (gdSelected in State) then
    if (Sender as TDBGridEh).SumList.RecNo mod 2 = 0 then
      Background := clInfobk;
end;

procedure Teco.DBGridEh1TitleClick(Column: TColumnEh);
begin
  if flag then
    ADOQuery1.Sort := Column.FieldName + ' ASC'
  else
    ADOQuery1.Sort := Column.FieldName + ' DESC';

  flag := not flag;
end;

procedure Teco.ADOQuery2AfterOpen(DataSet: TDataSet);
begin

  ADOQuery2.FieldByName('ID').OnGetText := SetADOText2;

end;


procedure Teco.alloydesign_Format; //合金组合表格格式化
var
  i: integer;
  mysql1: string;
begin
  for i := 0 to DBGridEh2.Columns.Count - 1 do
  begin
    DBGridEh2.Columns[i].Width := 90;
    DBGridEh2.Columns[i].Title.Alignment := taCenter;
    DBGridEh2.Columns[i].Alignment := taCenter;
  end;
  DBGridEh2.Columns[1].Alignment := taLeftJustify;
  DBGridEh2.Columns[0].Width := 40;
  DBGridEh2.Columns[2].Width := 120;
  DBGridEh2.Columns[0].Title.caption := '序号';
  DBGridEh2.Columns[1].Title.caption := '合金名称';
  DBGridEh2.Columns[2].Title.caption := '吨钢加入量（kg）';
  DBGridEh2.Columns[3].Visible := false;
end;




procedure Teco.ComboBox1Change(Sender: TObject);
begin
  if ADOQuery2.Active then ADOQuery2.Close;
  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('select ID, alloyname, alloynum,cases from alloydesign where cases= ' + #39 + trim(ComboBox1.Text) + #39);
  ADOQuery2.Open;
  alloydesign_Format;


end;

procedure Teco.ComboBox1KeyPress(Sender: TObject; var Key: Char);
begin
  if ComboBox1.Items.IndexOf(ComboBox1.Text) <> -1 then
  begin
    if ADOQuery2.Active then ADOQuery2.Close;
    ADOQuery2.SQL.Clear;
    ADOQuery2.SQL.Add('select ID, alloyname, alloynum,cases from alloydesign where cases= ' + #39 + trim(ComboBox1.Text) + #39);
    ADOQuery2.Open;
    alloydesign_Format;
  end;

end;


procedure Teco.bsSkinButton72Click(Sender: TObject);
var
  i: integer;
begin
  if DBGridEh2.SelectedRows.Count = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('请选择需要删除的合金！', (mtInformation), [mbOK], 0);
    exit;
  end;

  if form1.bsSkinMessage1.MessageDlg('确定删除所选记录？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := DBGridEh2.SelectedRows.Count - 1 downto 0 do
    begin
      DBGridEh2.DataSource.DataSet.GotoBookmark(pointer(DBGridEh2.SelectedRows[i]));
      DBGridEh2.DataSource.DataSet.Delete;
    end;
  end;

  ADOQuery2.UpdateBatch(arall);

end;

procedure Teco.bsSkinButton8Click(Sender: TObject);
var
  i: integer;
  Aqr1: TAdoQuery;
begin

  if form1.bsSkinMessage1.MessageDlg('确定删除本方案？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    for i := DBGridEh2.DataSource.DataSet.RecordCount downto 1 do
    begin
      DBGridEh2.DataSource.DataSet.RecNo := i;
      DBGridEh2.DataSource.DataSet.Delete;
    end;


    Aqr1 := TAdoQuery.Create(nil);
    Aqr1.Connection := form1.ADOConnection1;


    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('delete * from casesLCA where cases=' + #39 + trim(ComboBox1.Text) + #39);
    Aqr1.ExecSQL;

    Aqr1.Close;
    Aqr1.Free;


    ADOQuery2.UpdateBatch(arall);
    ADOQuery2.Refresh;
    ComboBox1.Items.Delete(ComboBox1.Items.IndexOf(ComboBox1.Text));
    ComboBox1.Text := '';

  end;



end;

procedure Teco.bsSkinButton1Click(Sender: TObject);
var
  Aqr1: TAdoQuery;
  i, j, k: integer;
  mysql: string;
begin
  if ADOQuery2.Active = false then exit;

  if trim(ComboBox1.Text) = '' then
  begin
    form1.bsSkinMessage1.MessageDlg('请先输入方案名称！', (mtInformation), [mbOK], 0);
    ComboBox1.SetFocus;
    exit;
  end;

 if  ADOQuery2.RecordCount  = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('no data ', (mtInformation), [mbOK], 0);
    exit;
  end;

  ADOQuery2.UpdateBatch(arall);

  StringGrid1.RowCount := 1;
  StringGrid1.ColCount := 1;


  Aqr1 := TAdoQuery.Create(nil);
  Aqr1.Connection := form1.ADOConnection1;
  try

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('delete * from casesLCA where cases=' + #39 + trim(ComboBox1.Text) + #39);
    Aqr1.ExecSQL;


    StringGrid1.Cells[0, 0] := '序号';
    StringGrid1.Cells[1, 0] := 'LCI因子';
    StringGrid1.Cells[2, 0] := '单位';


    mysql := 'select distinct lcifactor, unit  from up  order by  lcifactor';
  //  if CheckBox1.Checked = true then
  //    mysql := 'select distinct lcifactor, unit  from up where see_level=' + #39 + '1' + #39 + ' order by  lcifactor';

    if ADOQuery1.Active then ADOQuery1.Close;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add(mysql);
    ADOQuery1.Open;

    StringGrid1.RowCount := ADOQuery1.RecordCount + 1; //最后一列求和
    StringGrid1.ColCount := 4 + ADOQuery2.RecordCount;

    Matrix1.NrOfColumns := ADOQuery2.RecordCount + 1;
    Matrix1.NrOfRows := ADOQuery1.RecordCount; // 53
    for i := 1 to ADOQuery2.RecordCount do
    begin
      ADOQuery2.RecNo := i;
      mysql := 'select *  from up where product= ' + #39 + ADOQuery2.FieldByName('alloyname').AsString + #39 + '  order by  lcifactor';
      if ADOQuery1.Active then ADOQuery1.Close;
      ADOQuery1.SQL.Clear;
      ADOQuery1.SQL.Add(mysql);
      ADOQuery1.Open;
      for j := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := j;
        Matrix1.Elem[i, j] := ADOQuery1.FieldByName('sum').AsFloat * ADOQuery2.FieldByName('alloynum').AsFloat / 1000;
      end;
    end;


    for j := 1 to Matrix1.NrOfRows do
      Matrix1.Elem[Matrix1.NrOfColumns, j] := 0;

    for i := 1 to Matrix1.NrOfColumns - 1 do
      for j := 1 to Matrix1.NrOfRows do
        Matrix1.Elem[Matrix1.NrOfColumns, j] := Matrix1.Elem[Matrix1.NrOfColumns, j] + Matrix1.Elem[i, j]; //最后一列求和



    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from casesLCA');
    Aqr1.Open;

    mysql := 'select *  from lcifactor  order by inf_name';
    if ADOQuery1.Active then ADOQuery1.Close;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add(mysql);
    ADOQuery1.Open;

    for i := 1 to ADOQuery1.RecordCount do
    begin
      ADOQuery1.RecNo := i;
      Aqr1.Append;
      Aqr1.FieldByName('cases').AsString := trim(ComboBox1.Text);
      Aqr1.FieldByName('LCIfactor').AsString := ADOQuery1.FieldByName('inf_name').AsString;
      Aqr1.FieldByName('units').AsString := ADOQuery1.FieldByName('units').AsString;
      Aqr1.FieldByName('LCI').AsFloat := Matrix1.Elem[Matrix1.NrOfColumns, i];
      Aqr1.FieldByName('see_level').AsString :=ADOQuery1.FieldByName('see_level').AsString  ;
      Aqr1.Post;
    end;

    for i := 1 to ADOQuery1.RecordCount do
    begin
      ADOQuery1.RecNo := i;
      StringGrid1.Cells[0, i] := inttostr(i);
      StringGrid1.Cells[1, i] := ADOQuery1.Fields[0].AsString;
      StringGrid1.Cells[2, i] := ADOQuery1.Fields[1].AsString;
    end;


    StringGrid1.ColWidths[0] := 30;
    StringGrid1.ColWidths[1] := 150;
    StringGrid1.ColWidths[2] := 50;






    i := 3;

    for j := 1 to ADOQuery2.RecordCount do
    begin
      ADOQuery2.RecNo := j;

      mysql := 'select  lcifactor,sum  from up where product=' + #39 + ADOQuery2.FieldByName('alloyname').AsString + #39 + ' order by  lcifactor';
  //  if CheckBox1.Checked = true then
 //       mysql := 'select distinct lcifactor,sum  from up where see_level=' + #39 + '1' + #39 + ' and product=' + #39 + ADOQuery2.FieldByName('alloyname').AsString + #39 + ' order by  lcifactor';
      if ADOQuery1.Active then ADOQuery1.Close;
      ADOQuery1.SQL.Clear;
      ADOQuery1.SQL.Add(mysql);
      ADOQuery1.Open;
      StringGrid1.Cells[i, 0] := ADOQuery2.FieldByName('alloyname').AsString;
      for k := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := k;
        StringGrid1.Cells[i, k] := FormatFloat('#,##0.000', ADOQuery1.FieldbyName('sum').AsFloat);

      end;

      i := i + 1;

    end;

    StringGrid1.Cells[3 + ADOQuery2.RecordCount, 0] := '方案-' + trim(ComboBox1.Text);

    if Aqr1.Active then Aqr1.Close;
    Aqr1.SQL.Clear;
    Aqr1.SQL.Add('select * from casesLCA where cases=' + #39 + trim(ComboBox1.Text) + #39 + ' order by  lcifactor ');
    Aqr1.Open;
    for k := 1 to Aqr1.RecordCount do
    begin
      Aqr1.RecNo := k;
      StringGrid1.Cells[3 + ADOQuery2.RecordCount, k] := FormatFloat('#,##0.000', Aqr1.FieldbyName('LCI').AsFloat);

    end;


    StringGrid1.FixedRows := 1;



  finally
    Aqr1.Close;
    Aqr1.Free;
  end;



end;

procedure Teco.bsSkinButton2Click(Sender: TObject);
var
  mysql: string;
  i, j, k: integer;
begin
  if form1.checkedcount(bsSkinCheckListBox1) = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('未选择合金！', (mtInformation), [mbOK], 0);
    exit;
  end;

 



  StringGrid1.Cells[0, 0] := '序号';
  StringGrid1.Cells[1, 0] := 'LCI因子';
  StringGrid1.Cells[2, 0] := '单位';

  mysql := 'select distinct lcifactor, unit  from up  order by  lcifactor';
  if CheckBox1.Checked = true then
    mysql := 'select distinct lcifactor, unit  from up where see_level=' + #39 + '1' + #39 + ' order by  lcifactor';

  if ADOQuery1.Active then ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(mysql);
  ADOQuery1.Open;
  StringGrid1.RowCount := ADOQuery1.RecordCount + 1;
  StringGrid1.ColCount := 3 + form1.checkedcount(bsSkinCheckListBox1);


  for i := 1 to ADOQuery1.RecordCount do
  begin
    ADOQuery1.RecNo := i;
    StringGrid1.Cells[0, i] := inttostr(i);
    StringGrid1.Cells[1, i] := ADOQuery1.Fields[0].AsString;
    StringGrid1.Cells[2, i] := ADOQuery1.Fields[1].AsString;
  end;


  StringGrid1.ColWidths[0] := 30;
  StringGrid1.ColWidths[1] := 150;


  i := 3;
  for j := 0 to bsSkinCheckListBox1.Items.Count - 1 do
    if bsSkinCheckListBox1.Checked[j] = true then
    begin
      mysql := 'select  lcifactor,sum  from up where product=' + #39 + bsSkinCheckListBox1.Items.Strings[j] + #39 + ' order by  lcifactor';
      if CheckBox1.Checked = true then
        mysql := 'select distinct lcifactor,sum  from up where see_level=' + #39 + '1' + #39 + ' and product=' + #39 + bsSkinCheckListBox1.Items.Strings[j] + #39 + ' order by  lcifactor';
      if ADOQuery1.Active then ADOQuery1.Close;
      ADOQuery1.SQL.Clear;
      ADOQuery1.SQL.Add(mysql);
      ADOQuery1.Open;
      StringGrid1.Cells[i, 0] := bsSkinCheckListBox1.Items.Strings[j];
      for k := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := k;
        StringGrid1.Cells[i, k] := formatfloat('#,##0.000', ADOQuery1.Fieldbyname('sum').asfloat);
      end;

      i := i + 1;





    end;


 {
  if ADOQuery1.Active then ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(mysql);
  ADOQuery1.Open; }



end;

procedure Teco.bsSkinButton7Click(Sender: TObject);
var
  sl, t: TStrings;
  i: Integer;
begin
  sl := TStringList.Create;
  try
    for i := 0 to self.StringGrid1.RowCount - 1 do
    begin
      t := self.StringGrid1.Rows[i];
      t.Delimiter := #9;
      sl.Append(t.DelimitedText)
    end;
    form1.SaveDialog1.Filter := '*.txt';
    form1.SaveDialog1.Title := '保存';
    if form1.SaveDialog1.Execute then
      sl.SaveToFile(form1.SaveDialog1.FileName + '.txt');
  finally
    sl.Free;
  end;
end;



procedure Teco.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked = true then
  begin

    ComboBox2.ItemIndex := 0;
 //    ComboBox2.Text:='';
  end;

end;

procedure Teco.ComboBox2Change(Sender: TObject);
begin
  if ComboBox2.ItemIndex > 0 then
    CheckBox1.Checked := false;
end;

procedure Teco.bsSkinButton3Click(Sender: TObject);
var
  i, j,k: integer;
   mysql:string;
begin
  j := 0;
  for i := 0 to bsSkinCheckComboBox1.Items.Count - 1 do
    if bsSkinCheckComboBox1.Checked[i] = true then
      j := j + 1;

  if j = 0 then
  begin
    form1.bsSkinMessage1.MessageDlg('未选择要比较的方案！', (mtInformation), [mbOK], 0);
    exit;
  end;
        StringGrid1.ColCount := 3 + j;

  mysql := 'select distinct lcifactor, units  from casesLCA  order by  lcifactor';
  if CheckBox1.Checked = true then
    mysql := 'select distinct lcifactor, units  from  casesLCA  where see_level=' + #39 + '1' + #39 + ' order by  lcifactor';

  if ADOQuery1.Active then ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(mysql);
  ADOQuery1.Open;
  StringGrid1.RowCount := ADOQuery1.RecordCount + 1;


  StringGrid1.Cells[0, 0] := '序号';
  StringGrid1.Cells[1, 0] := 'LCI因子';
  StringGrid1.Cells[2, 0] := '单位';

  for i := 1 to ADOQuery1.RecordCount do
  begin
    ADOQuery1.RecNo := i;
    StringGrid1.Cells[0, i] := inttostr(i);
    StringGrid1.Cells[1, i] := ADOQuery1.Fields[0].AsString;
    StringGrid1.Cells[2, i] := ADOQuery1.Fields[1].AsString;
  end;


  StringGrid1.ColWidths[0] := 30;
  StringGrid1.ColWidths[1] := 150;


  i := 3;
  for j := 0 to bsSkinCheckComboBox1.Items.Count - 1 do
    if bsSkinCheckComboBox1.Checked[j] = true then
    begin
      mysql := 'select  lcifactor,LCI  from casesLCA where cases=' + #39 + bsSkinCheckComboBox1.Items.Strings[j] + #39 + ' order by  lcifactor';
      if CheckBox1.Checked = true then
        mysql := 'select distinct lcifactor,LCI  from casesLCA where see_level=' + #39 + '1' + #39 + ' and cases=' + #39 + bsSkinCheckComboBox1.Items.Strings[j] + #39 + ' order by  lcifactor';
      if ADOQuery1.Active then ADOQuery1.Close;
      ADOQuery1.SQL.Clear;
      ADOQuery1.SQL.Add(mysql);
      ADOQuery1.Open;
      StringGrid1.Cells[i, 0] := bsSkinCheckComboBox1.Items.Strings[j];
      for k := 1 to ADOQuery1.RecordCount do
      begin
        ADOQuery1.RecNo := k;
        StringGrid1.Cells[i, k] := formatfloat('#,##0.000', ADOQuery1.Fieldbyname('LCI').asfloat);
      end;

      i := i + 1;

  
    end;





end;

procedure Teco.bsSkinButton11Click(Sender: TObject);
begin
panel1.Visible:=true;
panel1.BringToFront;
panel1.Align:= alClient;
end;

procedure Teco.bsSkinButton12Click(Sender: TObject);
begin
  panel1.Visible:=false;
end;

end.

