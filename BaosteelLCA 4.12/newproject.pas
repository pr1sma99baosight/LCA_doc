unit newproject;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, fileCtrl,
  Dialogs, Menus, StdCtrls, bsSkinCtrls, ADODB, TFlatEditUnit, ExtCtrls, Comobj, ShellAPI;

 {

  Buttons,    DB,
  TeeProcs, TeEngine, Chart, newproject;    }


type
  TnewprojectForm = class(TForm)
    bsSkinPanel1: TbsSkinPanel;
    Label1: TLabel;
    Label2: TLabel;
    FlatEdit1: TFlatEdit;
    bsSkinButton9: TbsSkinButton;
    bsSkinButton7: TbsSkinButton;
    bsSkinButton8: TbsSkinButton;
    bsSkinPanel4: TbsSkinPanel;
    ListBox1: TListBox;
    ListBox2: TListBox;
    bsSkinButton1: TbsSkinButton;
    procedure FlatEdit1Click(Sender: TObject);
    procedure ListBox1Renew;
    function ListDirs(Path: string; List: TStringList): Integer;
    procedure bsSkinButton9Click(Sender: TObject);
    procedure bsSkinButton7Click(Sender: TObject);
    procedure bsSkinButton8Click(Sender: TObject);
    procedure bsSkinButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  newprojectForm: TnewprojectForm;

implementation
uses unit1;
{$R *.dfm}

procedure TnewprojectForm.ListBox1Renew; //
var
  s: TStringList;
begin
  ListBox1.Clear;
  s := TStringList.Create;
  ListDirs(ExtractFilePath(Application.ExeName) + '\projects\', s);
  ListBox1.Items.AddStrings(s);
  s.Free;

end;

function TnewprojectForm.ListDirs(Path: string; List: TStringList): Integer;
//��ȡĳһĿ¼�������ļ�������
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

procedure TnewprojectForm.FlatEdit1Click(Sender: TObject);
begin
  if FlatEdit1.Text = '����Ŀ����' then
    FlatEdit1.Text := '';
end;

procedure TnewprojectForm.bsSkinButton9Click(Sender: TObject);
var
 // Aqr1: TAdoQuery;
  i: integer;
  s: string;
  route: string;
 // myrec: TSHFILEOPSTRUCT;
 // Aqr1: TAdoQuery;
begin
  s := trim(FlatEdit1.Text);
  if trim(FlatEdit1.Text) = '' then exit;
 // showmessage(inttostr(length(trim(FlatEdit1.Text))));
  for i := 1 to length(trim(FlatEdit1.Text)) do
    if s[i] in ['\', '/', ':', '*', '?', '<', '>', '|'] then
    begin
      form1.bsSkinMessage1.MessageDlg('�ļ����а����Ƿ��ַ���' + #39 + s[i] + #39, (mtInformation), [mbOK], 0);
      exit;
    end;


  route := ExtractFilePath(Application.ExeName) + '\projects\' + trim(FlatEdit1.Text);
  form1.Label1.Caption := trim(FlatEdit1.Text);

  if DirectoryExists(route) = true then
  begin
    form1.bsSkinMessage1.MessageDlg('���� ' + #39 + trim(FlatEdit1.Text) + #39 + ' �Ѵ���! ', (mtInformation), [mbOK], 0);
    exit;
  end;

  with Form1 do
  begin
    label11.Caption := '�½�';
    label3.Caption := '<' + FlatEdit1.Text + '> ' + '��Ŀ������Ϣ';
    FlatEdit3.Text :=  datetimetostr(now);

  end;


  Createdir(route); //�����ļ���
  CopyFile(PChar(ExtractFilePath(Application.ExeName) + '\base.mdb'), PChar(route + '\' + trim(FlatEdit1.Text) + '.mdb'), True);
  CopyFile(PChar(ExtractFilePath(Application.ExeName) + '\tree.txt'), PChar(route + '\tree.txt'), True);


  form1.ADOConnection2.Close;
  form1.ADOConnection2.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + route + '\' + trim(FlatEdit1.Text) + '.mdb ; Persist Security Info=False';
  form1.ADOConnection2.Open;



  Form1.TreeView1Renew;

  for i := 0 to Form1.TreeView1.Items.Count - 1 do
  begin
    if Form1.TreeView1.Items[i].Text = trim(FlatEdit1.Text) then
    begin
      Form1.TreeView1.Items[i].Selected := true;
      Form1.TreeView1Click(Sender);
      break;
    end;
  end;

   Form1.FlatEdit3.Text :=  datetimetostr(now);



  Form1.TreeView1.SetFocus;
  newprojectForm.Close;
end;



procedure TnewprojectForm.bsSkinButton7Click(Sender: TObject);
var
  OpStruc: TSHFileOpStruct;
  frombuf, tobuf: array[0..128] of Char;
  route, s: string;
 // Aqr1: TAdoQuery;
begin
  if ListBox1.ItemIndex = -1 then exit;

  s := trim(ListBox1.Items[ListBox1.ItemIndex] + '�ĸ���');
  if InputQuery('�������������', '', s) = false then exit;

  if ListBox1.Items.IndexOf(s) <> -1 then
  begin
    form1.bsSkinMessage1.MessageDlg('�������Ѵ��ڣ����������룡', (mtInformation), [mbOK], 0);
    exit;
  end;

  route := ExtractFilePath(Application.ExeName) + 'projects\';

  FillChar(frombuf, Sizeof(frombuf), 0);
  FillChar(tobuf, Sizeof(tobuf), 0);
  StrPCopy(frombuf, route + ListBox1.Items[ListBox1.ItemIndex]);
  StrPCopy(tobuf, route + s);
  with OpStruc do begin
    Wnd := Handle;
    wFunc := FO_COPY;
    pFrom := @frombuf;
    pTo := @tobuf;
    fFlags := FOF_NOCONFIRMATION or FOF_RENAMEONCOLLISION;
    fAnyOperationsAborted := False;
    hNameMappings := nil;
    lpszProgressTitle := nil;
  end;
  ShFileOperation(OpStruc);


  RenameFile(route + s + '\' + ListBox1.Items[ListBox1.ItemIndex] + '.mdb', route + s + '\' + s + '.mdb');
  RenameFile(route + s + '\' + ListBox1.Items[ListBox1.ItemIndex] + '.flc', route + s + '\' + s + '.flc');
  RenameFile(route + s + '\' + ListBox1.Items[ListBox1.ItemIndex] + '.bmp', route + s + '\' + s + '.bmp');
  DeleteFile(route + s + '\' + ListBox1.Items[ListBox1.ItemIndex] + '.ldb');

  ListBox1.Items.Add(s);

  Form1.TreeView1Renew;

end;

procedure TnewprojectForm.bsSkinButton8Click(Sender: TObject);
var
  myrec: TSHFILEOPSTRUCT;
  route: string;
 // Aqr1: TAdoQuery;

begin
  if ListBox1.ItemIndex = -1 then
  begin
    form1.bsSkinMessage1.MessageDlg('��ѡ����Ҫɾ������Ŀ��', (mtInformation), [mbOK], 0);
    exit;
  end;

  if form1.bsSkinMessage1.MessageDlg('ȷ��ɾ��LCA��Ŀ�ļ�  "' + ListBox1.Items[ListBox1.ItemIndex] + '"  ?', mtConfirmation,
    [mbYes, mbNo], 0) = mrYes then
  begin
    form1.ADOConnection2.Close;
    route := ExtractFilePath(Application.ExeName) + '\projects\' + listbox1.Items[listbox1.ItemIndex];
    with myrec do begin
      Wnd := Handle;
      wFunc := FO_DELETE;
      pFrom := pChar(route);
      pTo := '';
      fFlags := FOF_NOCONFIRMATION or FOF_FILESONLY;
      fAnyOperationsAborted := False;
      hNameMappings := nil;
      lpszProgressTitle := nil;
    end;
    SHFileOperation(myrec);



  {  ListBox2.Items.Clear;
    form1.ADOConnection3_data.GetTableNames(ListBox2.Items);
    if ListBox2.Items.IndexOf(listbox1.Items[listbox1.ItemIndex]) <> -1 then
    begin
      Aqr1 := TAdoQuery.Create(nil);
      Aqr1.Connection := form1.ADOConnection3_data;
      try
        s := 'drop table ' + listbox1.Items[listbox1.ItemIndex];
        if Aqr1.Active then Aqr1.Close;
        Aqr1.SQL.Clear;
        Aqr1.SQL.Add(s);
        Aqr1.ExecSQL;
      finally
        Aqr1.Close;
        Aqr1.Free;
      end;
    end;     }

    Listbox1.Items.Delete(Listbox1.ItemIndex);
  end;
  Form1.TreeView1Renew;
    if Form1.TreeView1.Items.Count>0 then
  Form1.TreeView1.Items[0].Selected:=true;
end;

procedure TnewprojectForm.bsSkinButton1Click(Sender: TObject);
var
  Dir, s: string;
begin
  if ListBox1.ItemIndex = -1 then exit;

  s := ExtractFilePath(Application.ExeName) + 'projects\';

  if SelectDirectory('ѡ����Ŀ�ļ�����·��', '', Dir) then
  begin
    form1.CopyDir(s + ListBox1.Items[ListBox1.ItemIndex], dir);
    form1.bsSkinMessage1.MessageDlg('������ɣ�', (mtInformation), [mbOK], 0);
  end;

end;

end.

