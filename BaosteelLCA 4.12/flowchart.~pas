unit flowchart;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, BusinessSkinForm, StdCtrls, ComCtrls, ExtCtrls, dxflchrt, Mask,
  bsSkinBoxCtrls, bsSkinCtrls, DB, ADODB, ToolWin, ImgList, dxFcEdit,
  Buttons, Menus, matrix, ShellAPI, GridsEh, DBGridEh, bsSkinTabs,
  TFlatEditUnit, allocation;

type
  TChart = class(TdxFlowChart) end;

  TUndoManager = class
    FStream: TMemoryStream;
    FChart: TdxFlowChart;
    FUndoSteps: Integer;
    FCanUndo: Boolean;
    FStep: Integer;
    FLength: array[1..200] of integer;
    procedure SetChart(Value: TdxFlowChart);
    procedure SetUndoSteps(Value: integer);
  public
    constructor Create;
    destructor Destroy; override;
    property UndoSteps: Integer read FUndoSteps write SetUndoSteps;
    property Chart: TdxFlowChart read FChart write SetChart;
    property CanUndo: Boolean read FCanUndo;
    procedure Store;
    procedure Undo;
  end;



  TForm2 = class(TForm)
    Panel1: TPanel;
    Panel7: TPanel;
    ListBox1: TListBox;
    Panel2: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    TreeView1: TTreeView;
    bsSkinButton1: TbsSkinButton;
    bsSkinButton2: TbsSkinButton;
    bsSkinButton3: TbsSkinButton;
    bsSkinButton4: TbsSkinButton;
    bsSkinEdit1: TbsSkinEdit;
    bsSkinButton7: TbsSkinButton;
    ADOQuery1: TADOQuery;
    ColorDialog: TColorDialog;
    FontDialog: TFontDialog;
    ShapePopupMenu: TPopupMenu;
    iNone: TMenuItem;
    iRectangle: TMenuItem;
    iEllipse: TMenuItem;
    iRoundRect: TMenuItem;
    iDiamond: TMenuItem;
    iNorthTriangle: TMenuItem;
    itSouthTriangle: TMenuItem;
    itEastTriangle: TMenuItem;
    itWestTriangle: TMenuItem;
    itHexagon: TMenuItem;
    StylePopupMenu: TPopupMenu;
    iStraight: TMenuItem;
    iCurved: TMenuItem;
    iRectHorizontal: TMenuItem;
    iRectVertical: TMenuItem;
    SmallImages: TImageList;
    SaveDialog: TSaveDialog;
    Matrix1: TMatrix;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    dxFlowChart1: TdxFlowChart;
    bsSkinPanel1: TbsSkinPanel;
    SpeedButton8: TSpeedButton;
    Bevel2: TBevel;
    Bevel4: TBevel;
    Bevel1: TBevel;
    Label8: TLabel;
    bsSkinPanel2: TbsSkinPanel;
    SpeedButton1: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton7: TSpeedButton;
    sbUndo: TSpeedButton;
    bsSkinPanel3: TbsSkinPanel;
    btnCreateConnect: TSpeedButton;
    sbShape: TSpeedButton;
    sbStyle: TSpeedButton;
    sbObjectFont: TSpeedButton;
    Bevel3: TBevel;
    pColor: TPanel;
    fillingcolor: TPanel;
    linecolor: TPanel;
    bsSkinPanel4: TbsSkinPanel;
    SpeedButton2: TSpeedButton;
    SpeedButton4: TSpeedButton;
    sbFit: TSpeedButton;
    bsSkinButton5: TbsSkinButton;
    bsSkinButton9: TbsSkinButton;
    bsSkinButton11: TbsSkinButton;
    bsSkinButton6: TbsSkinButton;
    DataSource1: TDataSource;
    procedure TreeView1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dxFlowChart1DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure dxFlowChart1Change(Sender: TdxCustomFlowChart;
      Item: TdxFcItem);
    procedure dxFlowChart1DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure dxFlowChart1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dxFlowChart1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure dxFlowChart1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);


    function NormalizeRect(ARect: TRect): TRect;
    procedure FormCreate(Sender: TObject);
    procedure btnCreateConnectClick(Sender: TObject);
    procedure sbUndoClick(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure sbShapeClick(Sender: TObject);
    procedure sbStyleClick(Sender: TObject);
    procedure sbObjectFontClick(Sender: TObject);
    procedure pColorClick(Sender: TObject);
    procedure fillingcolorClick(Sender: TObject);
    procedure linecolorClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure sbFitClick(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure bsSkinButton5Click(Sender: TObject);
    procedure pColorDblClick(Sender: TObject);
    procedure fillingcolorDblClick(Sender: TObject);
    procedure linecolorDblClick(Sender: TObject);
    procedure bsSkinButton29Click(Sender: TObject);
    procedure DBGridEh1GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure bsSkinButton1Click(Sender: TObject); //物料属性 数据格式化
    function NameUsed(name: string; treeview0: Ttreeview): boolean;
    procedure bsSkinButton2Click(Sender: TObject);
    procedure bsSkinButton3Click(Sender: TObject);
    procedure bsSkinButton4Click(Sender: TObject);
    procedure DBGridEh1Enter(Sender: TObject); //判断是否重名
    procedure PaintCtrlToBmp(Ctrl: TWinControl; Bmp: TBitmap);
    procedure bsSkinButton6Click(Sender: TObject);
    procedure bsSkinButton11Click(Sender: TObject);
    procedure dxFlowChart1Deletion(Sender: TdxCustomFlowChart;
      Item: TdxFcItem);
    procedure iRectHorizontalClick(Sender: TObject);
    procedure iNoneClick(Sender: TObject);
    procedure iRectangleClick(Sender: TObject);
    procedure iEllipseClick(Sender: TObject);
    procedure iRoundRectClick(Sender: TObject);
    procedure iDiamondClick(Sender: TObject);
    procedure iNorthTriangleClick(Sender: TObject);
    procedure itSouthTriangleClick(Sender: TObject);
    procedure itEastTriangleClick(Sender: TObject);
    procedure itWestTriangleClick(Sender: TObject);
    procedure itHexagonClick(Sender: TObject);
    procedure iStraightClick(Sender: TObject);
    procedure iCurvedClick(Sender: TObject);
    procedure iRectVerticalClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure bsSkinButton9Click(Sender: TObject);
    procedure bsSkinButton7Click(Sender: TObject);

  private
    { Private declarations }

    DownPoint: TPoint;
    OldPoint: TPoint;
    MoveRect: TRect;
    moving: boolean;
    FUndo: TUndoManager;
    FStore: Boolean;
    Fchange: boolean;
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  Flowsave: boolean;
  flag: boolean;
  Saved: boolean;
implementation

uses Unit1, loadproject;

{$R *.dfm}



procedure TForm2.TreeView1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  a: TTreeNode;
begin
  a := TreeView1.GetNodeAt(X, Y);
  if a = nil then
    TreeView1.Selected := nil;

end;

procedure TForm2.dxFlowChart1DragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  l, t, w, h: integer;
  newobject: Tdxfcobject;
  xx: tpoint;
begin
  if listbox1.Items.IndexOf(treeview1.Selected.Text) = -1 then
    listbox1.Items.Add(treeview1.Selected.Text)
  else
  begin
    form1.bsSkinMessage1.MessageDlg(treeview1.Selected.Text + '已存在！', (mtInformation), [mbOK], 0);
    exit;
  end;
  getcursorpos(xx); //获取当前鼠标位置的坐标
  l := xx.X - 270; //在画图板上的坐标  x轴
  t := xx.Y - 50; //在画图板上的坐标  y轴
  h := 40;
  w := 75;
 // dxFlowChart1.Clear;
  FStore := False;
  newobject := dxFlowChart1.CreateObject(l, t, w, h, fcsrectangle);
  newobject.Text := trim(treeview1.Selected.Text);
  newobject.HorzTextPos := fchpcenter;
  newobject.VertTextPos := fcvpcenter;
  FStore := True;
  dxFlowChart1Change(dxFlowChart1, nil);



end;

//****************undo manager****************

constructor TUndoManager.Create;
begin
  FStream := TMemoryStream.Create;
  FStep := 0;
  FUndoSteps := 10;
  FCanUndo := False;
end;

destructor TUndoManager.Destroy;
begin
  FStream.Free;
  inherited;
end;

procedure TUndoManager.SetChart(Value: TdxFlowChart);
begin
  FChart := Value;
end;

procedure TUndoManager.SetUndoSteps(Value: integer);
begin
  FUndoSteps := Value;
  if FUndoSteps > 200 then FUndoSteps := 200;
end;

procedure TUndoManager.Store;
var
  Stream1, Stream2: TMemoryStream;
  i, StartPos: Integer;
  F: Boolean;
begin
  Stream1 := TMemoryStream.Create;
  Stream2 := TMemoryStream.Create;
  FChart.SaveToStream(Stream1);
  if FStep > 0 then begin
    StartPos := 0;
    for i := 1 to FStep - 1 do StartPos := StartPos + FLength[i];
    FStream.Position := StartPos;
    Stream2.Position := 0;
    Stream2.CopyFrom(FStream, FLength[FStep]);
  end;
  if Stream2.Size <> 0 then
    F := CompareMem(Stream1.Memory, Stream2.Memory, Stream1.Size)
  else F := False;
  if not (F and (Stream1.Size = Stream2.Size)) then begin
    if FStep >= FUndoSteps then begin
      Stream2.Clear;
      FStream.Position := FLength[1];
      Stream2.Position := 0;
      Stream2.CopyFrom(FStream, FStream.Size - FLength[1]);
      FStream.Clear;
      Stream2.Position := 0;
      FStream.Position := 0;
      FStream.CopyFrom(Stream2, 0);
      dec(FStep);
      for i := 1 to FStep do FLength[i] := FLength[i + 1];
    end;
    StartPos := 0;
    for i := 1 to FStep do StartPos := StartPos + FLength[i];
    FStream.Position := StartPos;
    FStream.CopyFrom(Stream1, 0);
    inc(FStep);
    FLength[FStep] := Stream1.Size;
  end;
  Stream1.Free;
  Stream2.Free;
  FCanUndo := FStep > 1;
end;

procedure TUndoManager.Undo;
var Stream: TMemoryStream;
  StartPos, i: Integer;
begin
  if not FCanUndo then exit;
  Stream := TMemoryStream.Create;
  StartPos := 0;
  for i := 1 to FStep - 2 do StartPos := StartPos + FLength[i];
  FStream.Position := StartPos;
  Stream.CopyFrom(FStream, FLength[FStep - 1]);
  Stream.Position := 0;
  FChart.LoadFromStream(Stream);
  dec(FStep);
  if FStep <= 1 then FCanUndo := False;
  StartPos := 0;
  for i := 1 to FStep do StartPos := StartPos + FLength[i];
  FStream.Size := StartPos;
  Stream.Free;
end;
//****************undo manager****************

procedure TForm2.dxFlowChart1Change(Sender: TdxCustomFlowChart;
  Item: TdxFcItem);
begin
  if Self.HandleAllocated then
  begin
    if FStore then
    begin
      FUndo.Store;
      sbUndo.Enabled := FUndo.CanUndo;
      Flowsave := false;
    end;
  end;
end;

procedure TForm2.dxFlowChart1DragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := Source is TtreeView;
end;

procedure TForm2.dxFlowChart1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  hTest: TdxFcHitTest;
  P: TPoint;
begin
  hTest := dxFlowChart1.GetHitTestAt(X, Y);
  if (hTest = [htNowhere]) and (Button = mbLeft) and (not (ssShift in Shift)) then
    dxFlowChart1.ClearSelection;

  if Button = mbLeft then
  begin
    DownPoint := Point(X, Y);
    OldPoint := DownPoint;
  end;

  if Button = mbLeft then
    if (dxFlowChart1.SelectedObjectCount = 0) and (dxFlowChart1.SelectedConnectionCount = 0) then
    begin
      SetCapture(dxFlowChart1.Handle);
      MoveRect.Left := X;
      MoveRect.Top := Y;
      MoveRect.BottomRight := MoveRect.TopLeft;
      TChart(dxFlowChart1).Canvas.DrawFocusRect(MoveRect);
      Moving := True;
    end;

  if Button = mbRight then
    if dxFlowChart1.SelCount > 0 then
    begin
      GetCursorPos(P);
      PopupMenu1.Popup(P.X, P.Y);
    end;

end;

procedure TForm2.dxFlowChart1MouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
var
  ARect: TRect;
begin
  if moving then
  begin
    ARect := NormalizeRect(MoveRect);
    TChart(dxFlowChart1).Canvas.DrawFocusRect(ARect);
    MoveRect.Right := X;
    MoveRect.Bottom := Y;
    ARect := NormalizeRect(MoveRect);
    TChart(dxFlowChart1).Canvas.DrawFocusRect(ARect);
  end;
 { else
    if ssShift in Shift then
      TChart(dxFlowChart1).Canvas.Pixels[X, Y] := clRed;  }

end;


function Tform2.NormalizeRect(ARect: TRect): TRect;
var
  tmp: Integer;
begin
  if ARect.Bottom < ARect.Top then
  begin
    tmp := ARect.Bottom;
    ARect.Bottom := ARect.Top;
    ARect.Top := tmp;
  end;
  if ARect.Right < ARect.Left then
  begin
    tmp := ARect.Right;
    ARect.Right := ARect.Left;
    ARect.Left := tmp;
  end;
  Result := ARect;
end;



procedure TForm2.dxFlowChart1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i: integer;
  SelectRect: TRect;
begin
  if moving then
    if Button = mbLeft then
    begin
      ReleaseCapture;
      moving := False;
      TChart(dxFlowChart1).Canvas.DrawFocusRect(MoveRect);
      dxFlowChart1.Refresh;

      if Button = mbLeft then
        DownPoint := Point(X, Y);
      SelectRect.Left := oldPoint.X;
      SelectRect.Right := DownPoint.X;
      SelectRect.Top := oldPoint.Y;
      SelectRect.Bottom := DownPoint.Y;
      for i := 0 to dxFlowChart1.ObjectCount - 1 do
        if dxFlowChart1.Objects[i].InRect(SelectRect) then
          dxFlowChart1.Objects[i].Selected := true;
      for i := 0 to dxFlowChart1.ConnectionCount - 1 do
        if dxFlowChart1.Connections[i].InRect(SelectRect) then
          dxFlowChart1.Connections[i].Selected := true;

    end;


end;

procedure TForm2.FormCreate(Sender: TObject);
begin

  moving := false;
  FChange := True;
  FUndo := TUndoManager.Create;
  FUndo.Chart := dxFlowChart1;
  FUndo.UndoSteps := 30;
  FStore := True;
  Flowsave := True;

end;

procedure TForm2.btnCreateConnectClick(Sender: TObject);
begin
  if (dxFlowChart1.SelectedObjectCount = 2) and (dxFlowChart1.SelectedConnectionCount = 0) then
  begin
    with dxFlowChart1 do
    begin
      dxFlowChart1.CreateConnection(SelectedObjects[0], SelectedObjects[1], 10, 2);
      FStore := False;
      Connections[ConnectionCount - 1].Selected := true;
      Connections[ConnectionCount - 1].ArrowDest.arrowtype := fcaArrow;
      Connections[ConnectionCount - 1].Style := TdxFclStyle(sbStyle.Tag - 1);
      Connections[ConnectionCount - 1].ArrowDest.Width := 1 * 5 + 5;
      Connections[ConnectionCount - 1].ArrowDest.Height := 1 * 5 + 5;
      Connections[ConnectionCount - 1].Color := clBlue;
      FStore := true;
    end;
  end;
end;

procedure TForm2.sbUndoClick(Sender: TObject);
var
  i: integer;
begin
  FStore := False;
  dxFlowChart1.Visible := false;
  FUndo.Undo;
  sbUndo.Enabled := FUndo.CanUndo;
  FStore := True; //showmessage('lll');
  for i := 0 to dxFlowChart1.ObjectCount - 1 do
    if listbox1.Items.IndexOf(dxFlowChart1.Objects[i].Text) = -1 then
      listbox1.Items.Add(dxFlowChart1.Objects[i].Text);


  dxFlowChart1.Visible := true;

end;

procedure TForm2.SpeedButton7Click(Sender: TObject);
begin
  dxFlowChart1.SetFocus;
  dxFlowChart1.SelectAll;
end;

procedure TForm2.SpeedButton5Click(Sender: TObject);
begin
  if dxFlowChart1.SelCount = 0 then
    exit;
  FStore := False;
//  if form1.bsSkinMessage1.MessageDlg('确定删除所选目标？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  dxFlowChart1.DeleteSelection;
  FStore := True;
  dxFlowChart1Change(dxFlowChart1, nil);
end;

procedure TForm2.SpeedButton1Click(Sender: TObject);
begin
  if dxFlowChart1.ObjectCount = 0 then
    if dxFlowChart1.ConnectionCount = 0 then
      exit;

  FStore := False;

  dxFlowChart1.Clear;
  listbox1.Clear;
  FStore := True;
  dxFlowChart1Change(dxFlowChart1, nil);
end;

procedure TForm2.sbShapeClick(Sender: TObject);
var
  pt: tpoint;
begin
  GetCursorPos(pt);
  ShapePopupMenu.Popup(pt.x, pt.y);

end;

procedure TForm2.sbStyleClick(Sender: TObject);
var
  pt: tpoint;
begin
  GetCursorPos(pt);
  StylePopupMenu.Popup(pt.x, pt.y);

end;

procedure TForm2.sbObjectFontClick(Sender: TObject);
var
  i: integer;
begin
  with TSpeedButton(Sender) do begin
    if dxFlowChart1.SelectedObjectCount > 0 then Font := dxFlowChart1.SelectedObjects[0].Font;
    FontDialog.Font.Assign(Font);
    if FontDialog.Execute then
    begin
      Font.Assign(FontDialog.Font);
      for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
        dxFlowChart1.SelectedObjects[i].Font := font;

    end;
  end;

end;

procedure TForm2.pColorClick(Sender: TObject);
var
  i: integer;
begin
  FStore := False;
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    with dxFlowChart1.SelectedObjects[i] do
      ShapeColor := pColor.Color;

  for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
    with dxFlowChart1.SelectedConnections[i] do
      Color := pColor.Color;
  FStore := True;
  dxFlowChart1Change(dxFlowChart1, nil);

end;

procedure TForm2.fillingcolorClick(Sender: TObject);
var
  i: integer;
begin
  FStore := False;
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    with dxFlowChart1.SelectedObjects[i] do
      BkColor := fillingcolor.Color;
  for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
    with dxFlowChart1.SelectedConnections[i] do
      Color := fillingcolor.Color;
  FStore := True;
  dxFlowChart1Change(dxFlowChart1, nil);

end;

procedure TForm2.linecolorClick(Sender: TObject);
var
  i: integer;
begin
  FStore := False;
  for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
    dxFlowChart1.SelectedConnections[i].Color := lineColor.Color;
  FStore := True;
  dxFlowChart1Change(dxFlowChart1, nil);

end;

procedure TForm2.SpeedButton2Click(Sender: TObject);
begin
  if dxFlowChart1.Zoom < 490 then dxFlowChart1.Zoom := dxFlowChart1.Zoom + 10 else SpeedButton2.Enabled := False;
  SpeedButton4.Enabled := True;
  sbFit.Enabled := (dxFlowChart1.Zoom <> 100) and (not sbFit.Down);
end;

procedure TForm2.SpeedButton4Click(Sender: TObject);
begin
  if dxFlowChart1.Zoom > 20 then dxFlowChart1.Zoom := dxFlowChart1.Zoom - 10 else SpeedButton4.Enabled := False;
  SpeedButton2.Enabled := True;
end;

procedure TForm2.sbFitClick(Sender: TObject);
begin
  SpeedButton2.Down := false;
  SpeedButton4.Down := false;
  dxFlowChart1.Zoom := 100;
end;

procedure TForm2.SpeedButton8Click(Sender: TObject);
var
  path: string;
begin
  if Form1.AdobeReaderInstalled = true then
  begin
    path := ExtractFilePath(Application.ExeName) + '\用户手册.pdf';
    ShellExecute(Handle, 'Open', PChar(path), nil, nil, SW_SHOW);
  end
  else
    Form1.bsSkinMessage1.MessageDlg('Without PDF viewer!', (mtInformation), [mbOK], 0);

end;

procedure TForm2.bsSkinButton5Click(Sender: TObject);

begin

  dxFlowChart1.SaveToFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\' + form1.Label1.Caption + '.flc');


//  Form1.ListBox1Renew;
  Flowsave := true;

  Form1.bsSkinMessage1.MessageDlg(form1.Label1.Caption + ' 保存成功!', (mtInformation), [mbOK], 0);
end;

procedure TForm2.pColorDblClick(Sender: TObject);
begin
  if ColorDialog.Execute then
  begin
    pColor.Color := ColorDialog.Color;
    pColorClick(Sender);
  end;
end;


procedure Tform2.PaintCtrlToBmp(Ctrl: TWinControl; Bmp: TBitmap);
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


procedure TForm2.fillingcolorDblClick(Sender: TObject);
begin
  if ColorDialog.Execute then
  begin
    fillingcolor.Color := ColorDialog.Color;
    fillingcolorClick(Sender);
  end;
end;

procedure TForm2.linecolorDblClick(Sender: TObject);
begin
  if ColorDialog.Execute then
  begin
    lineColor.Color := ColorDialog.Color;
    lineColorClick(Sender);
  end;
end;

procedure TForm2.bsSkinButton29Click(Sender: TObject);
begin

  panel2.Visible := true;
  bsskinpanel1.Visible := true;
end;

procedure TForm2.DBGridEh1GetCellParams(Sender: TObject; Column: TColumnEh;
  AFont: TFont; var Background: TColor; State: TGridDrawState);
begin
  if not (gdSelected in State) then
    if (Sender as TDBGridEh).SumList.RecNo mod 2 = 0 then
      Background := clInfobk;
end;





procedure TForm2.bsSkinButton1Click(Sender: TObject);
var
  Newnode, node: ttreenode;
begin
  if trim(bsSkinEdit1.Text) = '' then
  begin
    Form1.bsSkinMessage1.MessageDlg('Please input the new item name!', (mtInformation), [mbOK], 0);
    exit;
  end;

  if NameUsed(trim(bsSkinEdit1.Text), treeview1) then
  begin
    Form1.bsSkinMessage1.MessageDlg('The name ' + #39 + trim(bsSkinEdit1.Text) + #39 + ' has already exist!  Please input another name!', (mtInformation), [mbOK], 0);
    exit;
  end;

  if treeview1.Selected <> nil then
  begin
    node := treeview1.selected;
    if node.Level = 0 then
      Newnode := TreeView1.Items.Add(nil, trim(bsSkinEdit1.Text))
    else
      Newnode := TreeView1.Items.Addchild(node.Parent, trim(bsSkinEdit1.Text));
   // Newnode.MoveTo(node, naInsert);
  //  Newnode.MoveTo(node, naAdd);
  end
  else
    Newnode := treeview1.Items.Add(nil, trim(bsSkinEdit1.Text));
  treeview1.SetFocus;
  Newnode.Selected := true;
//   route := ExtractFilePath(Application.ExeName) + '\projects\' + trim(FlatEdit1.Text);
  treeview1.SaveToFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\tree.txt');
end;

function Tform2.NameUsed(name: string; treeview0: Ttreeview): boolean; //判断是否重名
var
  i: Integer;
begin
  result := false;
  for i := 0 to TreeView0.Items.Count - 1 do
    if name = TreeView0.Items[I].Text then
    begin
      result := true;
      break;
    end;
end;




procedure TForm2.bsSkinButton2Click(Sender: TObject);
var
  Newnode, node: ttreenode;
begin
  if trim(bsSkinEdit1.Text) = '' then
  begin
    Form1.bsSkinMessage1.MessageDlg('请输入名称!', (mtInformation), [mbOK], 0);
    exit;
  end;

  if NameUsed(trim(bsSkinEdit1.Text), treeview1) then
  begin
    Form1.bsSkinMessage1.MessageDlg('名称 ' + #39 + trim(bsSkinEdit1.Text) + #39 + ' 已存在，请重新输入！', (mtInformation), [mbOK], 0);
    exit;
  end;


  if treeview1.Selected = nil then exit;
  node := treeview1.Selected;
  Newnode := TreeView1.Items.AddChild(node, trim(bsSkinEdit1.Text));
  if node.Expanded = false then
    node.Expanded := true;
  Newnode.Selected := true;
  treeview1.SetFocus;

  treeview1.SaveToFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\tree.txt');

end;

procedure TForm2.bsSkinButton3Click(Sender: TObject);
begin
  if treeview1.Selected = nil then
  begin
    Form1.bsSkinMessage1.MessageDlg('Please select an item firstly!', (mtInformation), [mbOK], 0);
    exit;
  end;
  if form1.bsSkinMessage1.MessageDlg('确定删除' + treeview1.Selected.Text + '？', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    exit;
  treeview1.Items.Delete(treeview1.Selected);
  treeview1.Selected := nil;
  treeview1.SaveToFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\tree.txt');

end;

procedure TForm2.bsSkinButton4Click(Sender: TObject);
begin
  if treeview1.Selected = nil then
  begin
    Form1.bsSkinMessage1.MessageDlg('请选择需要修改的项!', (mtInformation), [mbOK], 0);
    exit;
  end;

  if trim(bsSkinEdit1.Text) = '' then
  begin
    Form1.bsSkinMessage1.MessageDlg('请输入修改名称!', (mtInformation), [mbOK], 0);
    exit;
  end;

  if NameUsed(trim(bsSkinEdit1.Text), treeview1) then
  begin
    Form1.bsSkinMessage1.MessageDlg('The name ' + #39 + trim(bsSkinEdit1.Text) + #39 + ' has already exist!  Please input another name!', (mtInformation), [mbOK], 0);
    exit;
  end;

  treeview1.Selected.Text := bsSkinEdit1.Text;
  treeview1.SaveToFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\tree.txt');

end;

procedure TForm2.DBGridEh1Enter(Sender: TObject);
begin
  Saved := false;
end;

procedure TForm2.bsSkinButton6Click(Sender: TObject);
var
  B: TBitmap;
  str: string;
begin

  B := TBitmap.Create;
   b.Width :=dxFlowChart1.Width;
    b.Height := dxFlowChart1.Height;
    dxFlowChart1.PaintTo(b.Canvas.Handle, 0, 0);


  if savedialog.Execute then
  begin
    savedialog.DefaultExt := 'bmp';
    str := savedialog.FileName + '.bmp';
  end;
  try
    if trim(str) <> '' then
      B.SaveToFile(str);
  finally
    B.Free;

  end;



  
end;

procedure TForm2.bsSkinButton11Click(Sender: TObject);
var
  B: TBitmap;
  str: string;
begin
  dxFlowChart1.SaveToFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\' + form1.Label1.Caption + '.flc');
  Flowsave := true;
  
  B := TBitmap.Create;
  //PaintCtrlToBmp(dxFlowChart1, B);

    b.Width :=dxFlowChart1.Width;
    b.Height := dxFlowChart1.Height;
    dxFlowChart1.PaintTo(b.Canvas.Handle, 0, 0);
  str :=ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\' + form1.Label1.Caption  + '.bmp';
  try
    if trim(str) <> '' then
      B.SaveToFile(str);
  finally
    B.Free;
  end;

    { b := TBitmap.Create;
  try

  
    b.SaveToFile('org.bmp');
  finally
    b.Free;
  end;}



  form1.dxFlowChart1.LoadFromFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\' + form1.Label1.Caption + '.flc');
  form1.dxFlowChart1.Zoom := 100;
  form1.PageControl1.ActivePageIndex := 1;
  close;
end;

procedure TForm2.dxFlowChart1Deletion(Sender: TdxCustomFlowChart;
  Item: TdxFcItem);
{var
  mysql1, product: string;
  Aqr1: TAdoQuery;  }
begin
{  if form2.Showing = false then exit;

  if Item is TDxFcObject then
  begin
    Aqr1 := TAdoQuery.Create(nil);
    try
      Aqr1.Connection := form1.ADOConnection2;
      mysql1 := 'select * from  data where process=' + #39 + Item.Text + #39 + ' and material_type=' + #39 + '产品' + #39;
      Aqr1.Close;
      Aqr1.SQL.Clear;
      Aqr1.SQL.Add(mysql1);
      Aqr1.Open;
      if Aqr1.RecordCount <> 0 then
      begin
        product := Aqr1.FieldByName('pd_name').AsString;

        if form1.bsSkinMessage1.MessageDlg('确定删除 ' + Item.Text + ' 及其所有数据？', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysql1 := 'delete * from  data where process=' + #39 + Item.Text + #39;
          Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add(mysql1);
          Aqr1.ExecSQL;

          mysql1 := 'delete * from LCA where product=' + #39 + product + #39;
          Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add(mysql1);
          Aqr1.ExecSQL;

          mysql1 := 'delete * from data_allocation where pd_name=' + #39 + product + #39;
          Aqr1.Close;
          Aqr1.SQL.Clear;
          Aqr1.SQL.Add(mysql1);
          Aqr1.ExecSQL;

          if Listbox1.Items.IndexOf(Item.Text) <> -1 then
            Listbox1.Items.Delete(Listbox1.Items.IndexOf(Item.Text));
          dxFlowChart1.Visible := false;
          dxFlowChart1.Visible := true;
        end
        else
          Abort;

      end;
    finally
      Aqr1.Close;
      Aqr1.Free;
    end;

  end;     }

  if Item is TDxFcObject then
  begin
    if Listbox1.Items.IndexOf(Item.Text) <> -1 then
      Listbox1.Items.Delete(Listbox1.Items.IndexOf(Item.Text));
    dxFlowChart1.Visible := false;
    dxFlowChart1.Visible := true;
  end;
end;

procedure TForm2.iRectHorizontalClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
    dxFlowChart1.SelectedConnections[i].Style := fclRectH;

end;

procedure TForm2.iNoneClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsNone;
end;

procedure TForm2.iRectangleClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsRectangle;

end;

procedure TForm2.iEllipseClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsEllipse;

end;

procedure TForm2.iRoundRectClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsRoundRect;

end;

procedure TForm2.iDiamondClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsDiamond;

end;

procedure TForm2.iNorthTriangleClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsNorthTriangle;

end;

procedure TForm2.itSouthTriangleClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsSouthTriangle;
end;

procedure TForm2.itEastTriangleClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsEastTriangle;

end;

procedure TForm2.itWestTriangleClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsWestTriangle;

end;

procedure TForm2.itHexagonClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedObjectCount - 1 do
    dxFlowChart1.SelectedObjects[i].ShapeType := fcsHexagon;
end;

procedure TForm2.iStraightClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
    dxFlowChart1.SelectedConnections[i].Style := fclStraight;

end;

procedure TForm2.iCurvedClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
    dxFlowChart1.SelectedConnections[i].Style := fclCurved;

end;

procedure TForm2.iRectVerticalClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to dxFlowChart1.SelectedConnectionCount - 1 do
    dxFlowChart1.SelectedConnections[i].Style := fclRectV;

end;

procedure TForm2.FormDestroy(Sender: TObject);
begin
  FUndo.Free;
end;

procedure TForm2.FormShow(Sender: TObject);
begin
  dxFlowChart1Change(dxFlowChart1, nil);
end;

procedure TForm2.N2Click(Sender: TObject);
var
  i, j: integer;
  s, s1, mysql1: string;
begin
  if dxFlowChart1.SelCount <> 1 then
  begin
    Form1.bsSkinMessage1.MessageDlg('只能选择一个目标重命名!', (mtInformation), [mbOK], 0);
    exit;
  end;

  s := dxFlowChart1.SelectedObject.Text;
  for i := 0 to treeview1.Items.Count - 1 do
    if treeview1.Items[i].Text = dxFlowChart1.SelectedObject.Text then
      break;

  s1 := trim(InputBox('模块重命名', '', dxFlowChart1.SelectedObject.Text));
  if s = s1 then exit;


  if listbox1.Items.IndexOf(s1) <> -1 then
  begin
    form1.bsSkinMessage1.MessageDlg(s1 + '已存在！', (mtInformation), [mbOK], 0);
    exit;
  end;
  listbox1.Items.Delete(listbox1.Items.IndexOf(s));
   listbox1.Items.Add(s1) ;

  dxFlowChart1.SelectedObject.Text := s1;
  if nameused(s1, treeview1) = false then
    treeview1.Items[i].Text := dxFlowChart1.SelectedObject.Text;

  
  treeview1.SaveToFile(ExtractFilePath(Application.ExeName) + '\projects\' + form1.Label1.Caption + '\tree.txt');

  AdoQuery1.Connection := form1.ADOConnection2;
  mysql1 := 'select *  from data where process=' + #39 + s + #39;
  if AdoQuery1.Active then AdoQuery1.Close; //载入原有数据
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(mysql1);
  ADOQuery1.Open;
  if ADOQuery1.RecordCount > 0 then
  begin
    for j := 1 to ADOQuery1.RecordCount do
    begin
      ADOQuery1.RecNo := j;
      ADOQuery1.Edit;
      ADOQuery1.FieldByName('process').AsString := dxFlowChart1.SelectedObject.Text;
      ADOQuery1.Post;
    end;
  end;

    

end;

procedure TForm2.N1Click(Sender: TObject);
begin
  FStore := False;
  dxFlowChart1.DeleteSelection;
  FStore := True;
  dxFlowChart1Change(dxFlowChart1, nil);
end;

procedure TForm2.bsSkinButton9Click(Sender: TObject);
var
  s: TStringList;
begin
  if form1.bsSkinMessage1.MessageDlg('载入流程图将覆盖原有流程图，确认载入？', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    exit;

  Form_loadproject.ListBox1.Clear;
  s := TStringList.Create;
  form1.ListDirs(ExtractFilePath(Application.ExeName) + '\projects\', s);
  Form_loadproject.ListBox1.Items.AddStrings(s);
  s.Free;

  Form_loadproject.ListBox1.Items.Delete(Form_loadproject.ListBox1.Items.IndexOf(form1.label1.Caption)); //删除自身的项目

  Form_loadproject.Label2.Caption := '载入流程图';
  Form_loadproject.ShowModal;

end;

procedure TForm2.bsSkinButton7Click(Sender: TObject);
var
  s: TStringList;
begin
  if form1.bsSkinMessage1.MessageDlg('载入目录树将覆盖现有目录树，确认载入？', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    exit;

  Form_loadproject.ListBox1.Clear;
  s := TStringList.Create;
  form1.ListDirs(ExtractFilePath(Application.ExeName) + '\projects\', s);
  Form_loadproject.ListBox1.Items.AddStrings(s);
  s.Free;

  Form_loadproject.Label2.Caption := '载入目录树';
  Form_loadproject.ShowModal;
     
end;

end.

