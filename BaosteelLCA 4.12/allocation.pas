unit allocation;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, bsSkinCtrls, TFlatEditUnit, bsSkinBoxCtrls,
  jpeg;

type
  Taboutform = class(TForm)
    Panel1: TPanel;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  aboutform: Taboutform;

implementation

{$R *.dfm}

procedure Taboutform.Button1Click(Sender: TObject);
begin
close;
end;

end.
