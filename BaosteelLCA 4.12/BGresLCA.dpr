program BGresLCA;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  flowchart in 'flowchart.pas' {Form2},
  allocation in 'allocation.pas' {aboutform},
  newproject in 'newproject.pas' {newprojectForm},
  loadproject in 'loadproject.pas' {Form_loadproject},
  ecodesign in 'ecodesign.pas' {eco};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(Taboutform, aboutform);
  Application.CreateForm(TnewprojectForm, newprojectForm);
  Application.CreateForm(TForm_loadproject, Form_loadproject);
  Application.CreateForm(Teco, eco);
  Application.Run;
end.

