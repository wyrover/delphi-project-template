program planes;

uses
  Forms,
  fmMain in 'fmMain.pas' {MainForm},
  fmDemo1 in 'fmDemo1.pas' {Demo1Form},
  demoFactory in 'demoFactory.pas',
  appUtils in 'appUtils.pas',
  fmExcel in 'fmExcel.pas' {Form1},
  fmDemo2 in 'fmDemo2.pas' {Demo2Form},
  fmDemo3 in 'fmDemo3.pas' {Demo3Form};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
