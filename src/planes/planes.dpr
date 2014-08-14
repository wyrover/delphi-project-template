program planes;

uses
  Forms,
  fmMain in 'fmMain.pas' {MainForm},
  roShell in '..\publish\roShell.pas',
  roDateTime in '..\publish\roDateTime.pas',
  roStrings in '..\publish\roStrings.pas',
  roFiles in '..\publish\roFiles.pas',
  roGraph in '..\publish\roGraph.pas',
  roUtils in '..\publish\roUtils.pas',
  roZLibEx in '..\publish\roZLibEx.pas',
  roIni in '..\publish\roIni.pas',
  fmDemo1 in 'fmDemo1.pas' {Demo1Form},
  demoFactory in 'demoFactory.pas',
  fmRotateFlip01 in 'GDIplus\fmRotateFlip01.pas' {FormRotateFlip01},
  appUtils in 'appUtils.pas',
  fmExcel in 'fmExcel.pas' {Form1},
  fmDemo2 in 'fmDemo2.pas' {Demo2Form},
  fmDemo3 in 'fmDemo3.pas' {Demo3Form},
  roAutoIE in '..\publish\roAutoIE.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
