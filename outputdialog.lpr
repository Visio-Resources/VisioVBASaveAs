program outputdialog;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, Unit1;

{$R *.res}


begin
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  Application.Title:='Visio VBA Save As';
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Form1.Close;
  Application.Run;
end.

