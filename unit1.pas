unit Unit1; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs;

type

  { TForm1 }

  TForm1 = class(TForm)
    SaveDialog1: TSaveDialog;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

{$R *.lfm}

{ TForm1 }

uses ComObj;

{$DEFINE ExitIfNoVisio}

const
  ServerName = 'Visio.Application';

var
  i: integer;
  theText: string;
  Server, activeDoc: Variant;


procedure TForm1.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  path, filename: string;
begin
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  theText := '';
  if ParamCount > 4 then
  begin
    for i := 5 to ParamCount do
    begin
      theText := theText + ParamStr(i) + ' ';
    end;
    //ShowMessage(theText);
  end;
  try
    Server := GetActiveOleObject(ServerName);
  except
    ShowMessage('This addon can only be run from inside Visio VBA.');
    {$IFDEF ExitIfNoVisio}
    Exit;
    {$ENDIF}
  end;
  try
    activeDoc := Server.ActiveDocument;
    path := ExtractFilePath(theText);
    fileName := ExtractFileName(theText);
    SaveDialog1.InitialDir := path;
    SaveDialog1.FileName := fileName;
    if SaveDialog1.Execute then
    begin
      activeDoc.SaveAs(SaveDialog1.FileName);
    end;
  except
    ShowMessage('This addon can only be run from inside Visio VBA.');
    {$IFDEF ExitIfNoVisio}
    Exit;
    {$ENDIF}
  end;
end;

end.

