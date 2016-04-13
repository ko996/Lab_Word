program Project1;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {Form1},
  Office_TLB in 'Office_TLB.pas',
  VBIDE_TLB in 'VBIDE_TLB.pas',
  Word_TLB in 'Word_TLB.pas',
  Unit2 in 'Unit2.pas';

{$R *.res}

begin
  vcl.Forms.Application.Initialize;
  vcl.Forms.Application.MainFormOnTaskbar := True;
  vcl.Forms.Application.CreateForm(TForm1, Form1);
  vcl.Forms.Application.Run;
end.
