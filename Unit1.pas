unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, VBIDE_TLB, Word_TLB, Office_TLB,
  Vcl.StdCtrls, unit2, Vcl.Samples.Spin, Vcl.ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Edit2: TEdit;
    DateTimePicker1: TDateTimePicker;
    Label3: TLabel;
    Label4: TLabel;
    Edit3: TEdit;
    Edit4: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    SpinEdit1: TSpinEdit;
    Label7: TLabel;
    SpinEdit2: TSpinEdit;
    Label8: TLabel;
    SpinEdit3: TSpinEdit;
    Label9: TLabel;
    Label10: TLabel;
    SpinEdit4: TSpinEdit;
    SpinEdit5: TSpinEdit;
    Label11: TLabel;
    Edit5: TEdit;
    Label12: TLabel;
    Edit6: TEdit;
    Label13: TLabel;
    DateTimePicker2: TDateTimePicker;
    Label14: TLabel;
    SpinEdit6: TSpinEdit;
    Label15: TLabel;
    Edit7: TEdit;
    Label16: TLabel;
    DateTimePicker3: TDateTimePicker;
    Label17: TLabel;
    Edit8: TEdit;
    Label18: TLabel;
    Edit9: TEdit;
    Label19: TLabel;
    Edit10: TEdit;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  wold: WordApplication;
  Doc: WordDocument;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  masstr: array of string;
begin
  SetLength(masstr, 19);
  masstr[0] := Edit1.Text;
  masstr[1] := Edit2.Text;
  masstr[2] := Edit8.Text;
  masstr[3] := Edit9.Text;
  masstr[4] := Edit10.Text;
  masstr[5] := Edit3.Text;
  masstr[6] := Edit4.Text;
  masstr[7] := DateToStr(DateTimePicker1.datetime);
  masstr[8] := IntToStr(SpinEdit1.Value);
  masstr[9] := IntToStr(SpinEdit2.Value);
  masstr[10] := IntToStr(SpinEdit3.Value);
  masstr[11] := IntToStr(SpinEdit4.Value);
  masstr[12] := IntToStr(SpinEdit5.Value);
  masstr[13] := Edit5.Text;
  masstr[14] := Edit6.Text;
  masstr[15] := DateToStr(DateTimePicker2.datetime);
  masstr[16] := IntToStr(SpinEdit6.Value);
  masstr[17] := Edit7.Text;
  masstr[18] := DateToStr(DateTimePicker3.datetime);
  fg(masstr);

end;

end.
