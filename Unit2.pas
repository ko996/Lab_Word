unit Unit2;

interface

uses word_tlb;

FUNCTION fg(STR: ARRAY OF STRING): WordDocument;

implementation

FUNCTION fg(STR: ARRAY OF STRING): WordDocument;
var
  WordApp: WordApplication;
  Docs: Documents;
  Doc: WordDocument;
  Pars: Paragraphs;
  Par: Paragraph;
  D: OleVariant;
  emptyparam:OleVariant;
begin
  WordApp := CoWordApplication.Create;
  WordApp.Visible := True;

  Docs := WordApp.Documents;
  Doc := Docs.Add('Normal', False, EmptyParam, True);
  Doc := (WordApp.Documents.Item(1) as WordDocument);


    With WordApp.Selection.ParagraphFormat  do
    begin
        LeftIndent := WordApp.CentimetersToPoints(0);
        RightIndent := WordApp.CentimetersToPoints(0);
        SpaceBefore := 0;
        SpaceBeforeAuto := 0;
        SpaceAfter := 0;
        SpaceAfterAuto := 0 ;
        LineSpacingRule := wdLineSpaceSingle ;
        Alignment := wdAlignParagraphRight;
        KeepWithNext := 0 ;
        KeepTogether := 0    ;
        PageBreakBefore := 0  ;
        NoLineNumber := 0   ;
        Hyphenation := 0   ;
        FirstLineIndent := 0 ;
        OutlineLevel := wdOutlineLevelBodyText  ;
        CharacterUnitLeftIndent:= 0   ;
        CharacterUnitRightIndent := 0  ;
        CharacterUnitFirstLineIndent := 0  ;
        LineUnitBefore := 0    ;
        LineUnitAfter := 0    ;
        MirrorIndents := 0   ;
        TextboxTightWrap := wdTightNone  ;

    end;


      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 8 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphleft;
      WordApp.Selection.TypeText('                                                                                                                                                                              ����������'
      +#13+'                                                                                                                                                                              �������� ��� ������'
      +#13+'                                                                                                                                                                         �� 04.07.2002 �. � ��-3-03/342');
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphcenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeText('              B '+ str[0]);
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
        WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
//      WordApp.Selection.TypeText('                                                                                                                                      (������������ ���������� ������)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11;
      WordApp.Selection.TypeText('                �� '+ str[1]);
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
//      WordApp.Selection.TypeText('                                                                                                                                      (������������, �.�.�. �����������������)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphcenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeText('                 '+ str[2]);
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                                                      (���/��� �����������������)');
      WordApp.Selection.TypeParagraph ;
       WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphcenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeText('                 '+ str[3]+','+str[4]);
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                                                      (����� �����������������, ���.)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeParagraph ;


      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 12 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.TypeText('�����������'
      +#13+ '�� ������������� ����� �� ������������'
      +#13+ '�� ���������� ������������ �����������������, ���������'
      +#13+ '� ����������� � ������� ������ �� ����������� ���������'
      +#13+'');


      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify;
      WordApp.Selection.TypeText('         � ������������ �� ������� 145 ���������� ������� ���������� ��������� ��������� �� ������������� ����� �� ������������ �� ���������� ������������ �����������������, ��������� � ����������� � ������� ������ �� ����������� ��������� ' + str[0]+','+str[2]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
//      WordApp.Selection.TypeText('(������������, �.�.�. �����������������-���������)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify;
      WordApp.Selection.TypeText('�� ���������� ���������������� ����������� �������, ������� � ' + str[7]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                                                                                                    (�����, �����, ���)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        1. �� �������������� ��� ����������� ������ ����� ������� �� ���������� ������� (�����, �����) ��������� � ������������ '+ str[5]+' ���. ������, � ��� ����� '+str[6]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
//      WordApp.Selection.TypeText('(����������� ���������)');
//      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2. ���������, �������������� ���������� ������� �������������� ������������ �� ���������� ������������ �����������������, ��������� � ����������� � ������� ������ �� ����������� ���������, ����������� �� '+str[8]+' ������:');
      WordApp.Selection.TypeParagraph ;
       WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.1. ������� �� �������������� ������� (������������ �����������), (� ������� ������ ���� ��������� ����� ������� �� ���������� ������� (�����, �����), ���������� ������� �����������, ��������� ������������ � �������� ����������), �� '+str[9]+'  ������');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.2. ������� �� ����� ������ �� '+str[10]+' ������.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.3. ������� �� ����� ����� ������� � �������� � ������������� �������� (������������ �������������� ���������������) �� '+str[11]+' ������');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.4. ����� �������� ���������� � ������������ ������-������ �� '+str[12]+' ������.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        3. ������������ �� ���������� ����������� ������� � (���) ������������ ������������ ����� � ������� 3-� �������������� ���������������� ����������� ������� �����������.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('������������ �����������,'
      +#13+'�������������� ���������������:'
      +#13+ str[13]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                          (�������, �.�.�.)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('	                                                          �.�.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('������� ���������'
      +#13+str[14]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                          (�������, �.�.�.)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('����  �� '+str[15]+' �.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeText('-------------------------------------------------------------------------------------------------------------------------------');
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphcenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.TypeText('�������� �����');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('������� ���������� ������ � ��������� ����������� � ����������:'
      +#13+'��������� ���������� '+str[16]+' �.�.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                (����� ������)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText(str[18]+' �. '+str[17]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                          (����)                                              (�������, �.�.�. ������������ ���� ���������� ������)');
//      WordApp.Selection.TypeParagraph ;
end;


end.
