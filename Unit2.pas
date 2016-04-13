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
      WordApp.Selection.TypeText('                                                                                                                                                                              УТВЕРЖДЕНО'
      +#13+'                                                                                                                                                                              приказом МНС России'
      +#13+'                                                                                                                                                                         от 04.07.2002 г. № БГ-3-03/342');
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
//      WordApp.Selection.TypeText('                                                                                                                                      (наименование налогового органа)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11;
      WordApp.Selection.TypeText('                От '+ str[1]);
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
//      WordApp.Selection.TypeText('                                                                                                                                      (наименование, Ф.И.О. налогоплательщика)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphcenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeText('                 '+ str[2]);
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                                                      (ИНН/КПП налогоплательщика)');
      WordApp.Selection.TypeParagraph ;
       WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphcenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeText('                 '+ str[3]+','+str[4]);
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                                                      (адрес налогоплательщика, тел.)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeParagraph ;


      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 12 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      WordApp.Selection.TypeText('УВЕДОМЛЕНИЕ'
      +#13+ 'ОБ ИСПОЛЬЗОВАНИИ ПРАВА НА ОСВОБОЖДЕНИЕ'
      +#13+ 'ОТ ИСПОЛНЕНИЯ ОБЯЗАННОСТЕЙ НАЛОГОПЛАТЕЛЬЩИКА, СВЯЗАННЫХ'
      +#13+ 'С ИСЧИСЛЕНИЕМ И УПЛАТОЙ НАЛОГА НА ДОБАВЛЕННУЮ СТОИМОСТЬ'
      +#13+'');


      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify;
      WordApp.Selection.TypeText('         В соответствии со статьей 145 Налогового кодекса Российской Федерации уведомляю об использовании права на освобождение от исполнения обязанностей налогоплательщика, связанных с исчислением и уплатой налога на добавленную стоимость ' + str[0]+','+str[2]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
//      WordApp.Selection.TypeText('(наименование, Ф.И.О. налогоплательщика-заявителя)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify;
      WordApp.Selection.TypeText('на двенадцать последовательных календарных месяцев, начиная с ' + str[7]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                                                                                                    (число, месяц, год)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        1. За предшествующие три календарных месяца сумма выручки от реализации товаров (работ, услуг) составила в совокупности '+ str[5]+' тыс. рублей, в том числе '+str[6]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter;
//      WordApp.Selection.TypeText('(указывается помесячно)');
//      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2. Документы, подтверждающие соблюдение условий предоставления освобождения от исполнения обязанностей налогоплательщика, связанных с исчислением и уплатой налога на добавленную стоимость, прилагаются на '+str[8]+' листах:');
      WordApp.Selection.TypeParagraph ;
       WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.1. Выписка из бухгалтерского баланса (представляют организации), (в выписке должна быть отраженна сумма выручки от реализации товаров (работ, услуг), заверенная печатью организации, подписями руководителя и главного бухгалтера), на '+str[9]+'  листах');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.2. Выписка из книги продаж на '+str[10]+' листах.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.3. Выписка из книги учета доходов и расходов и хозяйственных операций (представляют индивидуальные предприниматели) на '+str[11]+' листах');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        2.4. Копии журналов полученных и выставленных счетов-фактур на '+str[12]+' листах.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphJustify;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('        3. Деятельность по реализации подакцизных товаров и (или) подакцизного минерального сырья в течении 3-х предшествующих последовательных календарных месяцев отсутствует.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('Руководитель организации,'
      +#13+'индивидуальный предприниматель:'
      +#13+ str[13]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                          (подпись, Ф.И.О.)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('	                                                          М.П.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('Главный бухгалтер'
      +#13+str[14]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                          (подпись, Ф.И.О.)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('Дата  от '+str[15]+' г.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeText('-------------------------------------------------------------------------------------------------------------------------------');
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphcenter;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
      WordApp.Selection.TypeText('отрывная часть');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText('Отметки налогового органа о получении уведомления и документов:'
      +#13+'«Получено документов» '+str[16]+' М.П.');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                                                                                                (число листов)');
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 11 ;
      WordApp.Selection.TypeText(str[18]+' г. '+str[17]);
      WordApp.Selection.TypeParagraph ;
      WordApp.Selection.ParagraphFormat.Alignment :=  wdAlignParagraphleft;
      WordApp.Selection.Font.Name := 'Times New Roman';
      WordApp.Selection.Font.Size := 7 ;
//      WordApp.Selection.TypeText('                          (дата)                                              (подпись, Ф.И.О. должностного лица налогового органа)');
//      WordApp.Selection.TypeParagraph ;
end;


end.
