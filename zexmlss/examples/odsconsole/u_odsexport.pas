unit u_odsexport;

interface

procedure GenerateODS(const AJsonFile, AOutputFile: string);

implementation

uses
  SysUtils, Classes, Graphics, Math,
  JsonDataObjects,
  zexmlss, zeodfs;

const
  SGL_FEE = 6;

  COL_CONFIRMED  = 0;  // A - bestaetigt
  COL_NR         = 1;  // B - Nr.
  COL_GIVEN      = 2;  // C - abgegeben
  COL_PICKED     = 3;  // D - abgeholt
  COL_FIRSTNAME  = 4;  // E - Vorname
  COL_LASTNAME   = 5;  // F - Nachname
  COL_PHONE      = 6;  // G - Telefon
  COL_EMAIL      = 7;  // H - EMail
  COL_FEE        = 8;  // I - Gebuehr
  COL_REVENUE    = 9;  // J - Umsatz
  COL_PERCENT    = 10; // K - 10%
  COL_REFUND     = 11; // L - Rueckerstattung an Verkaeufer
  COL_COMMENT    = 12; // M - Bemerkung
  COL_PIECES_ID  = 13; // N - AnzTeileID
  COL_PIECES_TXT = 14; // O - AnzTeileText

  COLOR_WHITE  = $FFFFFF;
  COLOR_GRAY   = $C0C0C0;
  COLOR_YELLOW = $00FFFF;
  COLOR_GREEN  = $00FF00;
  COLOR_BLUE   = $FF0000;

var
  FStyleNormal: Integer;
  FStyleBold: Integer;
  FStyleGray: Integer;
  FStyleYellow: Integer;
  FStyleYellowBold: Integer;
  FStyleGreen: Integer;
  FStyleGreenBold: Integer;
  FStyleBlue: Integer;
  FStyleBlueBold: Integer;

function CreateStyle(AXMLSS: TZEXMLSS; const AFontName: string;
  AFontSize: Integer; ABold: Boolean; ABGColor: TColor): Integer;
begin
  Result := AXMLSS.Styles.Count;
  AXMLSS.Styles.Count := Result + 1;
  AXMLSS.Styles[Result].Font.Name := AFontName;
  AXMLSS.Styles[Result].Font.Size := AFontSize;
  if ABold then
    AXMLSS.Styles[Result].Font.Style := [fsBold]
  else
    AXMLSS.Styles[Result].Font.Style := [];
  AXMLSS.Styles[Result].BGColor := ABGColor;
  AXMLSS.Styles[Result].CellPattern := ZPSolid;
  AXMLSS.Styles[Result].Border.Left.LineStyle := ZEContinuous;
  AXMLSS.Styles[Result].Border.Left.Weight := 1;
  AXMLSS.Styles[Result].Border.Top.LineStyle := ZEContinuous;
  AXMLSS.Styles[Result].Border.Top.Weight := 1;
  AXMLSS.Styles[Result].Border.Right.LineStyle := ZEContinuous;
  AXMLSS.Styles[Result].Border.Right.Weight := 1;
  AXMLSS.Styles[Result].Border.Bottom.LineStyle := ZEContinuous;
  AXMLSS.Styles[Result].Border.Bottom.Weight := 1;
end;

procedure InitStyles(AXMLSS: TZEXMLSS);
begin
  FStyleNormal     := CreateStyle(AXMLSS, 'Arial', 10, False, COLOR_WHITE);
  FStyleBold       := CreateStyle(AXMLSS, 'Arial', 10, True,  COLOR_WHITE);
  FStyleGray       := CreateStyle(AXMLSS, 'Arial', 10, False, COLOR_GRAY);
  FStyleYellow     := CreateStyle(AXMLSS, 'Arial', 10, False, COLOR_YELLOW);
  FStyleYellowBold := CreateStyle(AXMLSS, 'Arial', 10, True,  COLOR_YELLOW);
  FStyleGreen      := CreateStyle(AXMLSS, 'Arial', 10, False, COLOR_GREEN);
  FStyleGreenBold  := CreateStyle(AXMLSS, 'Arial', 10, True,  COLOR_GREEN);
  FStyleBlue       := CreateStyle(AXMLSS, 'Arial', 10, False, COLOR_BLUE);
  FStyleBlueBold   := CreateStyle(AXMLSS, 'Arial', 10, True,  COLOR_BLUE);
end;

procedure SetCell(ASheet: TZSheet; ACol, ARow: Integer; const AValue: string;
  ACellType: TZCellType; AStyle: Integer);
begin
  ASheet.Cell[ACol, ARow].Data := AValue;
  ASheet.Cell[ACol, ARow].CellType := ACellType;
  ASheet.Cell[ACol, ARow].CellStyle := AStyle;
end;

procedure SetFormula(ASheet: TZSheet; ACol, ARow: Integer;
  const AFormula: string; AStyle: Integer);
begin
  ASheet.Cell[ACol, ARow].Formula := AFormula;
  ASheet.Cell[ACol, ARow].CellType := ZENumber;
  ASheet.Cell[ACol, ARow].CellStyle := AStyle;
end;

procedure WriteHeader(ASheet: TZSheet);
const
  HEADERS: array[0..14] of string = (
    'bestaetigt', 'Nr.', 'abgegeben', 'abgeholt', 'Vorname',
    'Nachname', 'Telefon', 'EMail', 'Gebuehr', 'Umsatz',
    '10% v. Umsatz', 'Rueckerstattung an Verkaeufer', 'Bemerkung',
    'AnzTeileID', 'AnzTeileText'
  );
var
  I: Integer;
begin
  for I := 0 to High(HEADERS) do
    SetCell(ASheet, I, 0, HEADERS[I], ZEString, FStyleGray);
end;

procedure WriteDataRow(ASheet: TZSheet; ARow: Integer; AObj: TJsonObject);
var
  LFee: Integer;
  LPieceIdx: Integer;
  LIsActive, LWillWork: Boolean;
  LRowStr: string;
begin
  LPieceIdx := AObj.I['numpiecesid'];
  LIsActive := AObj.B['isactivemember'];
  LWillWork := AObj.B['willwork'];

  if LIsActive and LWillWork then
    LFee := 0
  else
    LFee := SGL_FEE * LPieceIdx;

  LRowStr := IntToStr(ARow + 1); // 1-based for formulas

  SetCell(ASheet, COL_CONFIRMED, ARow, '', ZEString, FStyleNormal);
  SetCell(ASheet, COL_NR, ARow, AObj.S['usereventnum'], ZEString, FStyleNormal);
  SetCell(ASheet, COL_GIVEN, ARow, '', ZEString, FStyleNormal);
  SetCell(ASheet, COL_PICKED, ARow, '', ZEString, FStyleNormal);
  SetCell(ASheet, COL_FIRSTNAME, ARow, AObj.S['firstname'], ZEString, FStyleNormal);
  SetCell(ASheet, COL_LASTNAME, ARow, AObj.S['lastname'], ZEString, FStyleNormal);
  SetCell(ASheet, COL_PHONE, ARow, AObj.S['phone'], ZEString, FStyleNormal);
  SetCell(ASheet, COL_EMAIL, ARow, AObj.S['email'], ZEString, FStyleNormal);
  SetCell(ASheet, COL_FEE, ARow, IntToStr(LFee), ZENumber, FStyleNormal);
  SetCell(ASheet, COL_REVENUE, ARow, '', ZEString, FStyleNormal);
  // K: =J{row}*0.1
  SetFormula(ASheet, COL_PERCENT, ARow,
    'of:=[.J' + LRowStr + ']*0.1', FStyleNormal);
  // L: =J{row}*0.9
  SetFormula(ASheet, COL_REFUND, ARow,
    'of:=[.J' + LRowStr + ']*0.9', FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow, '', ZEString, FStyleNormal);
  SetCell(ASheet, COL_PIECES_ID, ARow, IntToStr(LPieceIdx), ZENumber, FStyleNormal);
  SetCell(ASheet, COL_PIECES_TXT, ARow, AObj.S['numpiecestext'], ZEString, FStyleNormal);
end;

procedure WriteSummary(ASheet: TZSheet; ASummaryRow, ALastDataRow: Integer);
var
  LFirst, LLast: string;
begin
  // Python: write_summary(row) where row is 1-based, uses I2:I{row-2}
  // In 0-based: first data row = 1, last data row = ALastDataRow
  LFirst := '2'; // row 2 in 1-based (first data row)
  LLast := IntToStr(ALastDataRow + 1); // convert 0-based to 1-based

  SetCell(ASheet, COL_EMAIL, ASummaryRow, 'Summe', ZEString, FStyleYellowBold);
  // I: =SUM(I2:I{last})
  SetFormula(ASheet, COL_FEE, ASummaryRow,
    'of:=SUM([.I' + LFirst + ':.I' + LLast + '])', FStyleYellow);
  // J: =SUM(J2:J{last})
  SetFormula(ASheet, COL_REVENUE, ASummaryRow,
    'of:=SUM([.J' + LFirst + ':.J' + LLast + '])', FStyleYellow);
  // K: =SUM(K2:K{last})
  SetFormula(ASheet, COL_PERCENT, ASummaryRow,
    'of:=SUM([.K' + LFirst + ':.K' + LLast + '])', FStyleYellow);
  // L: =SUM(L2:L{last})
  SetFormula(ASheet, COL_REFUND, ASummaryRow,
    'of:=SUM([.L' + LFirst + ':.L' + LLast + '])', FStyleYellow);
end;

procedure WriteAdditionalSubtractions(ASheet: TZSheet; ARow, ASummaryRow: Integer);
var
  LSumRowStr: string;
begin
  // Python: write_additional_subtractions(row)
  //   H{row} = 'Abzug Optional / Kasse'
  //   K{row+1} = =K{row-1}-K{row}    (row-1 = summary row)
  //   L{row+1} = 'Soll Kasse'
  LSumRowStr := IntToStr(ASummaryRow + 1); // 1-based summary row

  SetCell(ASheet, COL_EMAIL, ARow, 'Abzug Optional /  Kasse', ZEString, FStyleBold);
  // K{row+1} = =K{summaryRow} - K{row}
  SetFormula(ASheet, COL_PERCENT, ARow + 1,
    'of:=[.K' + LSumRowStr + ']-[.K' + IntToStr(ARow + 1) + ']', FStyleNormal);
  SetCell(ASheet, COL_REFUND, ARow + 1, 'Soll Kasse', ZEString, FStyleBold);
end;

procedure WriteFinalIncomeAndExpense(ASheet: TZSheet; ARow, ALastDataRow,
  ASummaryRow: Integer);
var
  LSumRowStr: string;
  LRowStr: string;
begin
  LSumRowStr := IntToStr(ASummaryRow + 1); // 1-based

  // Income header (green)
  SetCell(ASheet, COL_EMAIL, ARow, 'Einnahmen', ZEString, FStyleGreenBold);
  SetCell(ASheet, COL_FEE, ARow, '', ZEString, FStyleGreen);
  // Expense header (blue)
  SetCell(ASheet, COL_REFUND, ARow, 'Ausgaben', ZEString, FStyleBlueBold);
  SetCell(ASheet, COL_COMMENT, ARow, '', ZEString, FStyleBlue);

  // Income items
  SetCell(ASheet, COL_EMAIL, ARow + 1, 'Anmeldegebuehren', ZEString, FStyleNormal);
  SetFormula(ASheet, COL_FEE, ARow + 1,
    'of:=[.I' + LSumRowStr + ']', FStyleNormal);

  SetCell(ASheet, COL_EMAIL, ARow + 2, 'Kuchenkasse', ZEString, FStyleNormal);
  SetCell(ASheet, COL_FEE, ARow + 2, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_EMAIL, ARow + 3, 'Umsatz', ZEString, FStyleNormal);
  SetFormula(ASheet, COL_FEE, ARow + 3,
    'of:=[.J' + LSumRowStr + ']', FStyleNormal);

  SetCell(ASheet, COL_EMAIL, ARow + 4, 'Sonstiges 1', ZEString, FStyleNormal);
  SetCell(ASheet, COL_FEE, ARow + 4, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_EMAIL, ARow + 5, 'Sonstiges 2', ZEString, FStyleNormal);
  SetCell(ASheet, COL_FEE, ARow + 5, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_EMAIL, ARow + 6, 'Sonstiges 3', ZEString, FStyleNormal);
  SetCell(ASheet, COL_FEE, ARow + 6, '0', ZENumber, FStyleNormal);

  // Income sum
  LRowStr := IntToStr(ARow + 1 + 1); // first income item, 1-based
  SetCell(ASheet, COL_EMAIL, ARow + 7, 'Summe Einnahmen', ZEString, FStyleBold);
  SetFormula(ASheet, COL_FEE, ARow + 7,
    'of:=SUM([.I' + LRowStr + ':.I' + IntToStr(ARow + 6 + 1) + '])', FStyleBold);

  // Expense items (start at row+2 like Python: L{row+2}..L{row+11})
  SetCell(ASheet, COL_REFUND, ARow + 2, 'Halle', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 2, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 3, 'Hausmeister', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 3, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 4, 'Auslagen', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 4, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 5, 'Zeitung', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 5, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 6, 'Wechselgeld', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 6, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 7, 'Auer', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 7, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 8, 'Rueckerstattung', ZEString, FStyleNormal);
  SetFormula(ASheet, COL_COMMENT, ARow + 8,
    'of:=[.L' + LSumRowStr + ']', FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 9, 'Sonstiges 1', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 9, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 10, 'Sonstiges 2', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 10, '0', ZENumber, FStyleNormal);

  SetCell(ASheet, COL_REFUND, ARow + 11, 'Sonstiges 3', ZEString, FStyleNormal);
  SetCell(ASheet, COL_COMMENT, ARow + 11, '0', ZENumber, FStyleNormal);

  // Expense sum
  LRowStr := IntToStr(ARow + 2 + 1); // first expense item, 1-based
  SetCell(ASheet, COL_REFUND, ARow + 12, 'Summe Ausgaben', ZEString, FStyleBold);
  SetFormula(ASheet, COL_COMMENT, ARow + 12,
    'of:=SUM([.M' + LRowStr + ':.M' + IntToStr(ARow + 11 + 1) + '])', FStyleBold);

  // Final balance (row + 14)
  SetCell(ASheet, COL_EMAIL, ARow + 14, 'Auszahlung Verkaeufer', ZEString, FStyleBold);
  SetFormula(ASheet, COL_FEE, ARow + 14,
    'of:=[.L' + LSumRowStr + ']', FStyleBold);
  SetCell(ASheet, COL_REFUND, ARow + 14, 'Gewinn', ZEString, FStyleBold);
  // Gewinn = Summe Einnahmen - Summe Ausgaben
  SetFormula(ASheet, COL_COMMENT, ARow + 14,
    'of:=[.I' + IntToStr(ARow + 7 + 1) + ']-[.M' + IntToStr(ARow + 12 + 1) + ']',
    FStyleBold);
end;

function CompareByNr(AList: TJsonArray; AIdx1, AIdx2: Integer): Integer;
var
  LNr1, LNr2: Integer;
begin
  LNr1 := StrToIntDef(AList.O[AIdx1].S['usereventnum'], MaxInt);
  LNr2 := StrToIntDef(AList.O[AIdx2].S['usereventnum'], MaxInt);
  Result := LNr1 - LNr2;
end;

procedure SortJsonArray(AArray: TJsonArray; AIndices: TArray<Integer>);
var
  I, J, LTmp: Integer;
begin
  // Simple insertion sort (stable, fine for ~120 entries)
  for I := 1 to High(AIndices) do
  begin
    LTmp := AIndices[I];
    J := I - 1;
    while (J >= 0) and (CompareByNr(AArray, AIndices[J], LTmp) > 0) do
    begin
      AIndices[J + 1] := AIndices[J];
      Dec(J);
    end;
    AIndices[J + 1] := LTmp;
  end;
end;

procedure GenerateODS(const AJsonFile, AOutputFile: string);
var
  LXMLSS: TZEXMLSS;
  LSheet: TZSheet;
  LJsonArr: TJsonArray;
  LJsonBase: TJsonBaseObject;
  LSortedIdx: TArray<Integer>;
  I, LRow, LSummaryRow, LLastDataRow: Integer;
begin
  LJsonBase := TJsonBaseObject.ParseFromFile(AJsonFile);
  try
    if not (LJsonBase is TJsonArray) then
      raise Exception.Create('JSON root must be an array');
    LJsonArr := TJsonArray(LJsonBase);

    WriteLn('  Loaded ', LJsonArr.Count, ' entries from ', AJsonFile);

    // Build sorted index array by usereventnum
    SetLength(LSortedIdx, LJsonArr.Count);
    for I := 0 to LJsonArr.Count - 1 do
      LSortedIdx[I] := I;
    SortJsonArray(LJsonArr, LSortedIdx);
    WriteLn('  Sorted by Nr. field');

    LXMLSS := TZEXMLSS.Create(nil);
    try
      InitStyles(LXMLSS);

      LXMLSS.Sheets.Count := 1;
      LSheet := LXMLSS.Sheets[0];
      LSheet.Title := 'Boerse';
      LSheet.ColCount := 15;
      // data rows + header + summary section + income/expense section + margin
      LSheet.RowCount := LJsonArr.Count + 30;

      // Page setup: landscape, 75% scale
      LSheet.SheetOptions.PortraitOrientation := False;
      LSheet.SheetOptions.ScaleToPercent := 75;

      // Column widths (mm)
      LSheet.Columns[COL_CONFIRMED].WidthMM := 20;
      LSheet.Columns[COL_NR].WidthMM := 15;
      LSheet.Columns[COL_GIVEN].WidthMM := 20;
      LSheet.Columns[COL_PICKED].WidthMM := 20;
      LSheet.Columns[COL_FIRSTNAME].WidthMM := 35;
      LSheet.Columns[COL_LASTNAME].WidthMM := 35;
      LSheet.Columns[COL_PHONE].WidthMM := 35;
      LSheet.Columns[COL_EMAIL].WidthMM := 55;
      LSheet.Columns[COL_FEE].WidthMM := 20;
      LSheet.Columns[COL_REVENUE].WidthMM := 20;
      LSheet.Columns[COL_PERCENT].WidthMM := 25;
      LSheet.Columns[COL_REFUND].WidthMM := 45;
      LSheet.Columns[COL_COMMENT].WidthMM := 35;
      LSheet.Columns[COL_PIECES_ID].WidthMM := 20;
      LSheet.Columns[COL_PIECES_TXT].WidthMM := 30;

      // Header
      WriteHeader(LSheet);

      // Data rows
      LRow := 1;
      for I := 0 to High(LSortedIdx) do
      begin
        WriteDataRow(LSheet, LRow, LJsonArr.O[LSortedIdx[I]]);
        Inc(LRow);
      end;
      LLastDataRow := LRow - 1;

      // Empty row, then summary
      Inc(LRow);
      LSummaryRow := LRow;
      WriteSummary(LSheet, LSummaryRow, LLastDataRow);
      Inc(LRow);

      // Additional subtractions
      WriteAdditionalSubtractions(LSheet, LRow, LSummaryRow);
      Inc(LRow, 3);

      // Final income and expenses
      WriteFinalIncomeAndExpense(LSheet, LRow, LLastDataRow, LSummaryRow);

      WriteLn('  Styles: ', LXMLSS.Styles.Count);
      WriteLn('  Data rows: ', LLastDataRow);
      WriteLn('  Page: landscape, 75% scale');

      if SaveXmlssToODFS(LXMLSS, AOutputFile) <> 0 then
        raise Exception.Create('SaveXmlssToODFS failed');

    finally
      LXMLSS.Free;
    end;
  finally
    LJsonBase.Free;
  end;
end;

end.
