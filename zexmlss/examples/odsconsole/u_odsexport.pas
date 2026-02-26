unit u_odsexport;

interface

procedure GenerateODS(const AFileName: string);

implementation

uses
  SysUtils, Graphics, zexmlss, zeodfs;

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

type
  TParticipant = record
    Nr: string;
    FirstName: string;
    LastName: string;
    Phone: string;
    Email: string;
    PiecesId: Integer;
    PiecesText: string;
    IsActiveMember: Boolean;
    WillWork: Boolean;
  end;

const
  SAMPLE_DATA: array[0..4] of TParticipant = (
    (Nr: '300'; FirstName: 'Christina'; LastName: 'Sperr';
     Phone: '017661213354'; Email: 'sperrc@aol.com';
     PiecesId: 2; PiecesText: 'Bis 100 Teile';
     IsActiveMember: False; WillWork: False),
    (Nr: '301'; FirstName: 'Jasmin'; LastName: 'Steuer';
     Phone: '015772161567'; Email: 'jasminsteuer63@gmail.com';
     PiecesId: 2; PiecesText: 'Bis 100 Teile';
     IsActiveMember: False; WillWork: False),
    (Nr: '5'; FirstName: 'Romana'; LastName: 'Baars';
     Phone: '015753341443'; Email: 'romanazoidl@web.de';
     PiecesId: 1; PiecesText: 'Bis 50 Teile';
     IsActiveMember: True; WillWork: True),
    (Nr: '302'; FirstName: 'Sandra'; LastName: 'Schmid';
     Phone: '01702743812'; Email: 'sandygoerner@web.de';
     PiecesId: 2; PiecesText: 'Bis 100 Teile';
     IsActiveMember: False; WillWork: False),
    (Nr: '310'; FirstName: 'Imke'; LastName: 'Smorz';
     Phone: '017680457321'; Email: 'imkeherrmann@web.de';
     PiecesId: 1; PiecesText: 'Bis 50 Teile';
     IsActiveMember: False; WillWork: False)
  );

function CreateStyle(AXMLSS: TZEXMLSS; const AFontName: string;
  ABold: Boolean; ABGColor: TColor): Integer;
begin
  Result := AXMLSS.Styles.Count;
  AXMLSS.Styles.Count := Result + 1;
  AXMLSS.Styles[Result].Font.Name := AFontName;
  AXMLSS.Styles[Result].Font.Size := 10;
  if ABold then
    AXMLSS.Styles[Result].Font.Style := [fsBold]
  else
    AXMLSS.Styles[Result].Font.Style := [];
  AXMLSS.Styles[Result].BGColor := ABGColor;
  AXMLSS.Styles[Result].CellPattern := ZPSolid;
end;

procedure WriteHeader(ASheet: TZSheet; AStyleGray: Integer);
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
  begin
    ASheet.Cell[I, 0].Data := HEADERS[I];
    ASheet.Cell[I, 0].CellType := ZEString;
    ASheet.Cell[I, 0].CellStyle := AStyleGray;
  end;
end;

procedure WriteDataRow(ASheet: TZSheet; ARow: Integer;
  const AData: TParticipant; AStyleNormal: Integer);
var
  LFee: Integer;
begin
  if AData.IsActiveMember and AData.WillWork then
    LFee := 0
  else
    LFee := SGL_FEE * AData.PiecesId;

  ASheet.Cell[COL_CONFIRMED, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_NR, ARow].Data := AData.Nr;
  ASheet.Cell[COL_NR, ARow].CellType := ZEString;
  ASheet.Cell[COL_NR, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_GIVEN, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_PICKED, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_FIRSTNAME, ARow].Data := AData.FirstName;
  ASheet.Cell[COL_FIRSTNAME, ARow].CellType := ZEString;
  ASheet.Cell[COL_FIRSTNAME, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_LASTNAME, ARow].Data := AData.LastName;
  ASheet.Cell[COL_LASTNAME, ARow].CellType := ZEString;
  ASheet.Cell[COL_LASTNAME, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_PHONE, ARow].Data := AData.Phone;
  ASheet.Cell[COL_PHONE, ARow].CellType := ZEString;
  ASheet.Cell[COL_PHONE, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_EMAIL, ARow].Data := AData.Email;
  ASheet.Cell[COL_EMAIL, ARow].CellType := ZEString;
  ASheet.Cell[COL_EMAIL, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_FEE, ARow].Data := IntToStr(LFee);
  ASheet.Cell[COL_FEE, ARow].CellType := ZENumber;
  ASheet.Cell[COL_FEE, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_REVENUE, ARow].CellStyle := AStyleNormal;
  // Formula: =J{row}*0.1 in OpenFormula notation (0-based row becomes 1-based)
  ASheet.Cell[COL_PERCENT, ARow].Formula :=
    'of:=[.J' + IntToStr(ARow + 1) + ']*0.1';
  ASheet.Cell[COL_PERCENT, ARow].CellType := ZENumber;
  ASheet.Cell[COL_PERCENT, ARow].CellStyle := AStyleNormal;
  // Formula: =J{row}*0.9
  ASheet.Cell[COL_REFUND, ARow].Formula :=
    'of:=[.J' + IntToStr(ARow + 1) + ']*0.9';
  ASheet.Cell[COL_REFUND, ARow].CellType := ZENumber;
  ASheet.Cell[COL_REFUND, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_COMMENT, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_PIECES_ID, ARow].Data := IntToStr(AData.PiecesId);
  ASheet.Cell[COL_PIECES_ID, ARow].CellType := ZENumber;
  ASheet.Cell[COL_PIECES_ID, ARow].CellStyle := AStyleNormal;
  ASheet.Cell[COL_PIECES_TXT, ARow].Data := AData.PiecesText;
  ASheet.Cell[COL_PIECES_TXT, ARow].CellType := ZEString;
  ASheet.Cell[COL_PIECES_TXT, ARow].CellStyle := AStyleNormal;
end;

procedure WriteSummary(ASheet: TZSheet; ARow, ALastDataRow: Integer;
  AStyleYellow, AStyleYellowBold: Integer);
var
  LFirstRow, LLastRow: string;
begin
  LFirstRow := IntToStr(2);  // row 2 (1-based, after header)
  LLastRow := IntToStr(ALastDataRow + 1);  // convert 0-based to 1-based

  ASheet.Cell[COL_EMAIL, ARow].Data := 'Summe';
  ASheet.Cell[COL_EMAIL, ARow].CellType := ZEString;
  ASheet.Cell[COL_EMAIL, ARow].CellStyle := AStyleYellowBold;

  ASheet.Cell[COL_FEE, ARow].Formula :=
    'of:=SUM([.I' + LFirstRow + ':.I' + LLastRow + '])';
  ASheet.Cell[COL_FEE, ARow].CellType := ZENumber;
  ASheet.Cell[COL_FEE, ARow].CellStyle := AStyleYellow;

  ASheet.Cell[COL_REVENUE, ARow].Formula :=
    'of:=SUM([.J' + LFirstRow + ':.J' + LLastRow + '])';
  ASheet.Cell[COL_REVENUE, ARow].CellType := ZENumber;
  ASheet.Cell[COL_REVENUE, ARow].CellStyle := AStyleYellow;

  ASheet.Cell[COL_PERCENT, ARow].Formula :=
    'of:=SUM([.K' + LFirstRow + ':.K' + LLastRow + '])';
  ASheet.Cell[COL_PERCENT, ARow].CellType := ZENumber;
  ASheet.Cell[COL_PERCENT, ARow].CellStyle := AStyleYellow;

  ASheet.Cell[COL_REFUND, ARow].Formula :=
    'of:=SUM([.L' + LFirstRow + ':.L' + LLastRow + '])';
  ASheet.Cell[COL_REFUND, ARow].CellType := ZENumber;
  ASheet.Cell[COL_REFUND, ARow].CellStyle := AStyleYellow;
end;

procedure GenerateODS(const AFileName: string);
var
  LXMLSS: TZEXMLSS;
  LSheet: TZSheet;
  LStyleGray, LStyleNormal, LStyleBold: Integer;
  LStyleYellow, LStyleYellowBold: Integer;
  I, LRow, LSummaryRow: Integer;
begin
  LXMLSS := TZEXMLSS.Create(nil);
  try
    // Create styles
    LStyleNormal := CreateStyle(LXMLSS, 'Arial', False, clWhite);
    LStyleBold := CreateStyle(LXMLSS, 'Arial', True, clWhite);
    LStyleGray := CreateStyle(LXMLSS, 'Arial', False, $C0C0C0);
    LStyleYellow := CreateStyle(LXMLSS, 'Arial', False, $00FFFF);
    LStyleYellowBold := CreateStyle(LXMLSS, 'Arial', True, $00FFFF);

    // Setup sheet
    LXMLSS.Sheets.Count := 1;
    LSheet := LXMLSS.Sheets[0];
    LSheet.Title := 'Boerse';
    LSheet.ColCount := 15;
    LSheet.RowCount := Length(SAMPLE_DATA) + 10;

    // Set column widths (in mm)
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

    // Write header
    WriteHeader(LSheet, LStyleGray);

    // Write data rows
    LRow := 1;
    for I := 0 to High(SAMPLE_DATA) do
    begin
      WriteDataRow(LSheet, LRow, SAMPLE_DATA[I], LStyleNormal);
      Inc(LRow);
    end;

    // Skip a row, write summary
    Inc(LRow);
    LSummaryRow := LRow;
    WriteSummary(LSheet, LSummaryRow, Length(SAMPLE_DATA), LStyleYellow, LStyleYellowBold);

    // Export to ODS
    WriteLn('  Styles: ', LXMLSS.Styles.Count);
    WriteLn('  Rows: ', LSheet.RowCount);
    WriteLn('  Columns: ', LSheet.ColCount);

    if SaveXmlssToODFS(LXMLSS, AFileName) <> 0 then
      raise Exception.Create('SaveXmlssToODFS failed');

  finally
    LXMLSS.Free;
  end;
end;

end.
