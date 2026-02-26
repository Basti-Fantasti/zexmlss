program odsconsole;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  u_odsexport in 'u_odsexport.pas',
  JsonDataObjects in '..\..\..\..\..\dmvcframework\sources\JsonDataObjects.pas',
  zexmlss in '..\..\src\zexmlss.pas',
  zsspxml in '..\..\src\zsspxml.pas',
  zeodfs in '..\..\src\zeodfs.pas',
  zeformula in '..\..\src\zeformula.pas',
  zesavecommon in '..\..\src\zesavecommon.pas',
  zexmlssutils in '..\..\src\zexmlssutils.pas',
  zearchhelper in '..\..\src\zearchhelper.pas',
  zexlsx in '..\..\src\zexlsx.pas',
  zenumberformats in '..\..\src\zenumberformats.pas',
  zeZipper in '..\..\src\zeZipper.pas';

begin
  try
    WriteLn('zexmlss ODS export console example');

    if ParamCount < 1 then
    begin
      WriteLn('Usage: odsconsole <input.json> [output.ods]');
      WriteLn;
      WriteLn('  <input.json>   Path to JSON file with participant data');
      WriteLn('  [output.ods]   Output ODS file (default: response.ods)');
      ExitCode := 1;
      Exit;
    end;

    if not FileExists(ParamStr(1)) then
    begin
      WriteLn('Error: File not found: ', ParamStr(1));
      ExitCode := 1;
      Exit;
    end;

    if ParamCount >= 2 then
      GenerateODS(ParamStr(1), ParamStr(2))
    else
      GenerateODS(ParamStr(1), 'response.ods');

    if ParamCount >= 2 then
      WriteLn('Done. File saved to: ', ExpandFileName(ParamStr(2)))
    else
      WriteLn('Done. File saved to: ', ExpandFileName('response.ods'));
  except
    on E: Exception do
    begin
      WriteLn('Error: ', E.ClassName, ': ', E.Message);
      ExitCode := 1;
    end;
  end;
end.
