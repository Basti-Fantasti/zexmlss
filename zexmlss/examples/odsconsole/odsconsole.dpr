program odsconsole;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  u_odsexport in 'u_odsexport.pas',
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
    WriteLn('Generating response.ods...');
    GenerateODS('response.ods');
    WriteLn('Done. File saved to: ', ExpandFileName('response.ods'));
  except
    on E: Exception do
    begin
      WriteLn('Error: ', E.ClassName, ': ', E.Message);
      ExitCode := 1;
    end;
  end;
end.
