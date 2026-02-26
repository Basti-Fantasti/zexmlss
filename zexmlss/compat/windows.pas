{ Cross-platform Windows compatibility unit for zexmlss on non-Windows platforms.

  This unit must ONLY be on the search path for non-Windows targets (Linux64, etc.).
  On Windows, the unqualified 'windows' resolves to Winapi.Windows via namespace search.

  Provides the minimal subset of Windows types and constants that zexmlss requires:
  HWND, TRect, GetDeviceCaps stub with related constants. }
unit windows;

interface

uses
  System.Types;

type
  HWND = NativeUInt;
  TRect = System.Types.TRect;
  TPoint = System.Types.TPoint;
  TSize = System.Types.TSize;

const
  HORZSIZE = 4;
  HORZRES  = 8;
  VERTSIZE = 6;
  VERTRES  = 10;

/// <summary>
/// Stub for GetDeviceCaps. Returns 1 to avoid division by zero.
/// Not functional on non-Windows platforms.
/// </summary>
function GetDeviceCaps(hdc: HWND; nIndex: Integer): Integer;

implementation

function GetDeviceCaps(hdc: HWND; nIndex: Integer): Integer;
begin
  Result := 1;
end;

end.
