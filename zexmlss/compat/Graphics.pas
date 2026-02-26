{ Cross-platform Graphics compatibility unit for zexmlss on non-Windows platforms.

  This unit must ONLY be on the search path for non-Windows targets (Linux64, etc.).
  On Windows, the unqualified 'Graphics' resolves to Vcl.Graphics via namespace search.

  Provides the minimal subset of Graphics types that zexmlss requires:
  TColor, TFont, TFontStyles, TFontStyle, TRect, ColorToRGB, and color constants. }
unit Graphics;

interface

uses
  System.Classes,
  System.Types,
  System.UITypes;

type
  TColor = System.UITypes.TColor;
  TFontStyle = System.UITypes.TFontStyle;
  TFontStyles = System.UITypes.TFontStyles;

const
  fsBold      = System.UITypes.TFontStyle.fsBold;
  fsItalic    = System.UITypes.TFontStyle.fsItalic;
  fsUnderline = System.UITypes.TFontStyle.fsUnderline;
  fsStrikeOut = System.UITypes.TFontStyle.fsStrikeOut;

  clBlack       = TColors.Black;
  clMaroon      = TColors.Maroon;
  clGreen       = TColors.Green;
  clOlive       = TColors.Olive;
  clNavy        = TColors.Navy;
  clPurple      = TColors.Purple;
  clTeal        = TColors.Teal;
  clGray        = TColors.Gray;
  clSilver      = TColors.Silver;
  clRed         = TColors.Red;
  clLime        = TColors.Lime;
  clYellow      = TColors.Yellow;
  clBlue        = TColors.Blue;
  clFuchsia     = TColors.Fuchsia;
  clAqua        = TColors.Aqua;
  clWhite       = TColors.White;
  clNone        = TColors.SysNone;
  clDefault     = TColors.SysDefault;
  clWindow      = TColors.SysWindow;
  clWindowText  = TColors.SysWindowText;
  clWindowFrame = TColors.SysWindowFrame;

type
  TFontPitch = (fpDefault, fpVariable, fpFixed);

  { Minimal TFont replacement for non-Windows platforms.
    Provides the properties used by zexmlss: Name, Size, Height, Style, Color, Pitch. }
  TFont = class(TPersistent)
  private
    FName: string;
    FSize: Integer;
    FHeight: Integer;
    FStyle: TFontStyles;
    FColor: TColor;
    FPitch: TFontPitch;
  public
    constructor Create;
    procedure Assign(Source: TPersistent); override;
    property Name: string read FName write FName;
    property Size: Integer read FSize write FSize;
    property Height: Integer read FHeight write FHeight;
    property Style: TFontStyles read FStyle write FStyle;
    property Color: TColor read FColor write FColor;
    property Pitch: TFontPitch read FPitch write FPitch;
  end;

/// <summary>
/// Converts a TColor value to an RGB color. On non-Windows platforms,
/// system colors (clWindow, etc.) are mapped to reasonable defaults.
/// </summary>
function ColorToRGB(Color: TColor): Integer;

implementation

constructor TFont.Create;
begin
  inherited Create;
  FName := 'Arial';
  FSize := 10;
  FHeight := 0;
  FStyle := [];
  FColor := TColor(clBlack);
  FPitch := fpDefault;
end;

procedure TFont.Assign(Source: TPersistent);
var
  LSrc: TFont;
begin
  if Source is TFont then
  begin
    LSrc := TFont(Source);
    FName := LSrc.FName;
    FSize := LSrc.FSize;
    FHeight := LSrc.FHeight;
    FStyle := LSrc.FStyle;
    FColor := LSrc.FColor;
    FPitch := LSrc.FPitch;
  end
  else
    inherited Assign(Source);
end;

function ColorToRGB(Color: TColor): Integer;
begin
  // System colors have the high byte set ($FF or $80).
  // On non-Windows platforms, map common system colors to reasonable RGB values.
  if Integer(Color) < 0 then
  begin
    case Color of
      TColor(clWindow):      Result := Integer(clWhite);
      TColor(clWindowText):  Result := Integer(clBlack);
      TColor(clWindowFrame): Result := Integer(clBlack);
    else
      // Strip system flag and return as-is
      Result := Integer(Color) and $00FFFFFF;
    end;
  end
  else
    Result := Integer(Color);
end;

end.
