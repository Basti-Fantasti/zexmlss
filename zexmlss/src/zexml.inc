// If not using ZColorStringGrid
// ���� �� ������������ ZColorStringGrid 
{$DEFINE NOZCOLORSTRINGGRID}

// Conditional formatting
// �������� ��������������
{$DEFINE ZUSE_CONDITIONAL_FORMATTING}

// ���������� ������ XML, ��� �������� � ��������� �����
// Compact XML without indents and newlines
{$DEFINE NOXMLINDENTS}

// Charts
// �������
{$DEFINE ZUSE_CHARTS}

// Drawings
// �������
{$DEFINE ZUSE_DRAWINGS}

// Archivers for Delphi, not used for Free Pascal / Lazarus
{ use Delphi XE2 and above - System.Zip }
{$ifndef FPC}
{$if CompilerVersion >= 23.0}         // bds xe2 (2012)
  {$define XE2ZIP}
{$ifend}
{$endif}

// Free Pascal / Lazarus
{$IFDEF FPC}
  //��� FPC ���� ���� ZColorStringGrid-�
  {$DEFINE NOZCOLORSTRINGGRID}
  //���� ����� ������������ ������ ������������� LConvEncoding.pas
  {$DEFINE USELCONVENCODING}
{$ELSE}
  {$UNDEF USELCONVENCODING}
{$ENDIF}
