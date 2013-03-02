{ --▽---------------------------▼--
文字列処理関数
2002/08/22
  StringUnitからSysUtilsに依存しない
  部分を移動させて超軽量な
  文字列処理ユニットとして独立。
2002/09/24
  IsSameWideStrをコピーしてもってくる
2002/09/26
  KataToHiraなどが実装
2002/10/02
  StringPartsCompare系処理が大幅に変更
  IsSameWideStrの名前が変わる
2002/10/08
  DeleteStrにWideString版を追加
2002/11/01
  RangeAnsiPosなどすべてにIgnoreCaseフラグをつけて
  大小文字を区別しない検索を実装
2003/02/27
  StrCountを実装
2003/03/04
  CheckStrInTableを実装
2003/03/13
  ChangeLineBreakesをStringUnit(Heavy)から移動
  StringsReplace/ToEscapeSequence/FromEscapeSequenceを実装
2003/03/30
  BackAnsiPosを追加
  AnsiStringPartsCompare_Base/WideStringPartsCompare_Baseを
  interface節から削除、implementation節のforward参照に変更
  WideStringsReplaceを作成        
  WideStringReplaceAllの内部を変更した
2003/06/15
  ToEscapeSequence、FromEscapeSequence関数を
  テーブル変換方式にしてWideString対応に変更
  ConvertHanKataToZenKata、ConvertZenKataToHanKataを実装した
2004/04/14
  エンコード処理にヌル文字は\0に置き換える処理を追加
2004/09/19
  文字列を繰り返して出力するStringOfStr関数を追加
2005/02/10
  LCMapStringWを使ったToUpperCase/ToLowerCaseを追加
2005/09/22
・順方向が[forward direction]、逆方向が[backward direction]という
  英単語なのでAnsiPosをAnsiPosForward、BackAnsiPosをAnsiPosBackwardと
  関数名を変更してみた。同様にWidePosも影響度が大きそうだけど気にしない。(^-^;
  だってリファクタリングに躊躇したくないし、、、
  AnsiPosって関数名はSysUtilsと被るから不味いし。
・InStr関数を作った
2005/12/29
・InStr関数にIndexとCount引数版も追加した
・CheckWideCharInTableを作成した
・CheckStrInTableを使いやすく変更した
2006/04/22
・WordWrap処理のために
  WideCharByteLength/CharIndexToByteIndex/ByteIndexToCharIndex/ByteLengthを
  作成
2006/05/13
・WideCharByteLengthに半角カナ系のバグがあったので修正
2006/08/02
・OneTrimCharを追加した
2006/11/14
・DeleteStrInTableを追加した
2007/05/02
・CompareStringAnsiAPI/CompareStringWideAPIを追加した
2007/07/25
・DeleteStrのWideString版が思いっきりバグっていた...5年目にして初めて使ったか...
・Include/ExcludeBothEndsStr関数を追加
2007/08/01
・const群をConstUnit.pasに移動
・IncludeFirstStr等をinterface部に記述
2007/08/06
・IncludeLastPathDelim/ExcludeLastPathDelimを追加
2008/01/15
・StringUnitLightからStringUnitに名前変更
2008/03/04
・DelimiterCut.pasの関数を移動した
2008/11/18
・testStringUnit.pasを作成、とりあえずtestTrim系処理を
  そちらに移動した。
・文字列定数をConstUnit.pasに移動した。
2009/01/07
・TTextLineBreakStyleMultiPlatformを排除してTLineBreakStyleに統一する
・ChangeLineBreakesをTTextLineBreakStyleMultiPlatformからTLineBreakStyleに変更
・IsLineBreakをLineBreakStyleに名前変更
2010/03/04(木)
・	testコードをすべてtestStringUnit.pasに移行
	動作させた
・	StringReplaceをオープン配列パラメータ指定として
	引数に[str1, str2]という形での指定ができるようにした。
・ CopyIndexを作成した
2010/03/10(水)
・	ReplceStringにReplaceAppFlagをつけた
2010/03/10(水)
・	IsFirstStr/IsLastStr/IncludeFirstStr等にIgnoreCaseフラグを追加した
2010/03/30(火)
・	EncodeStringToUTF16CodeCsv/DecodeStringFromUTF16CodeCsvを追加
2010/10/29
・TabToSpaceLine/AddStringPreLineBreakをWideStrRecListUnitから移動
2010/11/10
・TrimLeftCharをTrimFirstCharへ名前変更しました。
・GetTagText/GetTagInfoを実装。
2011/04/22(金)
・AnsiStringPartsCompareやAnsiPosなどAnsi系のコードの引数を
  String>>AnsiStringに変更しました
・Ansi/Wide系の関数からUnicodeString版を作成しました。
・InStrもAnsi系Wide系を分離してUnicodeString対応しました。
2011/05/09(月)
・さまざまな関数をUnicodeString対応しました。
2011/05/11(水)
・Trim系の処理もUnicodeString対応しました。
2011/06/07(火)
・MECSUtilsをusesしてRight/LeftAlignTextを実装しました。
2011/08/12(金)
・容量の関係からMECSUtilsを分離。Right/LeftAlignTextをMECSUtilsUnit.pas
2011/12/17(土)
・TrimLeft/TrimRight/Trimを廃止(SysUtilsで実装されている)
・DelimiterLeft/Right等を廃止
2011/12/20(火)
・PosForward/Backward等をStringSearchUnitへ移動
  Search関数で置き換える
・コメントや古いコードを削除した
2011/12/22(木)
・StringPartsCompareやStringCountを修正した
・変換系の処理をStringConvertUnitへ移動
2012/08/21(火)
・TStringDynArrayを削除。array of Stringを使うようにした。
//--▲---------------------------△-- }
unit StringUnit;

interface

uses
//  Types,
  Windows,
  SysUtils,     //Exception
  ConstUnit,
  MathUnit,
uses_end;

type
//Typesユニットからコピペ
//  TStringDynArray       = array of string;

  TMbcsByteType = (mbSingleByte, mbLeadByte, mbTrailByte);
{mbSingle=半角 mbLead=全角の1バイト目 mbTrail=全角の2バイト目}

type
  TCaseCompare = (ccCaseSensitive, ccIgnoreCase);

function IsDBCSLeadChar(c:AnsiChar):boolean;
function ByteType(const S: AnsiString; Index: Integer): TMbcsByteType;
function CharToByteLen(const S: AnsiString; MaxLen: Integer): Integer;

function StringPartsCompare(const SubStr, S: String; StrIndex: Integer;
 CaseCompare: TCaseCompare = ccCaseSensitive): Boolean;
function StringPartsCompare_BaseSafe(S1, S2: String;
 S1Index, S2Index, CompareLength: Integer;
 CaseCompare: TCaseCompare): Boolean;

type TCompareResult = (crEqual, crLessThan, crGreaterThan);
function CompareStringAnsiAPI(const S1, S2: AnsiString;
 S1Index, S2Index, CompareS1Length, CompareS2Length: Integer;
  CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;
function CompareStringAnsi(const S1, S2: AnsiString;
 CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;
function CompareStringWideAPI(const S1, S2: WideString;
 S1Index, S2index, CompareS1Length, CompareS2Length: Integer;
  CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;
function CompareStringWide(const S1, S2: WideString;
 CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;

function DeleteStr(const S: String; const Index, Count: Integer): String; overload;
function DeleteStrIndex(const S: String; const StartIndex, EndIndex: Integer): String overload;
function DeleteFirstStr(const S: String; const Count: Integer): String overload;
function DeleteLastStr(const S: String; const Count: Integer): String overload;
function CopyIndex(const S: String; StartIndex, EndIndex: Integer): String; overload;

function KataToHira(const Source: String): String;
function HiraToKata(const Source: String): String;
function ZenkakuToHankaku(const Source: String): String;
function HankakuToZenkaku(const Source: String): String;
function ToUpperCase(const Source: String): String;
function ToLowerCase(const Source: String): String;

type TStringCountFlag = (scfIncChar, scfIncSubStr);
function StringCount(SubStr, S: String; Flag: TStringCountFlag = scfIncSubStr;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): Integer;

type TInTable = (itUnknown, itAllInclude, itAllExclude, itPartInclude);
function CheckStrInTable(const Str, Table: String): TInTable;
function CheckCharInTable(const Char: Char;
 const Table: String): Boolean;

//type TTextLineBreakStyleMultiPlatform = (tlbsmpCRLF, tlbsmpCR, tlbsmpLF);
//  //CRLF:Windows    CR:Mac  LF:Unix/Linux
//function ChangeLineBreakes(const S: String; Style: TTextLineBreakStyleMultiPlatform): String;

type TLineBreakStyle = (lbsCRLF, lbsCR, lbsLF ,lbsNoLineBreaks);
//改行コードは、CRLF:Windows CR:旧Mac  LF:Unix/Linux/MacOSX
function ChangeLineBreakes(const S: String; Style: TLineBreakStyle): String;

function LastLineBreakStyle(S: WideString): TLineBreakStyle;
procedure ExcludeLineBreakProc(var S: String);
function ExcludeLineBreak(const S: String): String;
function LineBreakString(Style: TLineBreakStyle): String;
function LineBreakStyle(S: WideString): TLineBreakStyle;
function AddStringPreLineBreak(const LineStr, AddStr: WideString): WideString;

function StringsReplace(const S: String; OldPatterns, NewPatterns: array of string;
 CaseCompare: TCaseCompare = ccCaseSensitive; ReplaceAll: Boolean = True): string;

function EncodeEscapeSequence(const Source: String): String;
function DecodeEscapeSequence(const Source: String): String;

type
  TEndian = (eBig, eLittle);
function EncodeStringToUTF16CodeCsv(const Source: WideString;
 Endian: TEndian = eLittle): String;
function DecodeStringFromUTF16CodeCsv(Source: String;
 Endian: TEndian = eLittle): WideString;
function EncodeStringToUTF8CodeCsv(const Source: WideString): String;
function DecodeStringFromUTF8CodeCsv(Source: String): WideString;

function StringsWordReplace(const S: WideString;
 OldPatterns, NewPatterns: array of string;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;

function TrimFirstCharCount(const S, Table: String): Integer;
function TrimFirstChar(const S, Table: String): String;
function TrimLastCharCount(const S, Table: String): Integer;
function TrimLastChar(const S, Table: String): String;
function TrimChar(const S, Table: String): String;
function OneTrimChar(S: String; Table: String): String;

function TrimFirst(const S: string): String;
function TrimLast(const S: string): String;

type
  TTagTextInfo = record
    StartIndex: Integer;
    EndIndex: Integer;
  end;
function GetTagInfo(StartTag, EndTag, TargetText: WideString;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): TTagTextInfo;
function GetTagText(StartTag, EndTag, TargetText: WideString;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): WideString;

function StringOfStr(Str: String; Count: Integer): String;
type TCharByteLength = (wcblSingle, wcblMulti);
function WideCharByteLength(source: WideChar): TCharByteLength;

function InStr(const SubStr, S: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): Boolean; overload;
function InStr(const SubStr, S: String; Index, Count: Integer;
 CaseCompare: TCaseCompare = ccCaseSensitive): Boolean; overload;



function CharIndexToByteIndex(Source: WideString; CharIndex: Integer): Integer;
function ByteIndexToCharIndex(Source: WideString; ByteIndex: Integer): Integer;
function ByteLength(Source: WideString): Integer;

function DeleteStrInTable(const S, Table: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;

function IsFirstStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): Boolean;
function IsLastStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): Boolean;
function IncludeFirstStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): String;
function IncludeLastStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): String;
function IncludeBothEndsStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): String;
function ExcludeFirstStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): String;
function ExcludeLastStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): String;
function ExcludeBothEndsStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): String;
function IncludeLastPathDelim(const Path: String): String;
function ExcludeLastPathDelim(const Path: String): String;


function FirstString(s: String; Delimiter: String): String;
function LastString(s: String; Delimiter: String): String;
function FirstStringLong(s: String; Delimiter: String): String;
function LastStringLong(s: String; Delimiter: String): String;

//下位互換関数
function DelimiterRight(Delimiter, Str: String): String;
function DelimiterLeft(Delimiter, Str: String): String;
function DelimiterLeftLong(Delimiter, Str: String): String;
function DelimiterRightLong(Delimiter, Str: String): String;

procedure StringAdd(var SourceText: String; AddText, ConnectString: String);
procedure StringLineAdd(var SourceText: String; AddText: String);

function GetStringConnect(ConnectText: String; SourceTests: array of String): String;
function GetStringLineConnect(SourceTests: array of String): String;

function TabToSpaceLine(const S: WideString): WideString;

implementation

uses
  StringSearchUnit,
  StringConvertUnit,
end_uses;


//(*--▽---------------------------▼--
//SysUtilsの代わりに実装している機能
//※DBCSはDouble-Byte Character Setの略
//※MBCSはMultiByte Character Set の略

    function IsDBCSLeadChar(c:AnsiChar):boolean;
    begin
      Result:=c in [#$81..#$9F,#$E0..#$FC]
    end;

    {↓SysUtils.ByteTypeをWinAPIを使って実装
       これを使うとSysUtilsをusesしなくても
       RangeAnsiPos(=AnsiPos)が実装できる}
    function ByteType(const S: AnsiString; Index: Integer): TMbcsByteType;
    var
      i: Integer;
    begin
      if Index = 1 then
      begin
        if IsDBCSLeadChar(S[Index]) then
        begin
          Result := mbLeadByte;
        end else
        begin
          Result := mbSingleByte;
        end;
      end else
      begin
        i := Index-1;
        while IsDBCSLeadChar(S[i]) do
        begin
          Dec(i);
          if i = 0 then break;
        end;
        if Odd(Index - i) then
        begin
          if IsDBCSLeadChar(S[Index]) then
          begin
            Result := mbLeadByte;
          end else
          begin
            Result := mbSingleByte;
          end;
        end else
        begin
          Result := mbTrailByte;
        end;
      end;
    end;

    //MaxLenまでの文字数までに何Byteあるかカウントする
    //SysUtilsにも同名関数があるが、別な実装方法(互換性不明)
    function CharToByteLen(const S: AnsiString; MaxLen: Integer): Integer;
    var
      i: Integer;
      SLen: Integer;
      SingleByteCount, DoubleByteCount: Integer;
    begin
      Result := 0;
      if MaxLen <= 0 then Exit;
      SLen := length(S);
      if SLen=0 then Exit;
      if SLen < MaxLen then MaxLen := SLen;

      SingleByteCount := 0;
      DoubleByteCount := 0;
      i:=1; while i<=SLen do
      begin
        case ByteType(S, i) of
          mbSingleByte:
            Inc(SingleByteCount);
          mbTrailByte:
            Inc(DoubleByteCount);
          mbLeadByte:;
        else
          Assert(False, 'エラー');
        end;
        if (SingleByteCount + DoubleByteCount) = MaxLen then
        begin
          Result := i;
          Exit;
        end;

        Inc(i);
      end;
      Result := i - 1;
    end;


//--▲---------------------------△--*)


{---------------------------------------
    文字列の範囲一致比較
機能:   
機能:   2つの文字列のIndex指定の比較
        StringPartsCompare(S, SubStr, Index)とすると
        SのIndex位置からSubStrが一致するかどうかを判定できる

備考:   _Baseの関数は
        内部で使うエラー処理の無い関数
        (エラーかどうか判定する分が無駄になるから)
参考：  Windows-API による文字列比較オブジェクト
        http://www.s34.co.jp/cpptechdoc/article/comparestring/index.html
履歴:   2002/09/29
            最初はRangeAnsiPosの為に作ったが
            応用範囲が広く他でも使用するようになったので
            関数名InSameStrなどから変更した
        2002/10/01
            _Base関数でCompareStringA/WのAPIを使うようにしたので
            大小文字関係ない比較が容易になった
        2002/10/02
            コード共通化のため_BaseでIgnoreFlagを使用して
            比較に大小文字の影響を切り替えられるようにした
            AnsiStringPartsCompare_BaseでShiftJISの2Byte文字が
            途中Byteで区切られている場合も正しく比較するようにした
            (多少OverHeadがあるが)
        2011/05/09(月)
        ・  UnicodeString対応しました
}//(*-----------------------------------
function StringPartsCompare_Base(S1, S2: String;
 S1Index, S2Index, CompareLength: Integer;
 CaseCompare: TCaseCompare): Boolean;
var
  I: Integer;
begin
  case CaseCompare of
    ccCaseSensitive:;
    ccIgnoreCase:
    begin
      StringConvert(S1, scUpperCase);
      StringConvert(S2, scUpperCase);
    end;
  end;

  for I := 0 to CompareLength-1 do
  begin
    if S1[S1Index + I] <> S2[S2Index + I] then
    begin
      Result := False;
      Exit;
    end;
  end;

  Result := True;
end;

function StringPartsCompare_BaseSafe(S1, S2: String;
 S1Index, S2Index, CompareLength: Integer;
 CaseCompare: TCaseCompare): Boolean;
begin
  Result := False;
  if (S1 = EmptyStr) or (S2 = EmptyStr) then Exit;
  if CheckRange(1, S1Index, Length(S1) - CompareLength + 1) = False then Exit;
  if CheckRange(1, S2Index, Length(S2) - CompareLength + 1) = False then Exit;

  Result :=
    StringPartsCompare_Base(S1, S2, S1Index, S2Index, CompareLength, CaseCompare);
end;

function StringPartsCompare(const SubStr, S: String; StrIndex: Integer;
 CaseCompare: TCaseCompare = ccCaseSensitive): Boolean;
begin
  Result := False;
  if (S = EmptyStr) or (SubStr = EmptyStr) then Exit;
  if CheckRange(1, StrIndex, Length(S) - Length(SubStr) + 1) = False then Exit;

  Result := StringPartsCompare_Base(SubStr, S, 1, StrIndex, Length(SubStr), CaseCompare);
end;
//------------------------------------*)

function CompareStringAnsiAPI(const S1, S2: AnsiString;
 S1Index, S2Index, CompareS1Length, CompareS2Length: Integer;
 CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;
var
  CompareFlag: longword;
begin
  if CaseCompare = ccIgnoreCase then    CompareFlag := NORM_IGNORECASE
                                else    CompareFlag := 0;

  case CompareStringA(LOCALE_USER_DEFAULT, CompareFlag,
         PAnsiChar(S1) + S1Index - 1, CompareS1Length,
         PAnsiChar(S2) + S2Index - 1, CompareS2Length) of
    CSTR_LESS_THAN:
      Result := crLessThan;
    CSTR_EQUAL:
      Result := crEqual;
    CSTR_GREATER_THAN:
      Result := crGreaterThan;
  else
    raise Exception.Create('CompareStringAの戻り値が不正です');
  end;
end;

function CompareStringAnsi(const S1, S2: AnsiString; CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;
begin
  Result := CompareStringAnsiAPI(S1, S2, 1, 1, Length(S1), Length(S2), CaseCompare);
end;

function CompareStringWideAPI(const S1, S2: WideString;
 S1Index, S2index, CompareS1Length, CompareS2Length: Integer;
 CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;
var
  CompareFlag: longword;
begin
  if CaseCompare = ccIgnoreCase then    CompareFlag := NORM_IGNORECASE
                                else    CompareFlag := 0;

  case CompareStringW(LOCALE_USER_DEFAULT, CompareFlag,
         PWideChar(S1) + S1Index - 1, CompareS1Length,
         PWideChar(S2) + S2Index - 1, CompareS2Length) of
    CSTR_LESS_THAN:
      Result := crLessThan;
    CSTR_EQUAL:
      Result := crEqual;
    CSTR_GREATER_THAN:
      Result := crGreaterThan;
  else
    raise Exception.Create('CompareStringWの戻り値が不正です');
  end;
end;

function CompareStringWide(const S1, S2: WideString; CaseCompare: TCaseCompare{IgnoreCase: Boolean}): TCompareResult;
begin
  Result := CompareStringWideAPI(S1, S2, 1, 1, Length(S1), Length(S2), CaseCompare);
end;
//--△----------------------▲--


{-----------------------------------------
//      戻り値を返すDelete
機能:       
備考:       おそらくメモリ確保を考えると
			Deleteをそのまま使うより動作は遅い
履歴:       だいぶ以前
//--▼--------------------------------▽--}
function DeleteStr(const S: String; const Index, Count: Integer): String; overload;
begin
  Result := S;
  Delete(Result, Index, Count);
end;

//StartIndexとEndIndexを指定したDeleteStr
function DeleteStrIndex(const S: String; const StartIndex, EndIndex: Integer): String overload;
begin
  Result := S;
  Delete(Result, StartIndex, EndIndex - StartIndex + 1);
end;
//※procedure DeleteIndex も作れると思う

//先頭文字を削除する
function DeleteFirstStr(const S: String; const Count: Integer): String overload;
begin
  Result := S;
  Delete(Result, 1, Count);
end;
//※procedure DeleteFirst も作れると思う

//終端文字を削除する
function DeleteLastStr(const S: String; const Count: Integer): String overload;
begin
  Result := S;
  Delete(Result, Length(S)-Count+1, Count);
end;
//※procedure DeleteLast も作れると思う


//最後から何文字かを削除する手続き
procedure DeleteEndCount(var S: String; Count: Integer);
begin
  Delete(S, Length(S)-Count+1, Count);
end;
//--△--------------------------------▲--

{----------------------------------------
//      Indexを指定するCopy
機能:       通常のCopyはIndexとCountだが
			CopyIndexはStartIndexとEndIndexを指定する
備考:       
履歴:       2010/03/04(木)
//----------------------------------------}
function CopyIndex(const S: String; StartIndex, EndIndex: Integer): String; overload;
begin
  if StartIndex < 1 then StartIndex := 1;
  Result := Copy(S, StartIndex, EndIndex - StartIndex + 1);
end;
//----------------------------------------

{-------------------------------
//  WideStringで文字列変換
機能:       APIのLCMapStringWを用いて
              ひらがな⇔カタカナ
              全角⇔半角
              大文字⇔小文字
            の変換処理をします
引数説明:   dwMapFlags:変換処理していフラグ
備考:       KanaToHira/HiraToKana
            /ZenkakuToHankaku/HankakuToZenkaku
            /ToUpperCase/ToLowerCase
履歴:       2002/09/26
            2005/02/10 ToUpper/ToLowerを追加
//--▼----------------------▽--}
function MapStringW(const Source: String; dwMapFlags: Longword): String;
var
  Len: Integer;
begin
  Result := '';
  Len := LCMapStringW(LOCALE_USER_DEFAULT, dwMapFlags,
           PWideChar(Source), -1, nil, 0);
  SetLength(Result, Len-1);
  LCMapStringW(LOCALE_USER_DEFAULT, dwMapFlags,
    PWideChar(Source), Length(Source)+1,
    PWideChar(Result), Len);
  //LOCALE_USER_DEFAULTの代わりに
  //GetUserDefaultLCIDを指定してもいいみたい
end;

function KataToHira(const Source: String): String;
begin
  Result := MapStringW(Source, LCMAP_HIRAGANA);
end;

function HiraToKata(const Source: String): String;
begin
  Result := MapStringW(Source, LCMAP_KATAKANA);
end;

function ZenkakuToHankaku(const Source: String): String;
begin
  Result := MapStringW(Source, LCMAP_HALFWIDTH);
end;

function HankakuToZenkaku(const Source: String): String;
begin
  Result := MapStringW(Source, LCMAP_FULLWIDTH);
end;

function ToUpperCase(const Source: String): String;
begin
  Result := MapStringW(Source, LCMAP_UPPERCASE);
end;

function ToLowerCase(const Source: String): String;
begin
  Result := MapStringW(Source, LCMAP_LOWERCASE);
end;
//--△----------------------▲--




{---------------------------------------
    文字列をカウントする
機能:   StringCount('ＡＡ', 'ＡＡＡＡＡ', scfIncSubStr)
                          ￣￣
                              ￣￣
            Result=2;   検索文字分を増加してカウント

        StringCount('ＡＡ', 'ＡＡＡＡＡ', scfIncChar)
                          ￣￣
                            ￣￣
                              ￣￣
                                ￣￣
            Result=4;   1文字分を増加してカウント
備考:   UnicodeString対応版
履歴:   2011/04/12(火)
        ・  作成
}//(*-----------------------------------
//function StringCount(const SubStr, S: String; Flag: TStringCountFlag = scfIncSubStr;
// CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): Integer;
//
//type
//  TCompareFunction = function (const S1, S2: string): Boolean;
//
//var
//  i, j: Integer;
//  SubStrLength, SLength: Integer;
//  SameCheckFlag: Boolean;
//  Compare: TCompareFunction;
//begin
//  Result := 0;
//  if (SubStr='') or (S='') then Exit;
//
//  if CaseCompare = ccIgnoreCase then
//    Compare := SameText
//  else
//    Compare := SameStr;
//
//  SubStrLength := Length(SubStr);
//  SLength := Length(S);
//
//  i := 1;
//  while i <= SLength do
//  begin
//    {↓先頭文字が検索したい文字と一致するなら}
//    if Compare( S[i] , SubStr[1]) then
//    begin
//      SameCheckFlag := True;
//      for j := 2 to SubStrLength do
//      begin
//        if (not Compare( SubStr[j] , S[i+j-1]) ) then
//        begin
//          SameCheckFlag := False;
//          break;
//        end
//      end; //for j
//      if SameCheckFlag then
//      begin
//        Inc(Result);
//        if Flag = scfIncSubStr then
//          Inc(i, SubStrLength-1);
//      end;
//    end;
//    Inc(i);
//  end;
//end;

function StringCount(SubStr, S: String; Flag: TStringCountFlag = scfIncSubStr;
 CaseCompare: TCaseCompare = ccCaseSensitive): Integer;
var
  I, J: Integer;
  SubStrLength, SLength: Integer;
begin
  Result := 0;
  if (SubStr = EmptyStr) or (S = EmptyStr) then Exit;

  case CaseCompare of
   ccCaseSensitive: ;
   ccIgnoreCase:
   begin
     StringConvert(S, scUpperCase);
     StringConvert(SubStr, scUpperCase);
   end;
  end;

  SLength := Length(S);
  SubStrLength := Length(SubStr);

  I := 1;
  while I <= SLength do
  begin
    if StringPartsCompare_Base(SubStr, S, 1, I, Length(SubStr), ccCaseSensitive) then
    begin
      Inc(Result);
      if Flag = scfIncSubStr then
        Inc(I, SubStrLength - 1);
    end;
    Inc(I);
  end;
end;
//------------------------------------*)

{-------------------------------
//  CheckStrInTable
機能:       StrがTable文字列の中の文字群に
            『すべて含まれるか』AllInclude
            もしくは
            『すべて含まれないか』AllExclude
            を判定します。
引数説明:   Str:判定文字列 Table:文字列テーブル
戻り値:     InTable
              itUnknown     判定不可能(空文字が入る場合があるから)
              itAllInclude  すべて含まれる
              itAllExclude  ひとつも含まれない
              itPartInclude 含まれるものもある
            動作結果は testCheckStrInTable を参照のこと
備考:       例えばTable="0123456789"とセットすると
            strが数値のみで成り立っているのかはわかる。
履歴:       2002/03/10
            2005/12/28
              わかりににくいので
              InTableフラグで動作を変更するのではなく
              戻り値でAllIncludeとAllExcludeを判断するようにした
//--▼----------------------▽--}
function CheckStrInTable(const Str, Table: String): TInTable;
var
  i: Integer;
begin
  Result := itUnknown; {←Strが空文字の場合にはこれを返す}
  if (Str = EmptyStr) or (Table = EmptyStr) then Exit;

  if CheckCharInTable(Str[1], Table) then
  begin {↓文字列が含まれるなら}
    Result := itAllInclude;
  end else
  begin {↓文字列が含まれないなら}
    Result := itAllExclude;
  end;

  for i:=2 to Length(Str) do
  begin
    if CheckCharInTable(Str[i], Table) then
    begin {↓文字列が含まれる場合}
      if Result <> itAllInclude then
      begin
        Result := itPartInclude;
        break;
      end;
    end else
    begin {↓文字列が含まれない場合}
      if Result <> itAllExclude then
      begin
        Result := itPartInclude;
        break;
      end;
    end;
  end;
end;

//--△----------------------▲--

{-------------------------------
//  CheckCharInTable
機能:       Charが文字列に含まれるかどうかを調べる関数
備考:       高速化できそうだがとりあえずWidePosForwardで実装した
履歴:       2005/12/29
            2011/10/05(水)
            ・CheckWideCharInTableからCheckCharInTableに変更
            2011/12/27(火)
            ・InStrなので廃止っぽい
            2012/01/11(水)
            ・高速化してしまった。
//--▼----------------------▽--}
//function CheckCharInTable(const Char: Char;
// const Table: String): Boolean;
//begin
//  Result := (1 <= PosForward(Char, Table));
//end;

function CheckCharInTable(const Char: Char;
 const Table: String): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 1 to Length(Table) do
  begin
    if Table[I] = Char then
    begin
      Result := True;
      Exit;
    end;
  end;
end;
//--△----------------------▲--

{-------------------------------
//  文字列が含まれるかどうか判断します
備考:       AnsiPosの条件判断をいつも間違うのでちょっと作った
            AnsiPosは文字列が無い場合0を返しますね。
履歴:       2005/09/22
			2011/04/22(金)
			・	Ansi版作成、UnicodeString対応
//--▼----------------------▽--}
function InStr(const SubStr, S: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): Boolean; overload;
begin
  Result := (1 <= (PosForward(SubStr, S, CaseCompare)));
end;

function InStr(const SubStr, S: String; Index, Count: Integer;
 CaseCompare: TCaseCompare = ccCaseSensitive): Boolean; overload;
begin
  Result := (1 <= (PosForward(SubStr, S, Index, Count, CaseCompare)));
end;
//--△----------------------▲--

{-------------------------------
//文字列の改行コードをそろえます
機能:       改行コードを
            WinCRLF形式 MacCR形式 UnixLF形式
            で相互変換します
引数説明:   S: 元の文字列
            Style: 変換する形式指定
戻り値:     変換された文字列
備考:
履歴:       2002/03/16
            2009/01/07
//--▼----------------------▽--}
function ChangeLineBreakes(const S: String; Style: TLineBreakStyle): String;
type StrArray = array[1..$10] of char;

 //CRとLFがいくつあるかをカウントする
 function CountCRLF(const S: String): Integer;
 var
   i: Integer;
   PS: ^StrArray;
 begin
   Result := 0;
   PS := @S[1];
   for i := 1 to Length(S) do
   begin
     case PS^[i] of
       CR: begin
           Inc(Result);
       end;
       LF: begin
           Inc(Result);
       end;
     else ;
     end;
   end; //for i
 end;

var
  ReadIndex, WriteIndex, SourceLength: Integer;
  ReplaceStr: array[0..1] of Char;
  ReplaceStrLen: Integer;
  PResultStr, PSourceStr: ^StrArray;
begin
  if S = '' then begin Result := ''; Exit; end;

  SourceLength := Length(S);
  case Style of
    lbsCRLF:
    begin
      ReplaceStr := CR+LF;
      ReplaceStrLen := 2;
      SetLength(Result, SourceLength+CountCRLF(S));
      //↑最大で改行コードの数だけ
      //  文字列長が増加する可能性があるので
      //  メモリを確保している
    end;

    lbsCR:
    begin
      ReplaceStr := CR+#0;
      ReplaceStrLen := 1;
      SetLength(Result, SourceLength);
    end;

    lbsLF:
    begin
      ReplaceStr := LF+#0;
      ReplaceStrLen := 1;
      SetLength(Result, SourceLength);
    end;

    lbsNoLineBreaks:
    begin
      ReplaceStr := #0+#0;
      ReplaceStrLen := 0;
      SetLength(Result, SourceLength-CountCRLF(S));
      //↑改行コードを削除するので
      //  CRLFの分、文字列長が短くなるのでその分を引いて
      //  メモリを確保している
    end;
  else
    ReplaceStrLen := 0;
    Assert(False, '');
  end;

  PResultStr := @Result[1];
  PSourceStr := @S[1];

  ReadIndex := 1;
  WriteIndex := 1;
  while (ReadIndex <= SourceLength-1) do
  begin
    //↓文字がCRかLFの場合は改行コード挿入処理へ
    //  CRLFかLFCRの場合は読み込み位置を1足している
    case PSourceStr^[ReadIndex] of
      CR: begin
        if PSourceStr^[ReadIndex+1]=LF then
          Inc(ReadIndex);
      end;

      LF: begin
        if PSourceStr^[ReadIndex+1]=CR then
          Inc(ReadIndex);
        //改行コード挿入コードへ
      end;

    //↓文字がCRかLFではない場合は
    //  Source文字をResult文字にコピーしている
    else
      PResultStr^[WriteIndex] := PSourceStr^[ReadIndex];
      Inc(WriteIndex);
      Inc(ReadIndex);
      Continue;
    end; //case

    //改行コード挿入処理
    PResultStr^[WriteIndex  ] := ReplaceStr[0];
    PResultStr^[WriteIndex+1] := ReplaceStr[1];
    inc(WriteIndex, ReplaceStrLen);

    Inc(ReadIndex);
  end; //while

  //↓読み込み位置が文字列の最後のときの場合だけ
  //  ループ外で処理する。
  //  PSourceStr^[ReadIndex+1]を読み込むとエラーになるが
  //  その防止処理をループ毎で判定したくないため。
  if ReadIndex = SourceLength then
  begin
    case PSourceStr^[SourceLength] of
      //↓文字がCRかLFの場合は改行コード挿入処理へ
      CR, LF: begin
        //改行コード挿入処理
        PResultStr^[WriteIndex  ] := ReplaceStr[0];
        PResultStr^[WriteIndex+1] := ReplaceStr[1];
        Inc(WriteIndex, ReplaceStrLen);
      end;
    //↓文字がCRかLFではない場合は
    //  Source文字をResult文字にコピーしている
    else
      PResultStr^[WriteIndex] := PSourceStr^[SourceLength];
      Inc(WriteIndex);
    end; //case
  end;
  //SourceLength < ReadIndex になっている場合は
  //文字の最後が改行変換されている場合なので
  //別に何もしない

  SetLength(Result,WriteIndex-1);
end;

(*--▽---------------------------▼--
//2009/01/07
//高速さが失われるかもしれないので必ず既存コードは残しておく。
function ChangeLineBreakes(const S: String; Style: TTextLineBreakStyleMultiPlatform): String;
type StrArray = array[1..$10] of char;

  //CRとLFがいくつあるかをカウントする
  function CountCRLF(const S: String): Integer;
  var
    i: Integer;
    PS: ^StrArray;
  begin
    Result := 0;
    PS := @S[1];
    for i := 1 to Length(S) do
    begin
      case PS^[i] of
        CR: begin
            Inc(Result);
        end;
        LF: begin
            Inc(Result);
        end;
      else ;
      end;
    end; //for i
  end;

var
  ReadIndex, WriteIndex, SourceLength: Integer;
  ReplaceChar: array[0..1] of Char;
  ReplaceStrLen: Integer;
  PResultStr, PSourceStr: ^StrArray;
begin
  if S = '' then begin Result := ''; Exit; end;

  SourceLength := Length(S);
  case Style of
    tlbsmpCRLF:
    begin
      ReplaceChar := CR+LF;
      ReplaceStrLen := 2;
      SetLength(Result, SourceLength+CountCRLF(S));
      //↑最大で改行コードの数だけ
      //  文字列長が増加する可能性があるので
      //  メモリを確保している
    end;

    tlbsmpCR:
    begin
      ReplaceChar := CR+#0;
      ReplaceStrLen := 1;
      SetLength(Result, SourceLength);
    end;

    tlbsmpLF:
    begin
      ReplaceChar := LF+#0;
      ReplaceStrLen := 1;
      SetLength(Result, SourceLength);
    end;
  else
    ReplaceStrLen := 0;
    Assert(False, '');
  end;

  PResultStr := @Result[1];
  PSourceStr := @S[1];

  ReadIndex := 1;
  WriteIndex := 1;
  while (ReadIndex <= SourceLength-1) do
  begin
    case PSourceStr^[ReadIndex] of
      CR: begin
        if PSourceStr^[ReadIndex+1]=LF then
          Inc(ReadIndex);
        //改行コード挿入コードへ
      end;

      LF: begin
        if PSourceStr^[ReadIndex+1]=CR then
          Inc(ReadIndex);
        //改行コード挿入コードへ
      end;

    else
      PResultStr^[WriteIndex] := PSourceStr^[ReadIndex];
      Inc(WriteIndex);
      Inc(ReadIndex);
      Continue;
    end; //case

    //改行コード挿入
    PResultStr^[WriteIndex  ] := ReplaceChar[0];
    PResultStr^[WriteIndex+1] := ReplaceChar[1];
    inc(WriteIndex, ReplaceStrLen);

    Inc(ReadIndex);
  end; //while

  if ReadIndex = SourceLength then
  begin
    case PSourceStr^[SourceLength] of
      CR, LF: begin
        //改行コード挿入
        PResultStr^[WriteIndex  ] := ReplaceChar[0];
        PResultStr^[WriteIndex+1] := ReplaceChar[1];
        Inc(WriteIndex, ReplaceStrLen);
      end;
    else
      PResultStr^[WriteIndex] := PSourceStr^[SourceLength];
      Inc(WriteIndex);
    end; //case
  end;
  //SourceLength < ReadIndex の場合
  //=文字の最後が???CRLFか???LFCRの場合
  //何もしない

  SetLength(Result,WriteIndex-1);
end;
//--▲---------------------------△--*)

{-------------------------------
//  文字列の最後の改行コードを調べる関数
戻り値:     type TLineBreakStyle = (lbsCRLF, lbsCR, lbsLF ,lbsNoLineBreaks);
            で定義される値が戻る
備考:
履歴:       2003/09/15 作成
//--▼----------------------▽--}
function LastLineBreakStyle(S: WideString): TLineBreakStyle;
begin
  Result := lbsNoLineBreaks;

  if 2 <= Length(S) then
  begin
    if (S[Length(S)-1]=#13)
      and (S[Length(S)]=#10) then
    begin
      Result := lbsCRLF;
      Exit;
    end;
  end;

  if 1 <= Length(S) then
  begin
    case S[Length(S)] of
      #13: begin Result := lbsCR; Exit; end;
      #10: begin Result := lbsLF; Exit; end;
    end;
  end;
end;
//--△----------------------▲--

{---------------------------------------
    文字列の最後の改行コードを取り除く手続き
機能:   
備考:   
履歴:   2009/01/15 作成
		2011/10/05(水)
        ・  UnicodeString対応しました。
}//(*-----------------------------------
procedure ExcludeLineBreakProc(var S: String);
begin
  DeleteEndCount(S, Length(LineBreakString(LastLineBreakStyle(S))))
end;

function ExcludeLineBreak(const S: String): String;
begin
  Result := S;
  ExcludeLineBreakProc(Result);
end;
//------------------------------------*)


{-------------------------------
//  改行スタイルから文字に変換する関数
備考:
履歴:       2009/01/15 作成
//--▼----------------------▽--}
function LineBreakString(Style: TLineBreakStyle): String;
begin
  case Style of
    lbsCRLF:    Result := CRLF;
    lbsCR:      Result := CR;
    lbsLF:      Result := LF;
    lbsNoLineBreaks: Result := EmptyStr;
  end;
end;
//--△----------------------▲--

{-------------------------------
//  文字が改行コードかどうか調べる関数
戻り値:     #13#10/#13/#10、それぞれを調べることができる。
備考:
履歴:       2007/06/18 作成
//--▼----------------------▽--}
function LineBreakStyle(S: WideString): TLineBreakStyle;
begin
  Result := lbsNoLineBreaks;

  if 2 = Length(S) then
  begin
    if (S[Length(S)-1]=#13)
      and (S[Length(S)]=#10) then
    begin
      Result := lbsCRLF;
      Exit;
    end;
  end;

  if 1 = Length(S) then
  begin
    case S[Length(S)] of
      #13: begin Result := lbsCR; Exit; end;
      #10: begin Result := lbsLF; Exit; end;
    end;
  end;
end;
//--△----------------------▲--

{---------------------------------------
    終端改行コードの前に文字列を挿入する関数
機能:
備考:
履歴:   2009/01/09
        ・IndentChangeからWideStrRecListUnitへ移動
        2010/10/29
        ・WideStrRecListUnitから移動
}//(*-----------------------------------
function AddStringPreLineBreak(const LineStr, AddStr: WideString): WideString;
var
  LineBreak: String;
  i: Integer;
  WriteIndex: Integer;
begin
  {↓改行コード文字を調べる}
  LineBreak := LineBreakString(LastLineBreakStyle(LineStr));
  {↓改行コードのない文字列に対しても正しく動作する}

  SetLength(Result, Length(LineStr)+Length(AddStr));
  WriteIndex := 1;
  for i:=1 to Length(LineStr)-Length(LineBreak) do
  begin
    Result[WriteIndex] := LineStr[i];
    Inc(WriteIndex);
  end;
  for i:=1 to Length(AddStr) do
  begin
    Result[WriteIndex] := AddStr[i];
    Inc(WriteIndex);
  end;
  for i:=1 to Length(LineBreak) do
  begin
    Result[WriteIndex] := WideChar(LineBreak[i]);
    Inc(WriteIndex);
  end;
end;
//------------------------------------*)

{-------------------------------
//  文字列を一気に置き換えます
機能:       StringReplaceと似ている機能だけど
            '\\'→'\'
            '\r'→#13
            等のように複数文字列を一気に置き換えるので
            結果として'\\r'→'\r'になります。
引数説明:   S: 変更したい文字列
            OldPatterns, NewPatterns: 変更文字列配列
            二つの文字列配列は大きさは同じにしておく必要がある
            IgnoreCase:大小文字を無視する
            ReplaceAll:文字列を全て置き換える
              Falseの場合、どのPatternでも
              1度置き換えたらあとは処理しない。
              少し特殊な使い方になる。
戻り値:     変更後の文字列
備考:
履歴:       2003/03/13
            2007/09/02
              OldPatternsの指定がなかったり
              NewとOldの個数があっていない場合
              ResultにEmptyStrではなくSを返すようにした
            2010/03/10(水)
            ・ReplaceAllFlagをつけた
//--▼----------------------▽--}

function StringsReplace(const S: String; OldPatterns, NewPatterns: array of String;
 CaseCompare: TCaseCompare = ccCaseSensitive; ReplaceAll: Boolean = True): String;

    function PlusValue(Value: Integer): Integer;
    begin
      if 0 < Value then
        Result := Value
      else
        Result := 0;
    end;

var
  ReadIndex, WriteIndex, SourceLength: Integer;
  ResultLength: Integer;
  I, J: Integer;
  ReplaceExecutedFlag: Boolean;
//  CompareFunction: TCompareFunction;
  ComparePatternFlag: Boolean;
begin
  Result := S;
  if Length(OldPatterns) = 0 then Exit;
  if ( Length(OldPatterns) <> Length(NewPatterns) ) then Exit;
  if S = EmptyStr then Exit;

  SourceLength := Length(S);

  {↓文字列長が伸びる事を予測して
     あらかじめ大きめのサイズを取得するための処理}
  ResultLength := SourceLength;
  for I := 0 to Length(OldPatterns)-1 do
  begin
    if OldPatterns[I] <> EmptyStr then
    begin
      ResultLength := ResultLength +
        PlusValue(Length(NewPatterns[I])-Length(OldPatterns[I]))
          * StringCount(OldPatterns[I], S);
    end;
  end;
  SetLength(Result, ResultLength);

  {↓置き換えが実行されたかどうか確認するフラグ}
  ReplaceExecutedFlag := False;
  {↓文字列比較を行うかどうか決めるフラグ}
  ComparePatternFlag := True;

  ReadIndex := 1;
  WriteIndex := 1;
  while (ReadIndex <= SourceLength) do
  begin
    if ComparePatternFlag then
      for I := 0 to Length(OldPatterns)-1 do
      begin
        if StringPartsCompare(OldPatterns[I], S, ReadIndex, CaseCompare) then
        begin
          ReplaceExecutedFlag := True;
          for J := 0 to Length(NewPatterns[I])-1 do
          begin
            Result[WriteIndex+J] := (NewPatterns[I])[J+1];
          end;
          Inc(WriteIndex, Length(NewPatterns[I]));
          Inc(ReadIndex, Length(OldPatterns[I]));
          break;
        end; //if
      end; //for

    if ReplaceExecutedFlag then
    begin
      {↓置き換え実行が1回だけ実行指定なら}
      if not (ReplaceAll) then
      begin
        {↓文字列比較を行わない指定をしてループを抜ける}
        ComparePatternFlag := False;
      end;
      ReplaceExecutedFlag := False;
      continue;
    end else
    begin
      Result[WriteIndex] := S[ReadIndex];
      Inc(WriteIndex);
      Inc(ReadIndex);
    end;
  end; //while
  SetLength(Result,WriteIndex-1);
end;




{-------------------------------
//  エスケープシーケンス変換関数
    EncodeEscapeSequence    CRLF→\r\nにする
    DecodeEscapeSequence    \r\n→CRLFにする
機能:       エスケープシーケンスと
            普通の文字列を相互変換します
引数説明:   Source: 元の文字列
戻り値:     変換後の文字列
備考:       \r/\n/\t
            に対応している
履歴:       2003/03/13
            2003/06/15
              WideStringでの実装に変更
            2004/04/13
              NULL文字を\0に置き換える処理を追加
            2011/05/24(火)
            ・  UnicodeString対応
            ・  内部コードの改善で行数を削減
            2011/12/21(水)
            ・  \0をヌル文字(#0)で置き換えていたので
                \を空文字で置き換えるように変更
//--▼----------------------▽--}
const
  ConvertTblEncodeEscapeSequence: array[0..4] of String =
       ('\r', '\n', '\t', '\\', '\');
  ConvertTblDecodeEscapeSequence: array[0..4] of String =
       ( CR,  LF,   TAB,  '\',  EmptyStr);

function EncodeEscapeSequence(const Source: String): String;
begin
  Result := StringsReplace(Source,
    ConvertTblDecodeEscapeSequence, ConvertTblEncodeEscapeSequence, ccIgnoreCase);
end;

function DecodeEscapeSequence(const Source: String): String;
begin
  Result := StringsReplace(Source,
    ConvertTblEncodeEscapeSequence, ConvertTblDecodeEscapeSequence, ccIgnoreCase);
end;
//--△----------------------▲--

{----------------------------------------
//      Unicode文字の16進数表示
機能:
備考:
履歴:       2010/03/29(月)
            ・作成
//----------------------------------------}
function EncodeStringToUTF16CodeCsv(const Source: WideString;
 Endian: TEndian = eLittle): String;
var
  i: Integer;
  HighByte, LowByte: Byte;
begin
  Result := '';
  for i := 1 to Length(Source) do
  begin
    HighByte := HiByte(word(Source[i]));
    LowByte  := LoByte(word(Source[i]));
    case Endian of
      eBig:
      begin
        Result := Result +
          IntToHex(HighByte, 2) + ',' +
          IntToHex(LowByte, 2) + ',';
      end;
      eLittle:
      begin
        Result := Result +
          IntToHex(LowByte, 2) + ',' +
          IntToHex(HighByte, 2) + ',';
      end;
    end;
  end;
  Result := ExcludeLastStr(Result, ',');
end;

function DecodeStringFromUTF16CodeCsv(Source: String;
 Endian: TEndian = eLittle): WideString;
var
  i, j, LoopCount: Integer;
  Chars: array[0..5] of char;
  CheckFlag: Boolean;
  HighByte, LowByte: Byte;
begin
//・カンマで区切ったものが偶数なこと
//  >>1足して6で割って割り切れること=Loop回数
//・0～Fと[,]で構成されている事
  if (Length(Source) + 1) mod 6 <> 0 then
  begin
    raise EConvertError.Create(
      Format('Error DecodeStringFromCodeCsv String Length %s',
      [Source]));
  end;
  LoopCount := (Length(Source)+1) div 6;
  Source := Source + ',';
  SetLength(Result, LoopCount);
  for i := 0 to LoopCount-1 do
  begin
    CheckFlag := True;
    for j := 0 to 5 do
    begin
      Chars[j] := Source[(i)*6+1+j];
      case j of
        0, 1, 3, 4:
        begin
          if CheckStrInTable(Chars[j], '0123456789abcdefABCDEF')<>itAllInclude then
            CheckFlag := False;
        end;

        2, 5:
        begin
          if Chars[j] <> ',' then
            CheckFlag := False;
        end;
      end;
    end;
    if CheckFlag = False then
    begin
      raise EConvertError.Create(
        Format('Error DecodeStringFromCodeCsv String Format %s',
        [Source]));
    end;

    LowByte := 0; HighByte := 0;
    case Endian of
      eBig:
      begin
        LowByte  := StrToInt('$'+Chars[0]+Chars[1]);
        HighByte := StrToInt('$'+Chars[3]+Chars[4]);
      end;
      eLittle:
      begin
        HighByte := StrToInt('$'+Chars[0]+Chars[1]);
        LowByte  := StrToInt('$'+Chars[3]+Chars[4]);
      end;
    end;
    Result[i+1] := WideChar( (HighByte) + (LowByte)*$100);
  end;
end;

function EncodeStringToUTF8CodeCsv(const Source: WideString): String;
var
  i: Integer;
  CharByte: Byte;
  us: UTF8String;
begin
  Result := '';
  us := UTF8Encode(Source);
  for i := 1 to Length(us) do
  begin
    CharByte := Byte(us[i]);
    Result := Result +
      IntToHex(CharByte, 2) + ',';
  end;
  Result := ExcludeLastStr(Result, ',');
end;

function DecodeStringFromUTF8CodeCsv(Source: String): WideString;
var
  i, j, LoopCount: Integer;
  Chars: array[0..2] of char;
  CheckFlag: Boolean;
  us: UTF8String;
//  HighByte, LowByte: Byte;
begin
//・カンマ区切り
//  >>1足して3で割って割り切れること=Loop回数
//・0～Fと[,]で構成されている事
  if (Length(Source) + 1) mod 3 <> 0 then
  begin
    raise EConvertError.Create(
      Format('Error DecodeStringFromUTF8CodeCsv String Length %s',
      [Source]));
  end;
  LoopCount := (Length(Source)+1) div 3;
  Source := Source + ',';
  SetLength(us, LoopCount);
  for i := 0 to LoopCount-1 do
  begin
    CheckFlag := True;
    for j := 0 to 2 do
    begin
      Chars[j] := Source[(i)*3+1+j];
      case j of
        0, 1:
        begin
          if CheckStrInTable(Chars[j], '0123456789abcdefABCDEF')<>itAllInclude then
            CheckFlag := False;
        end;

        2:
        begin
          if Chars[j] <> ',' then
            CheckFlag := False;
        end;
      end;
    end;
    if CheckFlag = False then
    begin
      raise EConvertError.Create(
        Format('Error DecodeStringFromUTF8CodeCsv String Format %s',
        [Source]));
    end;

    us[i+1] := AnsiChar( Byte(StrToInt('$'+Chars[0]+Chars[1])) );

  end;

//  Result := UTF8Decode(us);
  Result := UTF8ToString(us);
end;

//----------------------------------------


{-------------------------------
//  英単語の置き換えを処理する
機能:       指定された文字列群の単語を置き換える。
            単語の前後をアルファベットかどうかを判断している。
備考:       SQL文中の予約語をUpperCaseに
            そろえる処理の為に実装
            英文系しか実用的ではない。
            WideStringsReplaceから改造
※select→SELECTとすると
SQL文中のSELECTが全部大文字化されるけど
『selecter』という単語は変換されない。
履歴:       2004/01/11
            2007/09/08
              [s"]の["]を[']に変更しようとするときも
              sを前後に接続するアルファベットとして認識してしまうので
              sの隣が記号ならそれはちゃんと区切られているものとして
              認識させることにした。
              つまり、[ttp://]を[http://]に変換するとき
              [xttp://]はtの前にxがきて連続したアルファベットのために
              独立した単語として認識されないので変換されないが
              [ ttp://www]は/の後ろにwがきても連続したアルファベットでないので
              独立した単語として認識されて[ http://www]と変換される
            2011/08/10(水)
            ・TStringDynArrayをarray of Stringに変更した
//--▼----------------------▽--}

function StringsWordReplace(const S: WideString;
 OldPatterns, NewPatterns: array of string;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;

    function PlusValue(Value: Integer): Integer;
    begin
      if 0 < Value then
        Result := Value
      else
        Result := 0;
    end;

type
  TCompareFunction = function(const SubStr, S: WideString; StrIndex: Integer): Boolean;
var
  ReadIndex, WriteIndex, SourceLength: Integer;
  ResultLength: Integer;
  i, j: Integer;
  SameFlag: Boolean;
//  CompareFunction: TCompareFunction;
begin
  Result := '';
  if Length(OldPatterns) = 0 then Exit;
  if ( Length(OldPatterns) <> Length(NewPatterns) ) then Exit;
  if S = '' then Exit;

  SourceLength := Length(S);

  ResultLength := SourceLength;
  for i := 0 to Length(OldPatterns)-1 do
  begin
    if OldPatterns[i] <> '' then
    begin
      ResultLength := ResultLength +
        PlusValue(Length(NewPatterns[i])-Length(OldPatterns[i]))*StringCount(OldPatterns[i], S);
    end;
  end;

  SetLength(Result, ResultLength);


  ReadIndex := 1;
  WriteIndex := 1;
  while (ReadIndex <= SourceLength) do
  begin
    SameFlag := False;
    for i := 0 to Length(OldPatterns)-1 do
    begin
      if StringPartsCompare(OldPatterns[i], S, ReadIndex, CaseCompare) then
      begin
        {↓置き換え元の文字列前後が半角アルファベットかアンダーバーであり
           その隣の文字列も半角アルファベットかアンダーバーであるなら
           単語として独立していないとみなしてパスされる}
        if (1<=ReadIndex) and
          (CheckStrInTable(S[ReadIndex], hanAlphaTbl+'_')=itAllInclude) and
          (CheckStrInTable(S[ReadIndex-1], hanAlphaTbl+'_')=itAllInclude) then
        begin continue; end;

        if (ReadIndex+Length(OldPatterns[i]) <= Length(S)) and
          (CheckStrInTable(S[ReadIndex+Length(OldPatterns[i])-1], hanAlphaTbl+'_') = itAllInclude) and
          (CheckStrInTable(S[ReadIndex+Length(OldPatterns[i])], hanAlphaTbl+'_') = itAllInclude) then
        begin continue; end;

        SameFlag := True;
        for j := 0 to Length(NewPatterns[i])-1 do
        begin
          Result[WriteIndex+j] := (NewPatterns[i])[j+1];
        end;
        Inc(WriteIndex, Length(NewPatterns[i]));
        Inc(ReadIndex, Length(OldPatterns[i]));
        break;
      end; //if
    end; //for

    if SameFlag = False then
    begin
      Result[WriteIndex] := S[ReadIndex];
      Inc(WriteIndex);
      Inc(ReadIndex);
    end;
  end; //while
  SetLength(Result,WriteIndex-1);
end;
//--△----------------------▲--

{---------------------------------------
    指定文字をトリム
        TrimFirstChar
        TrimLastChar
        TrimChar
機能:   S: トリム対象
        Table: トリムしたい文字列テーブル
        SとTableに指定する文字列は
        全角文字やEmptyStrに対応している。
        Tableに'123'と指定すると1,2,3それぞれの文字でトリムする
備考:   
履歴:   2001/10/15
        2003/06/29
            高速化
        2010/10/27
            TrimeLeftCharCountを作成して書き直した
}//(*-----------------------------------


{----------------------------------------
トリムするべき文字がどこまであるかを調べる関数
戻り値は0～Length(S)のどれか
//----------------------------------------}
function TrimFirstCharCount(const S, Table: String): Integer;
var
  I: Integer;
begin
  for I := 1 to Length(S) do
  begin
    if Pos(S[I], Table) = 0 then
    begin
      Result := I - 1;
      Exit;
    end;
  end;
  Result := Length(S);
end;

function TrimFirstChar(const S, Table: String): String;
var
  I: Integer;
  StartIndex: Integer;
begin
  StartIndex := TrimFirstCharCount(S, Table) + 1;
  SetLength(Result, Length(S) - StartIndex + 1);
  for I := StartIndex to Length(S) do
  begin
    Result[I-StartIndex+1] := S[I];
  end;
end;


function TrimLastCharCount(const S, Table: String): Integer;
var
  I: Integer;
begin
  for I := Length(S) downto 1 do
  begin
    if Pos(S[I], Table) = 0 then
    begin
      Result := Length(S) - I;
      Exit;
    end;
  end;
  Result := Length(S);
end;

function TrimLastChar(const S, Table: String): String;
var
  I: Integer;
  EndIndex: Integer;
begin
  EndIndex := Length(S) - TrimLastCharCount(S, Table);
  SetLength(Result, EndIndex);
  for I := 1 to EndIndex do
  begin
    Result[I] := S[I];
  end;
end;

function TrimChar(const S, Table: String): String;
begin
  Result := TrimFirstChar(TrimLastChar(S, Table), Table);
end;
//------------------------------------*)

{---------------------------------------
    Trimする関数
機能:   SysUtilsのTrimLeft/TrimRightから実装をコピーして
		名前変更
備考:   
履歴:   2011/04/08(金)
        ・  作成、UnicodeString対応
		2011/12/17(土)
		・	実装をコピー
}//(*-----------------------------------
function TrimFirst(const S: String): String;
var
  I, L: Integer;
begin
  L := Length(S);
  I := 1;
  while (I <= L) and (S[I] <= ' ') do Inc(I);
  Result := Copy(S, I, Maxint);
end;

function TrimLast(const S: String): String;
var
  I: Integer;
begin
  I := Length(S);
  while (I > 0) and (S[I] <= ' ') do Dec(I);
  Result := Copy(S, 1, I);
end;
//------------------------------------*)


{-------------------------------
//  1文字だけトリムする関数
機能:
備考:       ※この関数なんのために使うんだ？
履歴:       2006/08/02(水) 12:50
//--▼----------------------▽--}
function OneTrimChar(S: String; Table: String): String;
begin
  Result := S;
  if (2 <= Length(S)) = False then Exit;
  if (Table = EmptyStr) then Exit;

  if Pos(S[1], Table)=0 then Exit;

  {↓前後の1文字がTableの中に含まれるなら}
  if Pos(S[1], Table) = Pos(S[Length(S)], Table) then
  begin
    Result := Copy(S, 2, Length(S)-2);
  end;
end;
//--△----------------------▲--

{-------------------------------
//  指定されたタグではさまれた文字列を取り出す
機能:
備考:
履歴:       2006/05/05(金) 15:02
            バグがあったので修正した
//--▼----------------------▽--}
function GetTagInfo(StartTag, EndTag, TargetText: WideString;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): TTagTextInfo;
begin
  Result.StartIndex := PosForward(StartTag, TargetText, 1, MaxInt, CaseCompare);
  Result.EndIndex := PosForward(EndTag, TargetText, Result.StartIndex + 1, MaxInt, CaseCompare);
end;

function GetTagText(StartTag, EndTag, TargetText: WideString;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): WideString;
var
  TagTextInfo: TTagTextInfo;
begin
  TagTextInfo := GetTagInfo(StartTag, EndTag, TargetText);
  if TagTextInfo.StartIndex = 0 then
  begin
    Result := '';
    Exit;
  end;

  if TagTextInfo.EndIndex = 0 then
  begin
    TagTextInfo.EndIndex := Length(TargetText);
  end;

  Result := CopyIndex(TargetText, TagTextInfo.StartIndex, TagTextInfo.EndIndex);
end;
//--△----------------------▲--

{-------------------------------
//  文字列を繰り返し出力します
備考:       StringOfCharの文字列版
履歴:       2004/09/19
            2004/11/30
              DupeString 関数がVCLにあるが
              WideString対応じゃない
//--▼----------------------▽--}
function StringOfStr(Str: String; Count: Integer): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Count do
  begin
    Result := Result + Str;
  end;
end;
//--△----------------------▲--

{-------------------------------
//  WideCharの描画文字幅が半角なのか全角なのかを返す関数
備考:       内部的に使うだけなのでtypeは外部に公開しない
履歴:       2006/04/22
              WordWrap処理のために作成
            2006/05/13
            ・半角カナに対応していなかったので
              HiByte=$00以外も比較条件にいれた
参考:   Unicode対応 文字コード表
        http://ash.jp/code/unitbl1.htm
//--▼----------------------▽--}

function WideCharByteLength(source: WideChar): TCharByteLength;
var
  Buffer: word;
begin
  Buffer := Ord(source);
  case Buffer of
    $0020..$007F, $203E, $FF61..$FF9F:
    begin
      Result := wcblSingle;
    end;
    else
    begin
      Result := wcblMulti;
    end;
  end;
end;

//--△----------------------▲--

{-------------------------------
//  文字位置からByte位置を返す関数
//  CharToByteInexのWideString版(実装には参考にしてない)
備考:   ByteIndexはAnsiStringと同く1オリジンとする
        ByteIndexは描画位置を示して半角80桁とかに使える値になる
履歴:   2006/04/22
            WordWrap処理のために作成
//--▼----------------------▽--}
function CharIndexToByteIndex(Source: WideString; CharIndex: Integer): Integer;
var
  i: Integer;
begin
  Result := 0;
  if not(1<=CharIndex) then Exit;
  if not(CharIndex<=Length(Source)+1) then Exit;

  Result := 1;
  for i := 1 to CharIndex - 1 do
  begin
    case WideCharByteLength(Source[i]) of
      wcblSingle: Inc(Result);
      wcblMulti:  Inc(Result, 2);
    end;
  end;
end;

//--△----------------------▲--

{-------------------------------
//  Byte位置から文字位置を返す関数
//  ByteToCharInexのWideString版(実装には参考にしてない)
備考:   ByteIndexはAnsiStringと同く1オリジンとする
        ByteIndexは描画位置を示して半角80桁とかに使える値になる
履歴:   2006/04/22
              WordWrap処理のために作成
//--▼----------------------▽--}
function ByteIndexToCharIndex(Source: WideString; ByteIndex: Integer): Integer;
var
  i: Integer;
  ByteCounter: Integer;
begin
  Result := 0;
  if not(1<=ByteIndex) then Exit;

  ByteCounter := 0;
  for i := 1 to Length(Source) do
  begin
    case WideCharByteLength(Source[i]) of
      wcblSingle: Inc(ByteCounter);
      wcblMulti:  Inc(ByteCounter, 2);
    end;
    if ByteIndex <= ByteCounter then
    begin
      Result := i;
      break;
    end;
  end;
end;

//--△----------------------▲--

{-------------------------------
//  WideStringの描画半角Byte桁を求める関数
備考:   文字列が半角では何桁になるのかを調べる
履歴:   2006/04/22
              WordWrap処理のために作成
//--▼----------------------▽--}
function ByteLength(Source: WideString): Integer;
begin
  Result := CharIndexToByteIndex(Source, Length(Source)+1)-1;
end;

//--△----------------------▲--

{-------------------------------
//  テーブルに含まれる文字列を削除する
機能:       Tableに"123"と指定すると
            文字列から1と2と3を削除する
備考:
履歴:       2006/11/14(火) 13:54
//--▼----------------------▽--}
function DeleteStrInTable(const S, Table: String;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): String;
var
  OldPatterns, NewPatterns: array of String;
  i: Integer;
begin
  SetLength(OldPatterns, Length(Table));
  SetLength(NewPatterns, Length(Table));
  for i := 0 to Length(Table) - 1 do
  begin
    OldPatterns[i] := Table[i+1];
    NewPatterns[i] := '';
  end;

  Result := StringsReplace(S, OldPatterns, NewPatterns, CaseCompare);
end;
//--△----------------------▲--


{-------------------------------
//  文字列の先頭(First)と終端(Last)に文字列を含ませたり取り除く関数
機能:       
備考:       IncludeBothEndsStrは
            AnsiQuotedStrと同じような仕様かもしれない
履歴:       2007/07/25(水) 17:14
            2010/03/09(火)
            ・  IsFirstStr/IsLastStrとして実装
            2011/05/09(月)
            ・  WideString>>StringにしてUnicodeString対応しました
//--▼----------------------▽--}
function IsFirstStr(const S, SubStr: String; CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): Boolean;
begin
  Result := StringPartsCompare(SubStr, S, 1, CaseCompare);
end;

function IsLastStr(const S, SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): Boolean;
begin
  Result := StringPartsCompare(SubStr, S, Length(S)-Length(SubStr)+1, CaseCompare);
end;

function IncludeFirstStr(const S, SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;
begin
  if IsFirstStr(S, SubStr, CaseCompare) then
  begin
    Result := S;
  end else
  begin
    Result := SubStr + S;
  end;
end;

function IncludeLastStr(const S, SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;
begin
  if IsLastStr(S, SubStr, CaseCompare) then
  begin
    Result := S;
  end else
  begin
    Result := S + SubStr;
  end;
end;

function IncludeBothEndsStr(const S, SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;
begin
  if S = '' then
  begin
    Result := SubStr + SubStr;
  end else
  begin
    Result := IncludeLastStr(IncludeFirstStr(S, SubStr, CaseCompare), SubStr, CaseCompare);
  end;
end;


function ExcludeFirstStr(const S, SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;
begin
  if IsFirstStr(S, SubStr, CaseCompare) then
  begin
    Result := DeleteStr(S, 1, Length(SubStr));
  end else
  begin
    Result := S;
  end;
end;

function ExcludeLastStr(const S, SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;
begin
  if IsLastStr(S, SubStr, CaseCompare) then
  begin
    Result := DeleteStr(S, Length(S)-Length(SubStr)+1, Length(SubStr));
  end else
  begin
    Result := S;
  end;
end;

function ExcludeBothEndsStr(const S, SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): String;
begin
  Result := ExcludeLastStr(ExcludeFirstStr(S, SubStr, CaseCompare), SubStr, CaseCompare);
end;
//--△----------------------▲--

{-------------------------------
//  文字列終端に\記号が追加されたり取り除かれたりする関数
機能:       Include/ExcludeTrailingPathDelimiter()互換(たぶん)
            関数名が長すぎて嫌だからこっちを使う事にした。
備考:
履歴:       2007/07/25(水) 17:14
//--▼----------------------▽--}
function IncludeLastPathDelim(const Path: String): String;
begin
  Result := IncludeLastStr(Path, PathDelim);
end;

function ExcludeLastPathDelim(const Path: String): String;
begin
  Result := ExcludeLastStr(Path, PathDelim);
end;
//--△----------------------▲--

{-------------------------------
//  区切り文字取得
機能:       AAA=BBB=CCCという文字から
            区切り文字[=]を指定すると
            [AAA]や[CCC]や[AAA=BBB]や[BBB=CCC]を取得できる関数
備考:
履歴:       2000/07/24
            2008/03/04(火)
              StringUnitに持ってきた
            2011/12/05(月)
              名称をDelimiterLeft/Rightから
              First/LastStringに変更した
//--▼----------------------▽--}
function FirstString(s: String; Delimiter: String): String;
var
  DeleteIndex: Integer;
begin
  Result := s;
  DeleteIndex := PosForward( Delimiter, s);
  if DeleteIndex = 0 then
    exit;

//  Result := copy( Result, 1, DeleteIndex - 1);
  Delete(Result, DeleteIndex, MaxInt);
end;

function LastString(s: String; Delimiter: String): String;
var
  DeleteIndex: Integer;
begin
  Result := s;
  DeleteIndex := PosBackward(Delimiter, s);
  if DeleteIndex = 0 then
    exit;

  DeleteIndex := DeleteIndex + Length(Delimiter) - 1;
  Delete( Result, 1, DeleteIndex);
end;

//最も後方に位置する区切り文字（文字列）で指定した物の前方方文字列を取得
function FirstStringLong(s: String; Delimiter: String): String;
var
  DeleteIndex: Integer;
begin
  Result := s;
  DeleteIndex := PosBackward(Delimiter, s);
  if DeleteIndex = 0 then
    exit;

//  Result := copy(Result, 1, DeleteIndex - 1);
  Delete(Result, DeleteIndex, MaxInt);
end;

//最も前方に位置する区切り文字（文字列）で指定した物の後方文字列を取得
function LastStringLong(s: String; Delimiter: String): String;
var
  DeleteIndex: Integer;
begin
  Result := s;
  DeleteIndex := PosForward( Delimiter, s);
  if DeleteIndex = 0 then
    exit;

  Delete( Result, 1, DeleteIndex + Length(Delimiter) - 1);
end;

//最も後方に位置する区切り文字（文字列）で指定した物の後方文字列を取得
function DelimiterRight(Delimiter, Str: String): String;
begin
  Result := LastString(Str, Delimiter);
end;
//
//最も前方に位置する区切り文字（文字列）で指定した物の前方文字列を取得
function DelimiterLeft(Delimiter, Str: String): String;
begin
  Result := FirstString(Str, Delimiter)
end;
//
//最も後方に位置する区切り文字（文字列）で指定した物の前方方文字列を取得
function DelimiterLeftLong(Delimiter, Str: String): String;
begin
  Result := FirstStringLong(Str, Delimiter);
end;
//
//最も前方に位置する区切り文字（文字列）で指定した物の後方文字列を取得
function DelimiterRightLong(Delimiter, Str: String): String;
begin
  Result := LastStringLong(Str, Delimiter);
end;
//--△----------------------▲--


{-------------------------------
//  TabToSpace
機能:       タブをスペースに置き換える関数
引数説明:   S: 変換する文字列
戻り値:     変換後の文字列
備考:       タブは4スペースに変換されるとしている
            Sに渡すのは行文字列
            先頭がIndex=0じゃないとタブカウントが変になるから
履歴:       2002/08/31
            ・EmEditorPlugin SpaceReplaceで
              String版とWideString版が実装されたらしい
            2010/10/29
            ・WideStrRecListUnit.pasからStringUnit.pasに移動
//--▼----------------------▽--}
function TabToSpaceLine(const S: WideString): WideString;
var
  ReadIndex: Integer;
  OutputIndex: Integer;
  SourceLen: Integer;
  TabReplaceSpaceCount: Integer;
  i: Integer;
const
  TabToSpaceOptionCount = 4;
begin
  Result := '';
  if S = '' then Exit;

  SourceLen := Length(S);
  SetLength(Result, SourceLen * 4);

  ReadIndex := 1;
  OutputIndex := 1;

  while (ReadIndex <= SourceLen) do
  begin
    if (S[ReadIndex] = TAB) then
    begin
      TabReplaceSpaceCount := TabToSpaceOptionCount - ( CharToByteLen(Result, (OutputIndex-1)) mod TabToSpaceOptionCount );
      for i:=0 to TabReplaceSpaceCount-1 do
      begin
        Result[OutputIndex+i] := ' ';
      end;
      Inc(OutputIndex, TabReplaceSpaceCount);
      Inc(ReadIndex);
    end else
    begin
      Result[OutputIndex] := S[ReadIndex];
      Inc(OutputIndex);
      Inc(ReadIndex);
    end;
  end; //while
  SetLength(Result, OutputIndex-1);
end;
//--△----------------------▲--


{---------------------------------------
    文字に区切り文字指定で文字を追加する
機能:
備考:
履歴:   2011/05/10(火)
        ・  作成
        2011/05/26(木)
        ・  GetStringAdd/GetStringLineAddは
            GetStringConnectで代替できるので削除した
}//(*-----------------------------------
procedure StringAdd(var SourceText: String; AddText, ConnectString: String);
begin
  if SourceText = EmptyStr then
  begin
    SourceText := AddText;
  end else
  begin
    SourceText := IncludeLastStr(SourceText, ConnectString, ccIgnoreCase) + AddText;
  end;
end;

procedure StringLineAdd(var SourceText: String; AddText: String);
begin
  StringAdd(SourceText, AddText, CRLF);
end;
//------------------------------------*)


{---------------------------------------
    文字列を連結する
機能:   
備考:   
履歴:   2011/05/18(水)
        ・  作成
}//(*-----------------------------------
function GetStringConnect(ConnectText: String; SourceTests: array of String): String;
var
  I: Integer;
begin
  Result := '';

  if 1 <= Length(SourceTests) then
  begin
    if SourceTests[0] = '' then
      Result := ''
    else
      Result := SourceTests[0] + ConnectText;
  end;
  for I := 1 to Length(SourceTests) - 1 do
  begin
    Result := Result + SourceTests[I] + ConnectText;
  end;
  Result := ExcludeLastStr(Result, ConnectText);
end;

//改行で連結する
function GetStringLineConnect(SourceTests: array of String): String;
begin
  Result := GetStringConnect(CRLF, SourceTests);
end;
//------------------------------------*)


end.
