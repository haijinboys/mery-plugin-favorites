(*----------------------------------------
文字列検索用の関数
PosForward/PosBackwardの置き換え用
2011/12/20(火)
・	作成
//----------------------------------------*)
unit StringSearchUnit;

interface

uses
  StringUnit,
  ConstUnit,
uses_end;

type
  TSearchDirection = (sdForwardToBackward, sdBackwardToForward);
  {↑ForwardToBackward:   前方検索    前から後ろへ検索する[→]
     BackwardToForward:   後方検索    後ろから前へ検索する[←]}

  TSearchResult = record
    SearchIndex: Integer;
    SubStrArrayIndex: Integer;
  end;

  TSearchRange = record
  private
    function GetCount: Integer;
  public
    StartIndex: Integer;
    EndIndex: Integer;
    property Count: Integer read GetCount;
  end;

function StartEndIndex(StartIndex, EndIndex: Integer): TSearchRange;
function StartIndexCount(StartIndex, Count: Integer): TSearchRange;
function StartIndexToLast(StartIndex: Integer): TSearchRange;

function Search(Direction: TSearchDirection; S: String; SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
function Search(Direction: TSearchDirection; S: String; SubStr: String;
 Range: TSearchRange; CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
function Search(Direction: TSearchDirection; S: String; SubStrs: array of string;
 CaseCompare: TCaseCompare = ccCaseSensitive): TSearchResult; overload;
function Search(Direction: TSearchDirection; S: String; SubStrs: array of string;
 Range: TSearchRange; CaseCompare: TCaseCompare = ccCaseSensitive): TSearchResult; overload;



{---------------------------------------
    下位互換関数
機能:   
備考:   
履歴:   2011/12/20(火)
        ・  下位互換ということにした
}//(*-----------------------------------
function PosForward(const SubStr, S: String; CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): Integer; overload;
function PosForward(const SubStr, S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
function PosBackward(const Substr, S: String; CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): Integer; overload;
function PosBackward(const Substr, S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;

type
  TPosResult = record
    SearchIndex: Integer;
    SubStrArrayIndex: Integer;
  end;

function PosForward(const SubStrs: array of string; const S: String;
 CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): TPosResult; overload;
function PosForward(const SubStrs: array of string; const S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): TPosResult; overload;
function PosBackward(const SubStrs: array of string; const S: String;
 CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): TPosResult; overload;
function PosBackward(const SubStrs: array of string; const S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): TPosResult; overload;

//下位互換
function RangePosForward(const SubStr, S: String;
 Index: Integer = 1; Count: Integer = MaxInt;
 CaseCompare: TCaseCompare = ccCaseSensitive): Integer;
function RangePosBackward(const SubStr, S: String;
 Index: Integer = 1; Count: Integer = MaxInt;
 CaseCompare: TCaseCompare = ccCaseSensitive): Integer;
//------------------------------------*)


implementation


function TSearchRange.GetCount: Integer;
begin
  Result := EndIndex - StartIndex + 1;
end;

function StartEndIndex(StartIndex, EndIndex: Integer): TSearchRange;
begin
  Result.StartIndex := StartIndex;
  Result.EndIndex := EndIndex;
end;

function StartIndexCount(StartIndex, Count: Integer): TSearchRange;
var
  Buffer: Int64;
begin
  Result.StartIndex := StartIndex;
  Buffer := StartIndex + Int64(Count) - 1;
  if Int64(MaxInt) < Buffer then
    Result.EndIndex := MaxInt
  else
    Result.EndIndex := Buffer;
end;

function StartIndexToLast(StartIndex: Integer): TSearchRange;
begin
  Result.StartIndex := StartIndex;
  Result.EndIndex := MaxInt;
end;

function AllRange: TSearchRange;
begin
  Result.StartIndex := 1;
  Result.EndIndex := MaxInt;
end;

function Search_Base(const S: String; const SubStrs: array of string;
 Range: TSearchRange; CaseCompare: TCaseCompare;
 SearchDirection: TSearchDirection): TSearchResult;

  function IsEmptyAllSubStrs: Boolean;
  var
    I: Integer;
  begin
    Result := True;
    for I := 0 to Length(SubStrs) - 1 do
    begin
      if SubStrs[I] <> EmptyStr then
      begin
        Result := False;
        Break;
      end;
    end;
  end;

var
  I, J: Integer;
begin
  Result.SearchIndex := 0;
  Result.SubStrArrayIndex := -1;

  if (S = EmptyStr) then Exit;
  if (Range.EndIndex <= 0) then Exit;
  if (Length(S) + 1 <= Range.StartIndex) then Exit;
  if (Range.EndIndex < Range.StartIndex) then Exit;

//  if IsEmptyAllSubStrs then Exit;
  {↑おそらく不要}

  { 123456789A
    __________←Length(S) = 10
       4____9 ←StartIndex=4/EndIndex=9 長さ=(9-4+1)=6
       ___    ←SubStr=3Char
        ___
         ___
          ___ ←ループは4から7←(9+1-3)
                逆方向は7から4
  }

  if Range.StartIndex <= 0 then       Range.StartIndex  := 1;
  if Length(S) < Range.EndIndex then  Range.EndIndex    := Length(S);

  case SearchDirection of
    sdForwardToBackward:
    begin
      for I := Range.StartIndex to Range.EndIndex do
        for J := 0 to Length(SubStrs) - 1 do
        begin
          if (Range.EndIndex - I + 1) < Length(SubStrs[J]) then Break;
          {↑EndIndexが短く指定されている場合この処理ではじく}
          if StringPartsCompare(SubStrs[J], S, I, CaseCompare) then
          begin
            Result.SearchIndex := I;
            Result.SubStrArrayIndex := J;
            Exit;
          end;
        end;
    end;

    sdBackwardToForward:
    begin
      for I := Range.EndIndex downto Range.StartIndex do
        for J := 0 to Length(SubStrs) - 1 do
        begin
          if (Range.EndIndex - I + 1) < Length(SubStrs[J]) then Break;

          if StringPartsCompare(SubStrs[J], S, I, CaseCompare) then
          begin
            Result.SearchIndex := I;
            Result.SubStrArrayIndex := J;
            Exit;
          end;
        end;
    end;
  end; //case SearchDirection
end;

//Index指定なしの呼び出し
function Search(Direction: TSearchDirection; S: String; SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
begin
  Result := Search_Base(S, [SubStr], AllRange, CaseCompare, Direction).SearchIndex;
end;

//Index指定ありの呼び出し
function Search(Direction: TSearchDirection; S: String; SubStr: String;
 Range: TSearchRange; CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
begin
  Result := Search_Base(S, [SubStr], Range, CaseCompare, Direction).SearchIndex;
end;

//SubStr複数指定でのIndex指定なしの呼び出し
function Search(Direction: TSearchDirection; S: String; SubStrs: array of string;
 CaseCompare: TCaseCompare = ccCaseSensitive): TSearchResult; overload;
begin
  Result := Search_Base(S, SubStrs, AllRange, CaseCompare, Direction);
end;

//SubStr複数指定でのIndex指定ありの呼び出し
function Search(Direction: TSearchDirection; S: String; SubStrs: array of string;
 Range: TSearchRange; CaseCompare: TCaseCompare = ccCaseSensitive): TSearchResult; overload;
begin
  Result := Search_Base(S, SubStrs, Range, CaseCompare, Direction);
end;

////////////////////////////////////////
//先頭引数省略前方指定
////////////////////////////////////////

//Index指定なしの呼び出し
function Search(S: String; SubStr: String;
 CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
begin
  Result := Search_Base(S, [SubStr], AllRange, CaseCompare, sdForwardToBackward).SearchIndex;
end;

//Index指定ありの呼び出し
function Search(S: String; SubStr: String;
 Range: TSearchRange; CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
begin
  Result := Search_Base(S, [SubStr], Range, CaseCompare, sdForwardToBackward).SearchIndex;
end;

//SubStr複数指定でのIndex指定なしの呼び出し
function Search(S: String; SubStrs: array of string;
 CaseCompare: TCaseCompare = ccCaseSensitive): TSearchResult; overload;
begin
  Result := Search_Base(S, SubStrs, AllRange, CaseCompare, sdForwardToBackward);
end;

//SubStr複数指定でのIndex指定ありの呼び出し
function Search(S: String; SubStrs: array of string;
 Range: TSearchRange; CaseCompare: TCaseCompare = ccCaseSensitive): TSearchResult; overload;
begin
  Result := Search_Base(S, SubStrs, Range, CaseCompare, sdForwardToBackward);
end;


////////////////////////////////////////
//下位互換用
////////////////////////////////////////

{---------------------------------------
    文字列検索 PosForward/PosBackward
機能:   指定範囲を検索するPos
        0:文字列が存在しない
        0以外:検索文字列が存在するSのIndex
        Index/Countで位置を指定する。
        EndIndex= Index + Count - 1 で求める事ができる
備考:   Indexが1～Length(S)以外だと空文字を返す
履歴:   2002/05/31
        2002/07/14
        Count が MaxInt の場合
            EndIndex := Index + Count -1 が
            負の値になるので
            EndIndex:Cardinalで対応をする
        2002/08/18
            StrPartsCompareで使っている
            ByteTypeをVCLのSysUtilsのものを
            使わずにAPIを使うことにした
        2002/08/22
            ByteTypeをようやく正しく実装。
            EndIndex:Cardinalを止めて
            Integerにして範囲を超えた時の負の値
            をはじくようにした
        2002/09/29
            StrPartsCompareをAnsiStringPartsCompare_Base
            という名称に変えて
            IsSameStrをAnsiStringPartsCompareに変更
            RangeAnsiPos/RangeBackAnsiPosと
            実装が同じだったのでまとめた
        2005/09/22
            関数名変えた
履歴:   2011/12/16(金)
        ・  RangePos系の機能をPosForwardに含めたので
            RangePos系は使用しない方向にする
        ・  RangePos_BaseからPos_Baseへ名前変更
}//(*-----------------------------------
type
  TPosSearchDirection = (sdForward, sdBackward);
  {↑Forward:   前方検索    前から後ろへ検索する[→]
     Backward:  後方検索    後ろから前へ検索する[←]}


function Pos_Base(const SubStrs: array of string; const S: String;
 Index, Count: Integer; CaseCompare: TCaseCompare;
 SearchDirection: TPosSearchDirection): TPosResult; overload;
var
  I, J: Integer;
  EndIndex: Integer;
  SubStrsEmptyFlag: Boolean;
begin
  Result.SearchIndex := 0;
  Result.SubStrArrayIndex := -1;

  if (S='') then Exit;
  if not (  (1<=Index) and (Index<=Length(S))  ) then Exit;
  if not (1<=Count) then Exit;

  SubStrsEmptyFlag := True;
  for I := 0 to Length(SubStrs) - 1 do
  begin
    if SubStrs[I] <> EmptyStr then
    begin
      SubStrsEmptyFlag := False;
      Break;
    end;
  end;
  if SubStrsEmptyFlag then Exit;

  {↓Index+Count-1を計算してMaxIntを超える場合
     負の値になるので修正}
  EndIndex := Index + Count - 1;
  if (EndIndex < 0) or (Length(S) < EndIndex) then
  begin
    EndIndex := Length(S);
  end;

  { 123456789A
    __________←Sは10Char
       4____9 ←Index=4/End=9の6Char
       ___    ←SubStr=3Char
        ___
         ___
          ___ ←ループは4から7←(9+1-3)
                逆方向は7から4
  }
  case SearchDirection of
    sdForward:
    begin
      for I := Index to EndIndex do
        for J := 0 to Length(SubStrs) - 1 do
        begin
//          if StringPartsCompare_Base(SubStrs[J], S, 1, I, Min(Length(SubStrs[J]), Count), CaseCompare) then
          if (EndIndex - I + 1) < Length(SubStrs[J]) then Break;

          if StringPartsCompare(SubStrs[J], S, I, CaseCompare) then
          begin
            Result.SearchIndex := I;
            Result.SubStrArrayIndex := J;
            Exit;
          end;
        end;
    end;

    sdBackward:
    begin
      for I := EndIndex downto Index do
        for J := 0 to Length(SubStrs) - 1 do
        begin
//          if StringPartsCompare_Base(SubStrs[J], S, 1, I, Min(Length(SubStrs[J]), Count), CaseCompare) then
          if (EndIndex - I + 1) < Length(SubStrs[J]) then Break;
          if StringPartsCompare(SubStrs[J], S, I, CaseCompare) then
          begin
            Result.SearchIndex := I;
            Result.SubStrArrayIndex := J;
            Exit;
          end;
        end;
    end;
  end; //case SearchDirection
end;

//function Pos_Base(const SubStr, S: String;
// Index, Count: Integer; CaseCompare: TCaseCompare;
// SearchDirection: TSearchDirection): Integer; overload;
//begin
//  Result := Pos_Base([SubStr], S, Index, Count, CaseCompare, SearchDirection).SearchIndex;
//end;

function PosForward(const SubStr, S: String; CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): Integer; overload;
begin
  Result := Pos_Base(SubStr, S, Index, Count, CaseCompare, sdForward).SearchIndex;
end;

function PosForward(const SubStr, S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
begin
  Result := Pos_Base(SubStr, S, Index, Count, CaseCompare, sdForward).SearchIndex;
end;

function PosForward(const SubStrs: array of string; const S: String;
 CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): TPosResult; overload;
begin
  Result := Pos_Base(SubStrs, S, Index, Count, CaseCompare, sdForward);
end;

function PosForward(const SubStrs: array of string; const S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): TPosResult; overload;
begin
  Result := Pos_Base(SubStrs, S, Index, Count, CaseCompare, sdForward);
end;

function PosBackward(const Substr, S: String; CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): Integer; overload;
begin
  Result := Pos_Base(SubStr, S, Index, Count, CaseCompare, sdBackward).SearchIndex;
end;

function PosBackward(const Substr, S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): Integer; overload;
begin
  Result := Pos_Base(SubStr, S, Index, Count, CaseCompare, sdBackward).SearchIndex;
end;

function PosBackward(const SubStrs: array of string; const S: String;
 CaseCompare: TCaseCompare = ccCaseSensitive;
 Index: Integer = 1; Count: Integer = MaxInt): TPosResult; overload;
begin
  Result := Pos_Base(SubStrs, S, Index, Count, CaseCompare, sdBackward);
end;

function PosBackward(const SubStrs: array of string; const S: String;
 Index: Integer; Count: Integer = MaxInt; CaseCompare: TCaseCompare = ccCaseSensitive): TPosResult; overload;
begin
  Result := Pos_Base(SubStrs, S, Index, Count, CaseCompare, sdBackward);
end;

//下位互換関数
//PosForward/PosBackwardで置き換え
function RangePosForward(const SubStr, S: String;
 Index: Integer = 1; Count: Integer = MaxInt;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): Integer;
begin
  Result := Pos_Base(SubStr, S, Index, Count, CaseCompare, sdForward).SearchIndex;
end;

function RangePosBackward(const SubStr, S: String;
 Index: Integer = 1; Count: Integer = MaxInt;
 CaseCompare: TCaseCompare = ccCaseSensitive{IgnoreCase=False}): Integer;
begin
  Result := Pos_Base(SubStr, S, Index, Count, CaseCompare, sdBackward).SearchIndex;
end;
//------------------------------------*)


end.

