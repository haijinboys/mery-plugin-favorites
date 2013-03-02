{ --▽---------------------------▼--
汎用共通処理関数ユニット
2004/08/18
・作成
・RoundOFF関数などを実装
・CeilEx/FloorExも実装
2004/10/02
・RandomRangeExを実装

//--▲---------------------------△-- }
unit MathUnit;

interface

uses
  Math,
uses_end;

function RoundOff(X: Extended): Int64;
function RoundOffEx(const X: Extended; DigitNumber: Integer): Extended;
function IsSame(A, B: Extended; e: Extended=0): Boolean;
function CeilEx(const X: Extended; DigitNumber: Integer): Extended;
function FloorEx(const X: Extended; DigitNumber: Integer): Extended;
function TruncEx(const X: Extended; DigitNumber: Integer): Extended;

function RandomRangeEx(const AFrom, ATo: Extended; Digit: Byte): Extended;

function CheckRange(Min, Value, Max: Integer): Boolean;
function Digit(Value: Cardinal): Integer;

implementation

{-------------------------------
//  HELPにのっているRoundOff
備考:
履歴:       2004/08/18
              LongintからInt64に変更した
//--▼----------------------▽--}
function RoundOff(X: Extended): Int64;
begin
  if x >= 0 then Result := Trunc(x + 0.5)
  else Result := Trunc(x - 0.5);
end;
//--△----------------------▲--

{-------------------------------
//  RoundOffEx
機能:       任意の桁で四捨五入します。
引数説明:   X: 四捨五入対象
            DigitNumber: 桁数
              3:100の桁
              2:10の桁
              1:一桁目
              0:処理しない
              -1:少数点第一位(通常のRoundOffと同じ処理)
              -2:少数点第二位
              -3:少数点第三位
戻り値:     四捨五入後のX
            処理できない場合Xの値そのままが入るはず
備考:
履歴:       2001/09/04
//--▼----------------------▽--}
function RoundOffEx(const X: Extended; DigitNumber: Integer): Extended;
var
  CalcDigit: Extended;
begin
  Result := X;
  if X = 0 then Exit;

  case DigitNumber of
    0: Exit;
    1..High(DigitNumber):
    begin
      CalcDigit := IntPower(10, DigitNumber);
      Result := Roundoff(X / CalcDigit) * CalcDigit;
    end;
    Low(DigitNumber)..-1:
    begin
      CalcDigit := IntPower(10, Abs(DigitNumber)-1);
      Result := Roundoff(X * CalcDigit) / CalcDigit;
    end;
  end;
end;
//--△----------------------▲--

{-------------------------------
//  IsSame
機能:       適当に小さい誤差の範囲内で
            浮動小数点値が等しいかどうかを調べる関数
引数説明:   A, B: 比較対照
            e: ±の誤差
戻り値:     true:等しい false:等しくない
備考:
履歴:       2001/09/05
//--▼----------------------▽--}
function IsSame(A, B: Extended; e: Extended=0): Boolean;
begin
  Result := Abs(A-B) <= e;
end;
//--△----------------------▲--

{---------------------------------------
    CeilEx/FloorEx/TruncEx
機能:   任意の桁で切り上げ/切り下げ/切捨てします
引数:	X: 計算対象
        DigitNumber: 桁数
          3:100の桁を処理、1000の桁で丸め
          2:10の桁を処理、 100の桁で丸め
          1:一桁目を処理、 10の桁を丸め
          0:処理しない
          -1:少数点第一位(通常のRoundOffと同じ処理)
          -2:少数点第二位
          -3:少数点第三位
備考:   切り上げは正の絶対値方向、切り下げは負の無限大方向に
        行われるので、負の値の場合動作に気をつけること。
履歴:   2011/09/12(月)
        ・  TruncExを追加
}//(*-----------------------------------
function CeilEx(const X: Extended; DigitNumber: Integer): Extended;
var
  CalcDigit: Extended;
begin
  Result := X;
  if X = 0 then Exit;

  case DigitNumber of
    0: Exit;
    1..High(DigitNumber):
    begin
      CalcDigit := IntPower(10, DigitNumber);
      Result := Math.Ceil(X / CalcDigit) * CalcDigit;
    end;
    Low(DigitNumber)..-1:
    begin
      CalcDigit := IntPower(10, Abs(DigitNumber)-1);
      Result := Math.Ceil(X * CalcDigit) / CalcDigit;
    end;
  end;
end;

function FloorEx(const X: Extended; DigitNumber: Integer): Extended;
var
  CalcDigit: Extended;
begin
  Result := X;
  if X = 0 then Exit;

  case DigitNumber of
    0: Exit;
    1..High(DigitNumber):
    begin
      CalcDigit := IntPower(10, DigitNumber);
      Result := Math.Floor(X / CalcDigit) * CalcDigit;
    end;
    Low(DigitNumber)..-1:
    begin
      CalcDigit := IntPower(10, Abs(DigitNumber)-1);
      Result := Math.Floor(X * CalcDigit) / CalcDigit;
    end;
  end;
end;

function TruncEx(const X: Extended; DigitNumber: Integer): Extended;
begin
  Result := X;
  if X = 0 then Exit;

  if 0 < X then
  begin
    Result := FloorEx(X, DigitNumber);
  end else
  begin
    Result := CeilEx(X, DigitNumber);
  end;
end;
//------------------------------------*)

{-------------------------------
//  RandomRangeEx
機能:       任意桁の範囲でRandomRangeする関数
引数説明:   AFrom/ATo: 範囲
            Digit: 桁数
              1:小数点1桁
              2:小数点2桁
              3:小数点3桁
              0:処理しない
戻り値:     AFrom/AToで示される範囲の乱数
備考:
履歴:       2004/10/02
//--▼----------------------▽--}
function RandomRangeEx(const AFrom, ATo: Extended; Digit: Byte): Extended;
var
  CalcDigit: Extended;
begin
  CalcDigit := IntPower(10, Digit); //
  Result := RandomRange(RoundOff(AFrom*CalcDigit), RoundOff(ATo*CalcDigit)) / CalcDigit;
end;
//--△----------------------▲--

{-------------------------------
//  数値が範囲内にあるかどうか調べる関数
備考:       Math Unit の InRange 関数と同じ
履歴:       2005/11/23
//--▼----------------------▽--}
function CheckRange(Min, Value, Max: Integer): Boolean;
begin
  if (Min <= Value) and (Value <= Max) then
  begin
    Result := True;
  end else
  begin
    Result := False;
  end;
end;
//--△----------------------▲--

{-------------------------------
//  桁数を求める
機能:   0～正の数の桁数を求めます
戻り値: 桁数
備考:   uses Mathが必要
        99999996～99999999の値が
        case1では誤動作して求まらない
        case2は正確だが計算回数が多くなるので
        case3が単純だが適切
履歴:   2003/03/13
        2011/09/12(月)
        ・  既存のコードをcase1にしてcase3まで作成
//--▼----------------------▽--}
function Digit(Value: Cardinal): Integer;
begin
  case 3 of
    1: 
    begin
      if Value = 0 then
        Result := 1
      else
        Result := Trunc(Log10(Value))+1;
    end;
  
    2: 
    begin
      Result := 1;
      while IntPower(10, Result) <= Value do
      begin
        Inc(Result);
        if High(Cardinal) < IntPower(10, Result) then Break;
      end;
    end;

    3: 
    begin
      Result := 1;
      if Value < 10 then Exit;
      Inc(Result);
      if Value < 100 then Exit;
      Inc(Result);
      if Value < 1000 then Exit;
      Inc(Result);
      if Value < 10000 then Exit;
      Inc(Result);
      if Value < 100000 then Exit;
      Inc(Result);
      if Value < 1000000 then Exit;
      Inc(Result);
      if Value < 10000000 then Exit;
      Inc(Result);
      if Value < 100000000 then Exit;
      Inc(Result);
      if Value < 1000000000 then Exit;
      Inc(Result);
      //High(Cardinal)=2147483647;で最大10桁
 	end;
  end; //case
end;
//--△----------------------▲--


end.

