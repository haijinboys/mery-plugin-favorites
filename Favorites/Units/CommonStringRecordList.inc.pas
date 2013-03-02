{ -----------------------------------
2003/06/15
  SetTextで文字列を分解するときの改行コード分割処理を
  Winの#13#10だけで対応していたものを
  #13、#10単独での場合も動作するように改造
2010/02/23(火)
・SetTextの実装をDelimiter指定でも
  出来るようにSetBaseTextを実装した。
2010/03/04(木)
・SetBaseTextにSplitFlagsを付属して
  機能を強化した。
・SplitFlagsではなく、dtfIncludeDelimiterのON/OFFにした
2010/03/05(金)
・DelimitStyleとして5タイプに分類して実装
//----------------------------------- }
{$ifdef interface}
{$undef interface}

  TDelimitStyle = (dsLineBreaks,
    dsNoDelimInEmpty, dsNoDelimNoEmpty, dsInDelimInEmpty, dsInDelimNoEmpty);
    {↑ SetBaseTextが使うときに動作する分解方法
        dsLineBreaks
            SetTextを使うときに使用するやり方
            SetBaseText('ABC\r\n\r\nDEF\r\n', [\r\n], dsLineBreaks)とすると
            [ABC\r\n][\r\n][DEF\r\n]となる
        dsNoDelimInEmpty
            区切り文字なしで空文字を含む処理
            SetBaseText('ABC,,DEF,', ',' dsNoDelimInEmpty)とすると
            [ABC][][DEF][]と分解される
        dsNoDelimNoEmpty
            区切り文字なしで空文字なしの処理
            SetBaseText('ABC,,DEF,', ',' dsNoDelimNoEmpty)とすると
            [ABC][DEF]と分解される
        dsInDelimInEmpty
            区切り文字ありで空文字を含む処理
            SetBaseText('ABC,,DEF,', ',' dsInDelimInEmpty)とすると
            [ABC][,][][,][DEF][,][]と分解される
        dsInDelimNoEmpty
            区切り文字ありで空文字なしの処理
            SetBaseText('ABC,,DEF,', ',' dsNoDelimNoEmpty)とすると
            [ABC][,][,][DEF][,]と分解される
    }

  TCommonStringRecordList = class(TRecordList)
  protected
    procedure SetText(const Value: TRecord);
    function GetText: TRecord;
    procedure SetBaseText(const Value: TRecord;
     Delimiters: array of TRecord; const Style: TDelimitStyle);
  public
    procedure SetDelimitedText(const Value: TRecord;
     Delimiters: array of TRecord);
    function GetDelimitedText(Delimiter: TRecord): String;
    property Text: TRecord read GetText write SetText;
  end;
{$endif}

{$ifdef RecordEqual}
{$undef RecordEqual}
function RecordEqual(const Value1, Value2: TRecord): Boolean;
begin
  if ( TRecord(Value1) = TRecord(Value2) ) then
  begin
    Result := True;
  end else
  begin
    Result := False;
  end;
end;
{$endif}

{$ifdef GetSetText}
{$undef GetSetText}

function TCommonStringRecordList.GetText: TRecord;
var
  StringLength, ResultIndex: Integer;
  i, j: Integer;
begin
  StringLength := 0;
  for i := 0 to Self.Count - 1 do
  begin
    Inc(StringLength, Length(Self.Items[i]));
  end;

  SetLength(Result, StringLength);
  ResultIndex := 1;
  for i := 0 to Self.Count - 1 do
  begin
    for j := 1 to Length(Self.Items[i]) do
    begin
      Result[ResultIndex] := Self.Items[i][j];
      Inc(ResultIndex);
    end;
  end;
end;


procedure TCommonStringRecordList.SetText(const Value: TRecord);
const
  LineBreakStrs: array[0..2] of TRecord = (#13#10, #13, #10);
  {↑#13#10,#13,#10の順番だから、CRLFを見つけて区切り
     CRかLFに一致しない場合にCRかLFで判定して区切る事ができる
	 #13,#10,#13#10などという並びではいけない}
begin
  SetBaseText(Value, LineBreakStrs, dsLineBreaks);
end;

{----------------------------------------
・	ABC\r\n\r\nDEF\r\nを分割した場合
	次のようになる。
	・	LineBreaksStyle
		[ABC\r\n][\r\n][DEF\r\n]となる
        区切り文字(\r\n)を前の要素に追加して分解する

	・	NoDelimInEmpty
		[ABC][][DEF]
        区切り文字(\r\n)は消去して要素を追加。
		空文字の場合も空文字要素が1つある状態になる
        StringList的に使える気がする

	・	NoDelimNoEmpty
		[ABC][DEF]
        区切り文字(\r\n)も空文字も追加されない

	・	InDelimInEmpty
		[ABC][\r\n][][\r\n][DEF][\r\n]
        区切り文字(\r\n)も空文字も分割して要素に登録する
        この処理を行ってからいろいろ処理をしやすいかもしれない

	・	InDelimNoEmpty
		[ABC][\r\n][\r\n][DEF][\r\n]
        区切り文字(\r\n)は分割し、空要素はなしにする処理。
		これは役に立つのだろうか…
//----------------------------------------}
procedure TCommonStringRecordList.SetBaseText(const Value: TRecord;
 Delimiters: array of TRecord; const Style: TDelimitStyle);

    function Delimiters_IndexOf(ValueIndex: Integer): Integer;
    var
      j: Integer;
    begin
      Result := -1;
      for j := Low(Delimiters) to High(Delimiters) do
      begin
        if StringPartsCompare(Delimiters[j], Value, ValueIndex) then
        begin
          Result := j;
          break;
        end;
      end;
    end;

    procedure Add_CopyIndex_NotEmptyStr(ValueStartIndex, ValueEndIndex: Integer;
      ValueLen: Integer);
    begin
      if ValueStartIndex <= ValueEndIndex then
        if (CheckRange(1, ValueStartIndex, ValueLen))
          or (CheckRange(1, ValueEndIndex, ValueLen)) then
        begin
          Self.Add(CopyIndex(Value, ValueStartIndex, ValueEndIndex));
        end;
    end;

var
  i, Len, StartIndex, EndIndex: Integer;
  DelimitersIndex: Integer;
begin
  Self.Clear;
  {↓ds…InEmptyフラグの場合に
     Valueに空文字が与えられたら
     要素数が0で返す場合は以下の処理を実行
     要素数が1で内容が空文字が入るようにするには
     以下の処理を実行しない。}
{----------------------------------------
  if Value = EmptyStr then
  begin
    Exit;
  end;
//----------------------------------------}

  Len := Length(Value);

  StartIndex := 0;
  EndIndex := 0;

  case Style of
    dsLineBreaks:
    begin
      {↓[abc\r\n][\r\n][def\r\n]となる
         区切り文字が見つかった場所=iと
         区切り文字の長さを足して、区切り文字の最後の位置
         ＝EndIndexを求めています}
      i := 1; while (i <= Len) do
      begin
        DelimitersIndex := Delimiters_IndexOf(i);

        if DelimitersIndex <> -1 Then
        begin
          EndIndex := i + Length(Delimiters[DelimitersIndex]) - 1;
          Self.Add(CopyIndex(Value, StartIndex, EndIndex));
          i := EndIndex + 1;
          StartIndex := i;
        end else
        begin
          EndIndex := i;
          Inc(i);
        end;
      end;
      Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
    end;

    dsNoDelimInEmpty:
    begin
      {↓[abc][][def]となる
         区切り文字が見つかった場所=iの
         1文字前が切り取るべき位置
         ＝EndIndexになります}
      i := 1; while (i <= Len) do
      begin
        DelimitersIndex := Delimiters_IndexOf(i);

        if DelimitersIndex <> -1 Then
        begin
          EndIndex := i - 1;
          Self.Add(CopyIndex(Value, StartIndex, EndIndex));
          i := i + Length(Delimiters[DelimitersIndex]);
          StartIndex := i;
        end else
        begin
          EndIndex := i;
          Inc(i);
        end;
      end;
      Self.Add(CopyIndex(Value, StartIndex, EndIndex));
    end;

    dsNoDelimNoEmpty:
    begin
		{↓[ABC][DEF]
		   空文字も区切り文字も追加されない}
      i := 1; while (i <= Len) do
      begin
        DelimitersIndex := Delimiters_IndexOf(i);

        if DelimitersIndex <> -1 Then
        begin
          EndIndex := i - 1;
          Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
          i := i + Length(Delimiters[DelimitersIndex]);
          StartIndex := i;
        end else
        begin
          EndIndex := i;
          Inc(i);
        end;
      end;
      Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
    end;

    dsInDelimInEmpty:
    begin
      {↓[ABC][\r\n][][\r\n][DEF][\r\n]}
      i := 1; while (i <= Len) do
      begin
        DelimitersIndex := Delimiters_IndexOf(i);

        if DelimitersIndex <> -1 Then
        begin
          EndIndex := i - 1;
          Self.Add(CopyIndex(Value, StartIndex, EndIndex));
          Self.Add(Copy(Value, EndIndex+1, Length(Delimiters[DelimitersIndex])));
          i := i + Length(Delimiters[DelimitersIndex]);
          StartIndex := i;
        end else
        begin
          EndIndex := i;
          Inc(i);
        end;
      end;
      Self.Add(CopyIndex(Value, StartIndex, EndIndex));
    end;

    dsInDelimNoEmpty:
    begin
      {↓[ABC][\r\n][\r\n][DEF][\r\n]}
      i := 1; while (i <= Len) do
      begin
        DelimitersIndex := Delimiters_IndexOf(i);

        if DelimitersIndex <> -1 Then
        begin
          EndIndex := i - 1;
          Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
          Self.Add(Copy(Value, EndIndex+1, Length(Delimiters[DelimitersIndex])));
          i := i + Length(Delimiters[DelimitersIndex]);
          StartIndex := i;
        end else
        begin
          EndIndex := i;
          Inc(i);
        end;
      end;
      Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
    end;
  end;

//  i := 1; while (i <= Len) do
//  begin
//    DelimitersIndex := Delimiters_IndexOf(i);
//
//    if DelimitersIndex <> -1 Then
//    begin
//      case Style of
//      dsLineBreaks:
//      begin
//        {↓[abc\r\n][\r\n][def\r\n]となる
//           区切り文字が見つかった場所=iと
//           区切り文字の長さを足して、区切り文字の最後の位置
//           ＝EndIndexを求めています}
//        EndIndex := i + Length(Delimiters[DelimitersIndex]) - 1;
//        Self.Add(CopyIndex(Value, StartIndex, EndIndex));
//        i := EndIndex + 1;
//        StartIndex := i;
//      end;
//
//      dsNoDelimInEmpty:
//      begin
//        {↓[abc][][def]となる
//           区切り文字が見つかった場所=iの
//           1文字前が切り取るべき位置
//           ＝EndIndexになります}
//        EndIndex := i - 1;
//        Self.Add(CopyIndex(Value, StartIndex, EndIndex));
//        i := i + Length(Delimiters[DelimitersIndex]);
//        StartIndex := i;
//      end;
//
//      dsNoDelimNoEmpty:
//      begin
//		{↓[ABC][DEF]
//		   空文字も区切り文字も追加されない}
//        EndIndex := i - 1;
//        Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
//        i := i + Length(Delimiters[DelimitersIndex]);
//        StartIndex := i;
//      end;
//
//      dsInDelimInEmpty:
//      begin
//        {↓[ABC][\r\n][][\r\n][DEF][\r\n]}
//        EndIndex := i - 1;
//        Self.Add(CopyIndex(Value, StartIndex, EndIndex));
//        Self.Add(Copy(Value, EndIndex+1, Length(Delimiters[DelimitersIndex])));
//        i := i + Length(Delimiters[DelimitersIndex]);
//        StartIndex := i;
//      end;
//
//      dsInDelimNoEmpty:
//      begin
//        {↓[ABC][\r\n][\r\n][DEF][\r\n]}
//        EndIndex := i - 1;
//        Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
//        Self.Add(Copy(Value, EndIndex+1, Length(Delimiters[DelimitersIndex])));
//        i := i + Length(Delimiters[DelimitersIndex]);
//        StartIndex := i;
//      end;
//
//      end;
//    end else
//    begin
//      EndIndex := i;
//      Inc(i);
//    end;
//  end;
//
//  {↓最後に残った文字がある場合は取得する
//     最後に区切り文字がある場合はStartIndexが
//     Lengthを超えるので取得できないからこれでOK}
//  case Style of
//
//  dsNoDelimInEmpty, dsInDelimInEmpty:
//  begin
//    Self.Add(CopyIndex(Value, StartIndex, EndIndex));
//  end;
//
//  dsLineBreaks, dsNoDelimNoEmpty, dsInDelimNoEmpty:
//  begin
//    Add_CopyIndex_NotEmptyStr(StartIndex, EndIndex, Len);
//  end;
//
//  end;

end;

{-------------------------------
//      区切り文字を指定して文字列を分割する
機能:
備考:       StringList.DelimitedTextの機能と
            同じような機能として使える
履歴:       2010/02/23(火)
//--▼----------------------▽--}
procedure TCommonStringRecordList.SetDelimitedText(const Value: TRecord;
 Delimiters: array of TRecord);
begin
  SetBaseText(Value, Delimiters, dsNoDelimInEmpty);
end;
//--△----------------------▲--


{---------------------------------------
    区切り文字を指定して連結した文字列を取得する
機能:   
備考:   
履歴:   2011/06/13(月)
        ・	作成
            CommaTextの入出力と同じようにする
}//(*-----------------------------------
function TCommonStringRecordList.GetDelimitedText(Delimiter: TRecord): String;
var
  I: Integer;
begin
  Result := '';
  for I := 0 to Self.Count - 1 do
  begin
    Result := Result + Items[I] + Delimiter;
  end;
  Result := ExcludeLastStr(Result, Delimiter);
end;
//------------------------------------*)


{$endif}




