(*----------------------------------------
文字列変換を行うユニット
2011/12/22(木)
・  作成
//----------------------------------------*)
unit StringConvertUnit;

interface

uses
  Types,      //下位互換関数でのTStringDynArrayの宣言のため
  StringUnit, //下位互換関数でのStringsReplaceの利用のため
  ConstUnit,
uses_end;

type
  TStringConvertOption = (
    scUpperCase, scLowerCase,
    scAlphabetToZenkaku,  scAlphabetToHankaku,
    scSymbolToZenkaku,    scSymbolToHankaku,
    scNumericToZenakaku,  scNumericToHankaku,
    scKatakanaToZenkaku,  scKatakanaToHankaku,
    scKatakanaToHiragana, scHiraganaToKatakana
  );
procedure StringConvert(var S: string; Option: TStringConvertOption); overload;
procedure StringConvert(Option: TStringConvertOption; var S: string); overload;
procedure StringConvert(var S: String;
 const OldPatterns, NewPatterns: String); overload;



////////////////////////////////////////
//下位互換
////////////////////////////////////////

function ConvertHanKataToZenKata(const Source: String): String;
function ConvertZenKataToHanKata(const Source: String): String;

function ConvertNumericHanToZen(const Source: String): String;
function ConvertNumericZenToHan(const Source: String): String;
function ConvertSymbolHanToZen(const Source: String): String;
function ConvertSymbolZenToHan(const Source: String): String;
function ConvertAlphabetHanToZen(const Source: String): String;
function ConvertAlphabetZenToHan(const Source: String): String;
function ConvertAlphabetUpperCase(const Source: String): String;


implementation

uses

end_uses;

{---------------------------------------
    文字列変換
機能:   大文字小文字の変換のような1文字と対になった1文字で
        表される場合の変換を行う関数
備考:   全角半角変換が行えるのはUnicodeStringのおかげだが
        AnsiStringであっても半角の高速変換にこの
        ロジックを使うとよいだろう
履歴:   2011/12/21(水)
        ・  作成
}//(*-----------------------------------
type
  TCharDynArray       = array of Char;

procedure StringConvert(var S: String;
 OldPatterns, NewPatterns: TCharDynArray); overload;
var
  I, J: Integer;
begin
  if Length(OldPatterns) = 0 then Exit;
  if Length(OldPatterns) <> Length(NewPatterns) then Exit;
  if S = EmptyStr then Exit;

  for I := 1 to Length(S) do
  begin
    for J := 0 to Length(OldPatterns) - 1 do
    begin
      if S[I] = OldPatterns[J] then
      begin
        S[I] := NewPatterns[J];
      end;
    end;
  end;
end;

procedure StringConvert(var S: String;
 const OldPatterns, NewPatterns: String); overload;
var
  I, J: Integer;
begin
  if Length(OldPatterns) = 0 then Exit;
  if Length(OldPatterns) <> Length(NewPatterns) then Exit;
  if S = EmptyStr then Exit;

  for I := 1 to Length(S) do
  begin
    for J := 1 to Length(OldPatterns) do
    begin
      if S[I] = OldPatterns[J] then
      begin
        S[I] := NewPatterns[J];
      end;
    end;
  end;
end;
//------------------------------------*)

type
  TConvertTable = class
  public
    const HankakuKatakana: array[0..86] of String =
        (
        'ｶﾞ','ｷﾞ','ｸﾞ','ｹﾞ','ｺﾞ',
        'ｻﾞ','ｼﾞ','ｽﾞ','ｾﾞ','ｿﾞ',
        'ﾀﾞ','ﾁﾞ','ﾂﾞ','ﾃﾞ','ﾄﾞ',
        'ﾊﾞ','ﾋﾞ','ﾌﾞ','ﾍﾞ','ﾎﾞ',
        'ﾊﾟ','ﾋﾟ','ﾌﾟ','ﾍﾟ','ﾎﾟ',
        'ｱ','ｲ','ｳ','ｴ','ｵ',
        'ｶ','ｷ','ｸ','ｹ','ｺ',
        'ｻ','ｼ','ｽ','ｾ','ｿ',
        'ﾀ','ﾁ','ﾂ','ﾃ','ﾄ',
        'ﾅ','ﾆ','ﾇ','ﾈ','ﾉ',
        'ﾊ','ﾋ','ﾌ','ﾍ','ﾎ',
        'ﾏ','ﾐ','ﾑ','ﾒ','ﾓ',
        'ﾔ','ﾕ','ﾖ',
        'ﾗ','ﾘ','ﾙ','ﾚ','ﾛ',
        'ﾜ','ｦ','ﾝ',
        'ｧ','ｨ','ｩ','ｪ','ｫ',
        'ｬ','ｭ','ｮ',
        'ｯ','ﾟ','ｰ','･','､','｡','｢','｣');

    const ZenkakuKatakana: array[0..86] of Char =
        (
        'ガ','ギ','グ','ゲ','ゴ',
        'ザ','ジ','ズ','ゼ','ゾ',
        'ダ','ヂ','ヅ','デ','ド',
        'バ','ビ','ブ','ベ','ボ',
        'パ','ピ','プ','ペ','ポ',
        'ア','イ','ウ','エ','オ',
        'カ','キ','ク','ケ','コ',
        'サ','シ','ス','セ','ソ',
        'タ','チ','ツ','テ','ト',
        'ナ','ニ','ヌ','ネ','ノ',
        'ハ','ヒ','フ','ヘ','ホ',
        'マ','ミ','ム','メ','モ',
        'ヤ','ユ','ヨ',
        'ラ','リ','ル','レ','ロ',
        'ワ','ヲ','ン',
        'ァ','ィ','ゥ','ェ','ォ',
        'ャ','ュ','ョ',
        'ッ','゜','ー','・','、','。','「','」');

    const ZenkakuHiragana: array[0..86] of Char =
        (
        'が','ぎ','ぐ','げ','ご',
        'ざ','じ','ず','ぜ','ぞ',
        'だ','ぢ','づ','で','ど',
        'ば','び','ぶ','べ','ぼ',
        'ぱ','ぴ','ぷ','ぺ','ぽ',
        'あ','い','う','え','お',
        'か','き','く','け','こ',
        'さ','し','す','せ','そ',
        'た','ち','つ','て','と',
        'な','に','ぬ','ね','の',
        'は','ひ','ふ','へ','ほ',
        'ま','み','む','め','も',
        'や','ゆ','よ',
        'ら','り','る','れ','ろ',
        'わ','を','ん',
        'ぁ','ぃ','ぅ','ぇ','ぉ',
        'ゃ','ゅ','ょ',
        'っ','゜','ー','・','、','。','「','」');

    const Numeric: String =
        ('0123456789-+/.');
    const ZenkakuNumeric: String =
        ('０１２３４５６７８９−＋／．');

    const Symbol: String =
        (
        '!?$\%&#''"_' +
        '()[]<>{}' +
        '-+/*=.,;:@| ');

    const ZenkakuSymbol: String =
        (
        '！？＄￥％＆＃’”＿' +
        '（）［］＜＞｛｝' +
        '−＋／＊＝．，；：＠｜　');

    const AlphabetUpper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const AlphabetLower = 'abcdefghijklmnopqrstuvwxyz';
    const Alphabet = AlphabetUpper + AlphabetLower;

    const ZenkakuAlphabetUpper =
      'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ';
    const ZenkakuAlphabetLower =
      'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ';
    const ZenkakuAlphabet = ZenkakuAlphabetUpper + ZenkakuAlphabetLower;
  end;

    function StringToCharDynArray(const Table: String): TCharDynArray;
    var
      I: Integer;
    begin
      SetLength(Result, Length(Table));
      for I := 0 to Length(Table) - 1 do
        Result[I] := Table[I + 1];
    end;

    function CharDynArrayToStringDynArray(Table: array of char): TStringDynArray;
    var
      I: Integer;
    begin
      SetLength(Result, Length(Table));
      for I := 0 to Length(Table) - 1 do
        Result[I] := Table[I];
    end;

procedure StringConvert(var S: string; Option: TStringConvertOption); overload;
var
  Table1, Table2: String;
begin
  case Option of
    scUpperCase:
    begin
      Table1 := TConvertTable.AlphabetLower + TConvertTable.ZenkakuAlphabetLower;
      Table2 := TConvertTable.AlphabetUpper + TConvertTable.ZenkakuAlphabetUpper;
    end;
    scLowerCase:
    begin
      Table1 := TConvertTable.AlphabetUpper + TConvertTable.ZenkakuAlphabetUpper;
      Table2 := TConvertTable.AlphabetLower + TConvertTable.ZenkakuAlphabetLower;
    end;

    scAlphabetToZenkaku:
    begin
      Table1 := TConvertTable.AlphabetUpper + TConvertTable.AlphabetLower;
      Table2 := TConvertTable.ZenkakuAlphabetUpper + TConvertTable.ZenkakuAlphabetLower;
    end;
    scAlphabetToHankaku:
    begin
      Table1 := TConvertTable.ZenkakuAlphabetUpper + TConvertTable.ZenkakuAlphabetLower;
      Table2 := TConvertTable.AlphabetUpper + TConvertTable.AlphabetLower;
    end;

    scSymbolToZenkaku:
    begin
      Table1 := TConvertTable.Symbol;
      Table2 := TConvertTable.ZenkakuSymbol;
    end;
    scSymbolToHankaku:
    begin
      Table1 := TConvertTable.ZenkakuSymbol;
      Table2 := TConvertTable.Symbol;
    end;

    scNumericToZenakaku:
    begin
      Table1 := TConvertTable.Numeric;
      Table2 := TConvertTable.ZenkakuNumeric;
    end;
    scNumericToHankaku:
    begin
      Table1 := TConvertTable.ZenkakuNumeric;
      Table2 := TConvertTable.Numeric;
    end;

    scKatakanaToZenkaku:
    begin
      S := StringsReplace(S,
        TConvertTable.HankakuKatakana, CharDynArrayToStringDynArray(TConvertTable.ZenkakuKatakana));
      Exit;
    end;
    scKatakanaToHankaku:
    begin
      S := StringsReplace(S,
        CharDynArrayToStringDynArray(TConvertTable.ZenkakuKatakana), TConvertTable.HankakuKatakana);
      Exit;
    end;

    scKatakanaToHiragana:
    begin
      S := StringsReplace(S,
        TConvertTable.HankakuKatakana, CharDynArrayToStringDynArray(TConvertTable.ZenkakuKatakana));
      Table1 := TConvertTable.ZenkakuKatakana;
      Table2 := TConvertTable.ZenkakuHiragana;
    end;
    scHiraganaToKatakana:
    begin
      Table1 := TConvertTable.ZenkakuHiragana;
      Table2 := TConvertTable.ZenkakuKatakana;
    end;
  end;

  StringConvert(S, Table1, Table2);

end;

procedure StringConvert(Option: TStringConvertOption; var S: string); overload;
begin
  StringConvert(S, Option);
end;

////////////////////////////////////////
//下位互換
////////////////////////////////////////

{-------------------------------
//  カタカナを半角⇔全角相互変換します
    ConvertHanKataToZenKata
    ConvertZenKataToHanKata
機能:       カタカナを変換します
引数説明:   Source: 元の文字列
戻り値:     変換後の文字列
備考:
履歴:       2003/06/15
//--▼----------------------▽--}
const
  ConvertTblHanKata: array[0..86] of String =
       (
        'ｶﾞ','ｷﾞ','ｸﾞ','ｹﾞ','ｺﾞ',
        'ｻﾞ','ｼﾞ','ｽﾞ','ｾﾞ','ｿﾞ',
        'ﾀﾞ','ﾁﾞ','ﾂﾞ','ﾃﾞ','ﾄﾞ',
        'ﾊﾞ','ﾋﾞ','ﾌﾞ','ﾍﾞ','ﾎﾞ',
        'ﾊﾟ','ﾋﾟ','ﾌﾟ','ﾍﾟ','ﾎﾟ',
        'ｱ','ｲ','ｳ','ｴ','ｵ',
        'ｶ','ｷ','ｸ','ｹ','ｺ',
        'ｻ','ｼ','ｽ','ｾ','ｿ',
        'ﾀ','ﾁ','ﾂ','ﾃ','ﾄ',
        'ﾅ','ﾆ','ﾇ','ﾈ','ﾉ',
        'ﾊ','ﾋ','ﾌ','ﾍ','ﾎ',
        'ﾏ','ﾐ','ﾑ','ﾒ','ﾓ',
        'ﾔ','ﾕ','ﾖ',
        'ﾗ','ﾘ','ﾙ','ﾚ','ﾛ',
        'ﾜ','ｦ','ﾝ',
        'ｧ','ｨ','ｩ','ｪ','ｫ',
        'ｬ','ｭ','ｮ',
        'ｯ','ﾟ','ｰ','･','､','｡','｢','｣');
  ConvertTblZenKata: array[0..86] of String =
       (
        'ガ','ギ','グ','ゲ','ゴ',
        'ザ','ジ','ズ','ゼ','ゾ',
        'ダ','ヂ','ヅ','デ','ド',
        'バ','ビ','ブ','ベ','ボ',
        'パ','ピ','プ','ペ','ポ',
        'ア','イ','ウ','エ','オ',
        'カ','キ','ク','ケ','コ',
        'サ','シ','ス','セ','ソ',
        'タ','チ','ツ','テ','ト',
        'ナ','ニ','ヌ','ネ','ノ',
        'ハ','ヒ','フ','ヘ','ホ',
        'マ','ミ','ム','メ','モ',
        'ヤ','ユ','ヨ',
        'ラ','リ','ル','レ','ロ',
        'ワ','ヲ','ン',
        'ァ','ィ','ゥ','ェ','ォ',
        'ャ','ュ','ョ',
        'ッ','゜','ー','・','、','。','「','」');
function ConvertHanKataToZenKata(const Source: String): String;
var
  HanKanaPatterns, ZenKanaPatterns: TStringDynArray;
  i: Integer;
begin
  SetLength(HanKanaPatterns, High(ConvertTblHanKata)+1);
  for i := 0 to High(ConvertTblHanKata) do
    HanKanaPatterns[i] := ConvertTblHanKata[i];
  SetLength(ZenKanaPatterns, High(ConvertTblZenKata)+1);
  for i := 0 to High(ConvertTblZenKata) do
    ZenKanaPatterns[i] := ConvertTblZenKata[i];

  Result := StringsReplace(Source, HanKanaPatterns, ZenKanaPatterns);
end;

function ConvertZenKataToHanKata(const Source: String): String;
var
  HanKanaPatterns, ZenKanaPatterns: TStringDynArray;
  i: Integer;
  SymbolCount: Integer;
begin
  SymbolCount := 5;
  {↓全角→半角の場合、ひらがな記号『・、。「」』これらはカタカナにしなくてよい}
  SetLength(HanKanaPatterns, High(ConvertTblHanKata)+1 - SymbolCount);
  for i := 0 to High(ConvertTblHanKata) - SymbolCount do
    HanKanaPatterns[i] := ConvertTblHanKata[i];
  SetLength(ZenKanaPatterns, High(ConvertTblZenKata)+1 - SymbolCount);
  for i := 0 to High(ConvertTblZenKata) - SymbolCount do
    ZenKanaPatterns[i] := ConvertTblZenKata[i];

  Result := StringsReplace(Source, ZenKanaPatterns, HanKanaPatterns);
end;
//--△----------------------▲--

{-------------------------------
//  英語と数値と記号半角⇔全角相互変換します
    ConvertAlphabetHanToZen
    ConvertAlphabetZenToHan
    ConvertNumericHanToZen
    ConvertNumericZenToHan
    ConvertSymbolHanToZen
    ConvertSymbolZenToHan
機能:       数値と記号を変換します
引数説明:   Source: 元の文字列
戻り値:     変換後の文字列
備考:
履歴:       2006/04/05
//--▼----------------------▽--}
const
  ConvertTblHanNumeric: String =
       ('0123456789-+/.');
  ConvertTblZenNumeric: String =
       ('０１２３４５６７８９−＋／．');
function ConvertNumericHanToZen(const Source: String): String;
var
  HanNumericPatterns, ZenNumericPatterns: TStringDynArray;
  i: Integer;
begin
  SetLength(HanNumericPatterns, Length(ConvertTblHanNumeric));
  for i := 0 to Length(ConvertTblHanNumeric)-1 do
    HanNumericPatterns[i] := ConvertTblHanNumeric[i+1];
  SetLength(ZenNumericPatterns, Length(ConvertTblZenNumeric));
  for i := 0 to Length(ConvertTblZenNumeric)-1 do
    ZenNumericPatterns[i] := ConvertTblZenNumeric[i+1];

  Result := StringsReplace(Source, HanNumericPatterns, ZenNumericPatterns);
end;

function ConvertNumericZenToHan(const Source: String): String;
var
  HanNumericPatterns, ZenNumericPatterns: TStringDynArray;
  i: Integer;
begin
  SetLength(HanNumericPatterns, Length(ConvertTblHanNumeric));
  for i := 0 to Length(ConvertTblHanNumeric)-1 do
    HanNumericPatterns[i] := ConvertTblHanNumeric[i+1];
  SetLength(ZenNumericPatterns, Length(ConvertTblZenNumeric));
  for i := 0 to Length(ConvertTblZenNumeric)-1 do
    ZenNumericPatterns[i] := ConvertTblZenNumeric[i+1];

  Result := StringsReplace(Source, ZenNumericPatterns, HanNumericPatterns);
end;


const
  ConvertTblHanSymbol: String =
       ('!?$\%&#''"_' +
        '()[]<>{}' +
        '-+/*=.,;:@| ');
  ConvertTblZenSymbol: String =
       ('！？＄￥％＆＃’”＿' +
        '（）［］＜＞｛｝' +
        '−＋／＊＝．，；：＠｜　');
function ConvertSymbolHanToZen(const Source: String): String;
var
  HanSymbolPatterns, ZenSymbolPatterns: TStringDynArray;
  i: Integer;
begin
  SetLength(HanSymbolPatterns, Length(ConvertTblHanSymbol));
  for i := 0 to Length(ConvertTblHanSymbol)-1 do
    HanSymbolPatterns[i] := ConvertTblHanSymbol[i+1];
  SetLength(ZenSymbolPatterns, Length(ConvertTblZenSymbol));
  for i := 0 to Length(ConvertTblZenSymbol)-1 do
    ZenSymbolPatterns[i] := ConvertTblZenSymbol[i+1];

  Result := StringsReplace(Source, HanSymbolPatterns, ZenSymbolPatterns);
end;

function ConvertSymbolZenToHan(const Source: String): String;
var
  HanSymbolPatterns, ZenSymbolPatterns: TStringDynArray;
  i: Integer;
begin
  SetLength(HanSymbolPatterns, Length(ConvertTblHanSymbol));
  for i := 0 to Length(ConvertTblHanSymbol)-1 do
    HanSymbolPatterns[i] := ConvertTblHanSymbol[i+1];
  SetLength(ZenSymbolPatterns, Length(ConvertTblZenSymbol));
  for i := 0 to Length(ConvertTblZenSymbol)-1 do
    ZenSymbolPatterns[i] := ConvertTblZenSymbol[i+1];

  Result := StringsReplace(Source, ZenSymbolPatterns, HanSymbolPatterns);
end;

const
  ConvertTblAlphabetUpper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  ConvertTblAlphabetLower = 'abcdefghijklmnopqrstuvwxyz';
  ConvertTblAlphabet = ConvertTblAlphabetUpper + ConvertTblAlphabetLower;

  ConvertTblZenkakuAlphabetUpper =
    'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ';
  ConvertTblZenkakuAlphabetLower =
    'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ';
  ConvertTblZenkakuAlphabet = ConvertTblZenkakuAlphabetUpper + ConvertTblZenkakuAlphabetLower;

function ConvertAlphabetHanToZen(const Source: String): String;
var
  HanAlphabetPatterns, ZenAlphabetPatterns: TStringDynArray;
  i: Integer;
begin
  SetLength(HanAlphabetPatterns, Length(ConvertTblAlphabet));
  for i := 0 to Length(ConvertTblAlphabet)-1 do
    HanAlphabetPatterns[i] := ConvertTblAlphabet[i+1];
  SetLength(ZenAlphabetPatterns, Length(ConvertTblZenkakuAlphabet));
  for i := 0 to Length(ConvertTblZenkakuAlphabet)-1 do
    ZenAlphabetPatterns[i] := ConvertTblZenkakuAlphabet[i+1];

  Result := StringsReplace(Source, HanAlphabetPatterns, ZenAlphabetPatterns);
end;

function ConvertAlphabetZenToHan(const Source: String): String;
var
  HanAlphabetPatterns, ZenAlphabetPatterns: TStringDynArray;
  i: Integer;
begin
  SetLength(HanAlphabetPatterns, Length(ConvertTblAlphabet));
  for i := 0 to Length(ConvertTblAlphabet)-1 do
    HanAlphabetPatterns[i] := ConvertTblAlphabet[i+1];
  SetLength(ZenAlphabetPatterns, Length(ConvertTblZenkakuAlphabet));
  for i := 0 to Length(ConvertTblZenkakuAlphabet)-1 do
    ZenAlphabetPatterns[i] := ConvertTblZenkakuAlphabet[i+1];

  Result := StringsReplace(Source, ZenAlphabetPatterns, HanAlphabetPatterns);
end;

  function StringTableToDynArray(Table: String): TStringDynArray;
  var
    I: Integer;
  begin
    SetLength(Result, Length(Table));
    for I := 0 to Length(Table) - 1 do
      Result[I] := Table[I + 1];
  end;

function ConvertAlphabetUpperCase(const Source: String): String;
begin
  Result := StringsReplace(Source,
    StringTableToDynArray(ConvertTblAlphabetLower + ConvertTblZenkakuAlphabetLower),
    StringTableToDynArray(ConvertTblAlphabetUpper + ConvertTblZenkakuAlphabetUpper));
end;
//--△----------------------▲--





end.
