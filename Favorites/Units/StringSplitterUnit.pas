(*--▽---------------------------▼--
セパレータで文字を分解する処理
00/05/11
00/08/22    実装を丸ごと変更しました。
            単語分解方法を幾つか選択出来るようにしたので
            ExcelのCSVファイルなど読むときも利用できます。
            他にもDelimitersを引数指定するようにして使い勝手が上がっています。
00/08/25    Wordxxx2という関数名をWordxxxにしました
01/03/29    タブ文字の対応が出来ていなかったので修正
            カンマ文字がDelimitersに含まれていない場合の処理を修正
            WordCountとWordGetの内部ルーチンをWordDecomposerとして1つにまとめた
01/10/29    TWordDecomposeを実装
            そのために無駄がないように function WordDecomposer に
            StringListを引数で渡すようにした
02/07/04    TWordDecomposeにIndexOfを実装した
2002/11/07
  Split関数をStringUnitLightから移動
  Split関数の内部を強化、WordDecomposerの内部処理をすべてまかす
2010/03/03(水)
・  WordDecompose.pasから
    DelimitedTextUnit.pasと名前変更
・  TSpliterをTWideStringSpliterとして実装をまとめた
・  WordCount/WordGetのフラグDecomposedModeを
    TSplitFlagsで置き換えた
2010/03/05(金)
・  WordCount/WordGetをTSplitterで置き換えた
2010/03/07(日)
・  Split関数をWideStringRecordList.SetBaseTextで置き換えて
    テストを通過した。
    SetBaseTextの文字分割機能は大幅に強化した。
・  TWordDecomposeの内部をSplit関数から
    WideStringRecordList.SetBaseTextで置き換えた
・  TSplitterの役目はなくなったので廃止した。
2011/05/11(水)
・  GrepExtensionを追加した。
2011/06/10(金)
・名前を DelimitedTextUnit から StringSplitterUnit に変更した。
・TWordSplited から TStringSplitter に変更した。
2011/08/12(金)
・GrepExtensionを改良して uses Classes を排除しました
//--▲---------------------------△--*)
unit StringSplitterUnit;

interface

uses
  Types,
  SysUtils,
  StringUnit,
  StringRecordList,
  ConstUnit,

uses_end;

type
  TSplitFlags = set of (sfInDelimiter, sfInEmptyStr);
  {↑sfIncludeDelimiter
       分割時の戻り値に区切り文字自体も含まれる
     sfEmptyStr
       空文字も分割したとみなす
       例えば[A,B,]を分割する時[A][B][空文字]に分解される}

const
  dmUserFriendly: TSplitFlags = [];
  dmDelimiterExactly: TSplitFlags = [sfInEmptyStr];
  {単語分解方法
   UserFriendlyは区切り文字が複数個でも単語単位に分解
   DelimiterExactlyは区切り文字に完全に正確に分解
   接頭語のdmはDelimitModeという事にする}

//  TDecomposeMode = (dmUserFriendly, dmDelimiterExactly);
//  元は列挙型だったが、TSplitFlagsに統合した。

function WordCount(const S: String; Delimiters: array of String;
 const SplitFlag: TSplitFlags): Integer;
function WordGet(const S: String; Delimiters: array of String;
 WordIndex: Integer; SplitFlag: TSplitFlags): String;

type
  TStringSplitter = class(TObject)
  private
    FWords: TStringRecordList;
    FSentence: String;
    FDelimiters: TStringDynArray;
    function GetWords(Index: Integer): String;
    function GetCount: Integer;
  public
    constructor Create(const Sentence: String;
     Delimiters: array of String; SplitFlag: TSplitFlags);
    destructor Destroy; override;
    property Words[Index: Integer]: String read GetWords;
    property Count: Integer read GetCount;
    property Sentence: String read FSentence;
    property Delimiters: TStringDynArray read FDelimiters;
    function IndexOf(const Word: String; IgnoreCase: Boolean): Integer;
  end;
  {Create時に単語分解を行ってしまうクラス
   WordCount/WordGet毎に処理する無駄をはぶく時に利用できる}

function Split(const S: String; const Delimiters: array of String;
 SplitFlag: TSplitFlags): TStringDynArray; overload;

function GrepExtension(Filter, Target: String): Boolean;


implementation


type
  TWideStringRecordListAccess = class(TStringRecordList);

function SplitFlagsToDelimitStyle(SplitFlag: TSplitFlags): TDelimitStyle;
begin
  if SplitFlag = [sfInDelimiter, sfInEmptyStr] then
    Result := dsInDelimInEmpty
  else
  if SplitFlag = [sfInEmptyStr] then
    Result := dsNoDelimInEmpty
  else
  if SplitFlag = [sfInDelimiter] then
    Result := dsInDelimNoEmpty
  else
  if SplitFlag = [] then
    Result := dsNoDelimNoEmpty
  else
  begin
    Result := dsLineBreaks;
    Assert(False, 'Error SplitFlagsToDelimitStyle');
  end;
end;

{----------------------------------------
//      WordGet/WordCount
        
機能:       区切り文字を指定して単語に分割
			単語を取得したり個数を調べる事ができる
備考:       毎回、分割処理が走るので
			処理によっては他のSplit/TStringSplitterを使うとよい
            もっと細かい制御は TWideStringRecordList 本体を使って
            実装するとよいかもしれない
履歴:       2010/03/07(日)
            ・  TWideStringRecordList.SetBaseTextを使って
                実装を更新した。
//----------------------------------------}
function WordCount(const S: String; Delimiters: array of String;
 const SplitFlag: TSplitFlags): Integer;
begin
  with TWideStringRecordListAccess.Create do try
  SetBaseText(S, Delimiters,
    SplitFlagsToDelimitStyle(SplitFlag));
  Result := Count;
  finally Free; end;
end;

function WordGet(const S: String; Delimiters: array of String;
 WordIndex: Integer; SplitFlag: TSplitFlags): String;
begin
  with TWideStringRecordListAccess.Create do try
  SetBaseText(S, Delimiters,
    SplitFlagsToDelimitStyle(SplitFlag));
  Result := Items[WordIndex];
  finally Free; end;
end;
//----------------------------------------

{----------------------------------------
//      WordGet/WordCountのクラス版
機能:       WordCountしてループして
            要素一つ一つをWordGetするというコードでは
            毎回、分割処理が走るので無駄なので
            クラス化して分割処理は最初の1回に限定して
            あとはキャッシュを読む実装になっている。
備考:       
履歴:       2010/03/07(日)
            ・  WordGet/Countと同様に内部実装を変更した
				TWideStringRecordList本体をラッピングしている
				という形にもなっている。
//----------------------------------------}
{ TStringSplitter }

constructor TStringSplitter.Create(const Sentence: String;
 Delimiters: array of String; SplitFlag: TSplitFlags);
var
  i: Integer;
begin
  FSentence := Sentence;
  FWords := TStringRecordList.Create;
  TWideStringRecordListAccess(FWords).SetBaseText(Sentence, Delimiters,
    SplitFlagsToDelimitStyle(SplitFlag));

  SetLength(FDelimiters, Length(Delimiters));
  for i := 0 to Length(Delimiters) - 1 do
  begin
    FDelimiters[i] := Delimiters[i];
  end;
end;

destructor TStringSplitter.Destroy;
begin
  FWords.Free;
  inherited;
end;

function TStringSplitter.GetCount: Integer;
begin
  Result := FWords.Count;
end;

function TStringSplitter.GetWords(Index: Integer): String;
begin
  if ( 0 <= Index ) and ( Index <= (GetCount-1) ) then
  begin
    Result := FWords.Items[Index];
  end else
  begin
    Result := '';
  end;
end;

function TStringSplitter.IndexOf(const Word: String;
 IgnoreCase: Boolean): Integer;
var
  i: Integer;
  CompareFunction: function(const S1, S2: WideString): Boolean;
begin
  Result := -1;

  for i := 0 to GetCount - 1 do
  begin
    if (IgnoreCase=True) then
    begin
      CompareFunction := WideSameText;
    end else
    begin
      CompareFunction := WideSameStr;
    end;

    if CompareFunction(FWords[i], Word) then
    begin
      Result := i;
      Exit;
    end;
  end;
end;
//----------------------------------------

{----------------------------------------
//      Split
機能:       区切り文字で分解した後
			動的配列で返す関数
備考:       
履歴:       2002/09/27
              作成
            2002/10/02
              ItemAddDelimiterフラグを追加
            2002/11/07
              ItemAddEmptyStrフラグを追加
			2010/03/07(日)
			・	実装を大幅に変えている。
			・	全てのテストに通過しているので
				品質には影響ないはずです。
//----------------------------------------}
function Split(const S: String; const Delimiters: array of String;
 SplitFlag: TSplitFlags): TStringDynArray;
var
  Spliter: TWideStringRecordListAccess;
  i: Integer;
begin
  Spliter := TWideStringRecordListAccess.Create; try
  Spliter.SetBaseText(S, Delimiters,
    SplitFlagsToDelimitStyle(SplitFlag));

  SetLength(Result, Spliter.Count);
  for i := 0 to Spliter.Count - 1 do
  begin
    Result[i] := Spliter.Items[i];
  end;
  finally Spliter.Free; end;
end;
//----------------------------------------


{---------------------------------------
    AndとOrを組み合わせた検索をするための関数
機能:   Filterに「AA BB CC\r\nDD EE」を指定すると
        ある文字列を、AAとBBとCCが含まれているか、DDとEEが含まれているか
        どうかを調べることができる。
        演算子で表現すると (AA and BB and CC)or(DD and EE)となる。
備考:   空文字を指定すると戻り値がTrueになるのは
        フィルタとして使う時に全て全ての項目が有効になるべきだから。
履歴:   2011/05/11(水)
        ・  作成
        2011/08/12(金)
        ・  容量を食うのでTStringsを廃止。uses Classes が不要になった。
}//(*-----------------------------------
function GrepExtension(Filter, Target: String): Boolean;
var
  K, J: Integer;
  AndFilterWords: TStringSplitter;
  AndFilterFlag: Boolean;
  FilterStrRecList: TStringRecordList;
begin
  if TrimChar(Filter, ' ' + '　'+CtrlCharTbl)=EmptyStr then
  begin
    Result := True;
  end else
  begin
    Result := False;

    FilterStrRecList := TStringRecordList.Create; try
    FilterStrRecList.Text := Filter;
    for K := 0 to FilterStrRecList.Count - 1 do
    begin
      AndFilterWords := TStringSplitter.Create(FilterStrRecList[K], [' ', '　'], dmUserFriendly); try
      if AndFilterWords.Count = 0 then
      begin
        AndFilterFlag := False;
      end else
      begin
        AndFilterFlag := True;

        for J := 0 to AndFilterWords.Count - 1 do
        begin
          if not InStr(AndFilterWords.Words[J], Target, ccIgnoreCase) then
          begin
            AndFilterFlag := False;
            Break;
          end;
        end;
      end;
      Result := Result or AndFilterFlag;
      finally AndFilterWords.Free; end;
    end;
    finally FilterStrRecList.Free; end;
  end;
end;
//------------------------------------*)

end.
