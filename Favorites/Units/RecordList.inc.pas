{ --▽---------------------------▼--
    構造体リスト/クラスリストの
    テンプレートのような実装
履歴:
2002/05/15
  MyClassList.pas/NumericListなどを実装
2002/08/15
  StringClassに文字列引数を取るコンストラクタを実装
  StringClassList/StringRecordListにTextプロパティを実装
2002/08/16
  StringRecordListのためにMyRecordListにInitialize/Finalizeを実装
  TListに依存しないためにTMｙListを実装
2003/09/19
  構造体破棄の処理を修正(SetCountなどでメモリリークしてた)
  GetMem Initialize/FreeMem Finalize処理を
  New/Disposeの処理に変更した
  参考:[Delphi-ML:18256] Re: RE: Re: 動的に配列を作る方法
2003/11/09
  リファクタリングしてTMyRecordListをTRecordListにする
2007/04/29
・FList/GetItem/SetItem/GetCount/SetCountを
  privateからprotectedにした。
  ソート実装で継承クラスでFListにアクセスする必要があったから
//--▲---------------------------△-- }

////////////////////////////////////////////////////////////
{$ifndef RecordList}
{$define RecordList}
// interface
////////////////////////////////////////////////////////////

  PRecord = ^TRecord;

  TRecordList = class(TObject)
  private
  protected
    FList: TListClone;
    function GetItem(Index: Integer): TRecord;
    procedure SetItem(Index: Integer; const Value: TRecord);
    function GetCount: Integer;
    procedure SetCount(const Value: Integer);
  public
    constructor Create;
    destructor Destroy; override;
    function Add(Item:TRecord): Integer;
    procedure Delete(Index: Integer);
    procedure Insert(Index: Integer; Item: TRecord);
    procedure Move(CurIndex, NewIndex: Integer);
    procedure Exchange(Index1, Index2: Integer);
    procedure Clear;
	procedure Assign(Source: TRecordList); virtual;
    function IndexOf(const Value: TRecord): Integer;
    property Items[Index:Integer]:TRecord read GetItem write SetItem;default;
    property Count: Integer read GetCount write SetCount;
  end;

////////////////////////////////////////////////////////////
{$else}
// implementation
////////////////////////////////////////////////////////////

//-------------------------------
{ TRecordList }
//-------------------------------


constructor TRecordList.Create;
begin
  FList := TListClone.Create;
end;

destructor TRecordList.Destroy;
begin
  Self.Clear;
  FList.Free;
  inherited;
end;

function TRecordList.Add(Item:TRecord): Integer;
var
  p:PRecord;
begin
  {↓Newの変わりに
       GetMem(p,SizeOf(Item));
       Initialize(p^);
     でもよいらしい}
  New(p);
  p^ := Item;
  result := FList.Add(p);
end;

procedure TRecordList.Delete(Index: Integer);
begin
  {↓Disposeの変わりに
       Finalize(PRecord(FList[Index])^);
       FreeMem(PRecord(FList[Index]));
     でもよいらしい}
  Dispose(PRecord(FList[Index]));
  FList.Delete(Index);
end;

procedure TRecordList.Insert(Index: Integer; Item: TRecord);
var
  p: PRecord;
begin
  New(p); {←GetMem(p,SizeOf(Item))でもいいのかな？
             Initializeは代入しているから要らない気がする}
  p^ := Item;
  FList.Insert(Index, p);
end;

procedure TRecordList.Clear;
var
  i:integer;
begin
  for i := 0 to FList.Count-1 do
  begin
    Dispose(PRecord(FList[i]));
  end;
  FList.Clear;
end;

procedure TRecordList.Exchange(Index1, Index2: Integer);
begin
  FList.Exchange(Index1, Index2);
end;

procedure TRecordList.Move(CurIndex, NewIndex: Integer);
begin
  FList.Move(CurIndex, NewIndex);
end;

function TRecordList.GetItem(Index: Integer): TRecord;
begin
  result := PRecord(FList[Index])^;
end;

procedure TRecordList.SetItem(Index: Integer; const Value: TRecord);
begin
  PRecord(FList[Index])^ := Value;
end;

function TRecordList.IndexOf(const Value: TRecord): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to FList.Count-1 do
  begin
    if RecordEqual(Value, GetItem(i)) then
    begin
      Result := i;
      Break;
    end
  end;
end;

function TRecordList.GetCount: Integer;
begin
  result := FList.Count;
end;

procedure TRecordList.SetCount(const Value: Integer);
var
  i: Integer;
  p: PRecord;
begin
  if FList.Count = Value then
    Exit
  else
  if Value < FList.Count then
  begin
    {↓SetCountが現在のListItemの数より少ないので
       リストに定義してある余分な分を削除する}
    for i := FList.Count-1 downto Value do
    begin
      Self.Delete(i);
    end;
  end else
  if FList.Count < Value then
  begin
    {↓SetCountが現在のListItemの数より多いので}
    for i := FList.Count+1 to Value do
    begin
      {↓メモリを確保して0クリアした値をAddする
         Newの変わりに
           GetMem(p, SizeOf(TRecord));
           FillChar(p^, SizeOf(TRecord), 0);
         というコードでもよいらしい}
      New(p);
      FList.Add(p);
    end;
  end;
end;

procedure TRecordList.Assign(Source: TRecordList);
var
  i: Integer;
begin
  if Source is TRecordList then
  begin
    Self.Clear;
    for I := 0 to TRecordList(Source).Count - 1 do
    begin
      Self.Add(TRecordList(Source).Items[I]);
    end;
  end;

//  inherited Assign(Source);
end;

////////////////////////////////////////////////////////////
{$endif}
////////////////////////////////////////////////////////////

