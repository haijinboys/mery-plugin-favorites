{ --▽---------------------------▼--
長いファイル名、短いファイル名相互変換関数
2002/06/05
  作成されていた
2003/03/17
  関数の説明を記述、全体的にコメントなどを修正
  GetShortFullPathName/GetLongFullPathNameは
  不要なのでimplementationだけに実装して
  interfaceにはGetLongFileName/GetShortFileNameを
  準備した。
2005/12/19
  LongShortFileName.pasからFileNameUnit.pasに名前変更
2005/12/29
  CheckUNCPathやCheckDrivePathを正しく実装
2006/12/24(日)
・CheckDrivePathをデバッグ
・PathLevel/CutPathLevelを実装
・testTrailingPathDelimiterを追加
2007/07/22(日) 00:04
・FileUnit.pasからExtractFolderFileName関数を移動
2010/03/10(水)
・GetAbsolutePathFromRelativePath/GetRelativePathFromAbsolutePath
  を実装した。
2010/11/10
・RepaceFlagをccIgnoreCase/ccCaseSensitiveに変更した。
2011/05/23(月)
・GetEnvironmentStringArrayを追加した。
2012/01/15(日)
・CheckCanRelativePathを追加
  そのために、CheckUNCPathHeaderとCheckDrivePathHeaderを作成して
  CheckUNCPathとCheckDrivePathをリファクタリング
//--▲---------------------------△-- }
unit FileNameUnit;

interface

uses
  Windows, SysUtils, Types,

//  ShellUnit,
//  ShellFileCtrl,
//  NetworkUnit,
  ConstUnit,
  StringUnit,
  StringSearchUnit,
  StringSplitterUnit,
uses_end;

function GetLongFileName(const ShortFileName:string):string;
function GetShortFileName(const LongFileName:string):string;

function CheckSameFileName(A, B: String): Boolean;
function CheckUNCPath(Path: String): Boolean;
function CheckDrivePath(Path: String): Boolean;
function CheckPathFollowRule(Path: String): Boolean; forward;
function CheckCanRelativePath(Path: string): Boolean;

function ExtractFileNameExcludeExt(const FileName: String): String;
function ExtractFolderFileName(Path: String): String;

function CheckSameExt(const Filename, Ext: String): Boolean;

function PathLevel(Path: String): Integer;

function CutPathLevel(Path: String; CutCount: Integer): String;

function GetUpFolderPath(FilePath: string): String;

function GetAbsolutePathFromRelativePath(BasePath, RelativePath: WideString): WideString;
function GetRelativePathFromAbsolutePath(BasePath, AbsolutePath: WideString): WideString;

function GetEnvironmentStringArray: TStringDynArray;
function ReplaceEnvironmentVariableNameToValue(Source: string): String;


implementation


{-------------------------------
//ロングファイル名/ショートファイル名を取得する
    GetLongFileName
    GetShortFileName
備考:       DelphiML検索ワード「ロングファイル」
            2chの情報
            結局、FDelphiのコードよりこちらの方が
            短くてよいだろう
履歴:       2003/03/17
              存在しないファイルを変換した場合
              空文字を返すように修正
//--▼----------------------▽--}
function GetLongFileName(const ShortFileName:string):string;
var
  Path:string;
  SearchRec:TSearchRec;
begin
  Result := '';
  if ShortFileName = '' then Exit;

  if FileExists(ShortFileName)
    or DirectoryExists(ShortFileName) then
  begin
    Result := ExcludeTrailingPathDelimiter( ShortFileName );
    if ( Length( Result ) = 2 ) and ( Result[2] = ':' ) then Exit;
    if SysUtils.FindFirst( Result, faAnyFile, SearchRec )=0 then
    begin
      Path := GetLongFileName( ExtractFileDir( Result ) );
      Path := IncludeTrailingPathDelimiter( Path );
      Result := Path + SearchRec.Name;
    end;
    SysUtils.FindClose( SearchRec );
  end;
end;

function GetShortFileName(const LongFileName:string):string;
var
  Path:string;
  SearchRec:TSearchRec;
begin
  Result := '';
  if LongFileName = '' then Exit;

  if FileExists(LongFileName)
    or DirectoryExists(LongFileName) then
  begin
    Result := ExcludeTrailingPathDelimiter( LongFileName );
    if ( Length( Result ) = 2 ) and ( Result[2] = ':' ) then Exit;
    if SysUtils.FindFirst( Result, faAnyFile, SearchRec )=0 then
    begin
      Path := GetShortFileName( ExtractFileDir( Result ) );
      Path := IncludeTrailingPathDelimiter( Path );
      if SearchRec.FindData.cAlternateFileName <> '' then
        Result := Path + SearchRec.FindData.cAlternateFileName
      else
        Result := Path + SearchRec.Name;
    end;
    SysUtils.FindClose( SearchRec );
  end;
end;
//--△----------------------▲--


{-------------------------------
//ファイル名を比較して
機能:       文字列が同一のファイルを示しているのかどうかを判断します。
            存在するファイルしか同一かどうか判断できません。
            A,Bは短いファイル名でも長いファイル名でもOKです。
戻り値:     true:同じファイル false:異なるファイル
備考:       ユニットSysUtilsのSameFileName関数で代用できるかもしれません
履歴:       2002/02/11
//--▼----------------------▽--}
function CheckSameFileName(A, B: String): Boolean;
var
  OldMode: UINT;// エラーモード保持用
begin
  OldMode := SetErrorMode(SEM_FAILCRITICALERRORS);
  if FileExists(A) and FileExists(B) then
  begin
    A := GetLongFileName(A);
    B := GetLongFileName(B);

    if SameText(A, B) then
    begin
      Result := True;
    end else
    begin
      Result := False;
    end;

  end else
  begin
    Result := False;
  end;
  SetErrorMode(OldMode);
end;
//--△----------------------▲--



{-------------------------------
//  UNC(=Universal Naming Convention)パスかどうかを判断する関数
機能:       文字列が3文字以上であって
            先頭が『\\?』って形になっていて
            途中でファイル禁止文字や『\\』がない
            かどうかを判断する関数
備考:
履歴:       2005/12/28
2012/01/11(水)
・  修正
//--▼----------------------▽--}
  function CheckUNCPathHeader(Path: string): Boolean;
  begin
    Result := False;
    if Length(Path) <= 2 then Exit;
    if ((Path[1]=PathDelim)
      and (Path[2]=PathDelim)
      and (not (Path[3]=PathDelim))) = False then Exit;
    Result := True;
  end;

function CheckUNCPath(Path: String): Boolean;
begin
  Result := False;

  if CheckUNCPathHeader(Path) = False then Exit;

  if CheckPathFollowRule(Path) then
  begin
    Result := True;
  end;
end;

//--△----------------------▲--

{-------------------------------
//  通常のドライブパスかどうかを判断する関数
機能:       文字列が2文字以上であって
            先頭が『X:』って形になっていて
            途中でファイル禁止文字や『\\』がない
            かどうかを判断する関数
備考:
履歴:       2005/12/28
            2006/12/24(日) 誤動作を若干修正
//--▼----------------------▽--}

  function CheckDrivePathHeader(Path: string): Boolean;
  begin
    Result := False;
    if Length(Path)<=1 then Exit;
    {↑一文字の場合はどうやってもパスにならない}

    if 2 <= Length(Path) then
      if ((Path[2]=DriveDelim)
        and CheckCharInTable(Path[1], hanAlphaTbl)) = False then Exit;

    if 3 <= Length(Path) then
      if Path[3] <> PathDelim then Exit;

    Result := True;
  end;

function CheckDrivePath(Path: String): Boolean;
begin
  Result := False;

  if CheckDrivePathHeader(Path) = False then Exit;

  if CheckPathFollowRule(Path) then
  begin
    Result := True;
  end;
end;
//--△----------------------▲--


{---------------------------------------
    相対パスになれる文字列かどうか確認する関数
機能:   
備考:   
履歴:   2012/01/15(日)
        ・  作成
}//(*-----------------------------------
function CheckCanRelativePath(Path: string): Boolean;
begin
  Result := False;
  if CheckPathFollowRule(Path) = False then Exit;

  if CheckUNCPathHeader(Path) then Exit;
  if CheckDrivePathHeader(Path) then Exit;
  Result := True;
end;
//------------------------------------*)



{-------------------------------
//  CheckDrivePathとCheckUNCPathで使われる内部関数
機能:       文字列が1文字しかなかったり
            UNC指定以外の場所(つまり1文字目以外)で"\\"が使われていたり
            その他禁止文字が使われていたりして
            Pathがファイルパスとして成り立つかどうかを判断する関数
仕様:
・パスとして正しいルールに含まれている時はTrue
  正しくないならFalse

  ・\記号によって空文字を含む分割を行う。分割記号は含まない。
  ・先頭に\\と連続している場合
    つまり、空文字要素が先頭の2連続で続くのはOK
    1つだけの空文字要素ならNotOK
  ・[:]が先頭要素の2文字目以外にあるとエラー
    >>でも「..\D:\という場合があるな…」>>対応した
  ・区切られた要素に空文字がある場合はエラー
    (つまり\\記号があった場合はエラー)
  ・区切られた要素に[...]で始まる内容があるとエラー
  ・区切られた要素に[/*?"<>|]が含まれているとエラー
備考:
履歴:       2005/12/29
            2011/05/23(月)
            ・  仕様をまとめた。Unicode版で動作確認
//--▼----------------------▽--}

function CheckPathFollowRule(Path: String): Boolean;
var
  Split: TStringSplitter;
  SearchStartIndex: Integer;
  FirstItemFlag: Boolean;
//  DriveItemIndex: Integer;
  i: Integer;
begin
  Result := True;
  Split := TStringSplitter.Create(ExcludeLastPathDelim(Path), [PathDelim], [sfinEmptyStr]); try
  if (Split.Count = 1) and (Split.Words[0] = EmptyStr) then
  begin
    Result := False;
    Exit;
  end;

  {↓\で区切ったときの空項目、つまり\が連続で並んでいる場合に
     許容されるのは[\\abc]となっている時だけ
     それ以外は\が連続とみなされるので
     ファイルパスではない}
  SearchStartIndex := 0;
  if (2 <= Split.Count) and (Split.Words[0] = EmptyStr) then
  begin
    {↑↓Words[0]がEmptyStrでWords[1]がEmptyStrではないなら
       [\]が先頭に1文字の場合だから
       ファイルパスではない}
    if Split.Words[1] <> EmptyStr then
    begin
      Result := False;
      Exit;
    end else
    if Path = PathDelim + PathDelim then
    begin
      Result := False;
      Exit;
    end else
    begin
      SearchStartIndex := 2;
    end;
  end;

  for i := SearchStartIndex to Split.Count - 1 do
  begin
    if Split.Words[i] = EmptyStr then
    begin
      Result := False;
      Exit;
    end;
  end;

  FirstItemFlag := True;
  for i := 0 to Split.Count - 1 do
  begin
    if (Split.Words[i] = '.') or (Split.Words[i] = '..') then
    begin
      if FirstItemFlag then
      begin
        Continue;
      end else
      begin
        {↓先頭項目じゃないのに[.]や[..]なら
           ファイルパスではない(ことにする)}
        Result := False;
        Exit;
      end;
    end;

    if FirstItemFlag then
    begin
      FirstItemFlag := False;
      {↓最初の項目で[:]が含まれていて
         それが2文字目ではない場合はファイルパスではない}
      if InStr(DriveDelim, Split.Words[i]) then
        if not InStr(DriveDelim, Split.Words[i], 2, 1) then
        begin
          Result := False;
          Exit;
        end else
        begin
          {↓2文字目に[:]でも1文字目がアルファベットではないなら
             ファイルパスではない}
          if not (CheckStrInTable(Split.Words[i][1], hanAlphaTbl)=itAllInclude) then
          begin
            Result := False;
            Exit;
          end;
        end;
    end else
    begin
      {↓最初の項目以外で[:]が含まれているなら
         ファイルパスではない}
      if InStr(DriveDelim, Split.Words[i])  then
      begin
        Result := False;
        Exit;
      end;
    end;
  end;

  for i := 0 to Split.Count - 1 do
  begin
    if (Split.Words[i] = '.') or (Split.Words[i] = '..') then continue;

    {↓[.]や[..]以外で先頭に[.]がある場合は
       ファイルパスではない
       [A...]という項目はOK}
    if PosForward('.', Split.Words[i]) = 1 then
    begin
      Result := False;
      Exit;
    end;

    if CheckStrInTable(Path, '/*?"<>|') <> itAllExclude then
    begin
      Result := False;
      Exit;
    end;
  end;

  finally Split.Free; end;
end;

//--△----------------------▲--

{-------------------------------
//  ExtractFileNameExcludeExt
機能:       拡張子もディレクトリパスもない
            ファイル名を取得する
備考:       ファイル名が正しくない場合
            戻り値はChangeFileExtやExtractFileNameの処理に依存する
履歴:       2006/02/18
//--▼----------------------▽--}
function ExtractFileNameExcludeExt(const FileName: String): String;
begin
  Result := ChangeFileExt(ExtractFileName(FileName), '');
end;
//--△----------------------▲--

{-------------------------------
//  パスの末尾から"AAA\BBB.txt"という形式で
//  フォルダ名\ファイル名、という形式で文字列を取り出す
備考:
履歴:       2005/11/08
//--▼----------------------▽--}
function ExtractFolderFileName(Path: String): String;
begin
  Result :=
    IncludeTrailingPathDelimiter(ExtractFileName(ExtractFileDir(Path))) +
    ExtractFileName(Path);
end;
//--△----------------------▲--

{-------------------------------
//  ファイルパスの階層を調べる関数
機能:
備考:
履歴:       2006/12/24(日) 00:33
//--▼----------------------▽--}
function PathLevel(Path: String): Integer;
begin
  Result := -1;
  if CheckDrivePath(Path) then
  begin
    Result := WordCount(Path, [PathDelim], dmUserFriendly)-1;
  end else
  if CheckUNCPath(Path) then
  begin
    Result := WordCount(Path, [PathDelim], dmUserFriendly)-1;
  end else
  begin
    Exit;
  end;
end;
//--△----------------------▲--

{-------------------------------
//  ファイルパスの階層を切り取る関数
機能:       指定した階層(CutCount)分にフォルダを上に上がる関数
            戻り値の最後には[\]記号がつく
備考:
履歴:       2006/12/24(日) 00:33
//--▼----------------------▽--}
function CutPathLevel(Path: String; CutCount: Integer): String;
var
  Level: Integer;
  i: Integer;
begin
  Result := '';
  if CutCount < 0 then Exit;
  if CutCount = 0 then
  begin
    Result := Path;
    Exit;
  end;

  Level := PathLevel(Path);
  if Level = -1 then Exit;

  if Level < CutCount then
  begin
    Exit;
  end else
  begin
    for i := 0 to Level - CutCount do
    begin
      Result := Result +
        WordGet(Path, [PathDelim], i, dmUserFriendly) + PathDelim;
    end;
    if CheckUNCPath(Path) then
    begin
      Result := StringOfChar(PathDelim, 2) + Result;
    end;
  end;

end;
//--△----------------------▲--

{---------------------------------------
    ひとつ上のフォルダの同名ファイルパスを得る
機能:   
備考:   
履歴:   2011/07/15(金)
        ・  作成
}//(*-----------------------------------
function GetUpFolderPath(FilePath: string): String;
begin
  Result := ExtractFilePath( ExtractFileDir(FilePath) ) +
            ExtractFileName(FilePath)
end;
//------------------------------------*)



{-------------------------------
//  拡張子の一致を調べる関数
機能:       Extはピリオドを先頭に付けても付けなくてもよい
備考:       ".pas"や".pas.bak"というような拡張子で一致を調べる事ができる
履歴:       2007/08/30(木) 15:09
//--▼----------------------▽--}
function CheckSameExt(const Filename, Ext: String): Boolean;
begin
//  Result := StringLastCompareCase(
//       IncludeFirstStr(Ext, '.'), Filename, ccIgnoreCase);
  Result := IsLastStr(Filename, IncludeFirstStr(Ext, '.'), ccIgnoreCase);
end;
//--△----------------------▲--

{----------------------------------------
//      絶対パスを取り出す関数
        GetAbsolutePathFromRelativePath
機能:       BasePath:基準のパス(ドライブorUNC)
            RelativePath:相対パス(..\や.\やABC等)
            戻り値はFullPath
備考:
BasePathはC:\か\\である必要あり
RelativePathはパスである必要あり
(ピリオドでスタートしてなくてもいい)

SetCurrentDirとExpandFileNameの場合、
存在しないフォルダの場合困る。
相対パスを絶対パスに変換する
http://www.wwlnk.com/boheme/delphi/tips/tec1600.htm

自作モノもある。
サンプル: "相対パスから絶対パスへ変換"
C:/Software/FirefoxPortable/Data/profile/ScrapBook/data/20100218100202/index.html

      ・\記号で分解する(空文字含む分解)
      ・要素に空文字がきたらエラー
        >>  CheckPathFollowRuleでチェック済み
      ・要素が
        [..]単独なら一つ上
        [.]単独なら処理なし
        [...]以上ならエラー
        [.x][..x][...x]などというファイルはエラー(存在しない)
        [A..]このようなファイルは存在する
        >>  CheckPathFollowRuleでチェック済み
      ・元のパスが[C:\TEST1\TEST2\]と[C:\TEST1\TEST2]は
        最後の[\]の有無だけなので区別はしない
      ・最後の[\]を除外して、[\]で分解
        [C:][TEST1][TEST2]となるので、
        それを相対パスで[..]が来た場合削っていく
履歴:
        2010/03/09(火)
        ・  フォルダ作成してSetCurrentDirで移動し
            ExpandFileNameをするという仕様を変更して
            全て自作化した。テストも通過している。
//----------------------------------------}
function GetAbsolutePathFromRelativePath(BasePath, RelativePath: WideString): WideString;
var
  WordSplited: TStringSplitter;
  i, j: Integer;
  LastPathDelimFlag: Boolean;
begin
  Result := '';
  {↓ドライブパスでもなく、UNCパスでもない}
  if (not CheckDrivePath(BasePath)) and (not CheckUNCPath(BasePath)) then
  begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
    Exit;
  end;
  if not CheckPathFollowRule(RelativePath) then
  begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
    Exit;
  end;
      {
      }
  BasePath := ExcludeLastPathDelim(BasePath);
  if IsLastStr(RelativePath, PathDelim) then
    LastPathDelimFlag := True
  else
    LastPathDelimFlag := False;

  WordSplited := TStringSplitter.Create(ExcludeLastPathDelim(RelativePath),
    [PathDelim], [sfInEmptyStr]); try
  if WordSplited.Count = 0 then
  begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
    Exit;
  end;

  j := 0;
  for i := 0 to WordSplited.Count - 1 do
  begin
    {↓このチェックはBasePathに対して上記でやっているので不要
    if WordSplited.Words[i] = EmptyStr then
    begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
      Exit;
    end;}

    if WordSplited.Words[i] = '..' then
    begin
      {↓相対パスが上にたどりすぎている場合は不適合}
      if BasePath = '' then
      begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
        Exit;
      end;
          
      BasePath := CutPathLevel(BasePath, 1);
      j := i + 1;
    end else
    if WordSplited.Words[i] = '.' then
    begin
      j := i + 1;
    end else
    begin
      j := i;
      break;
    end;
  end;

  if BasePath  <> '' then
    Result := IncludeLastPathDelim(BasePath);

  for i := j to WordSplited.Count - 1 do
  begin
    Result := Result + WordSplited.Words[i] + PathDelim;
  end;

  if LastPathDelimFlag then
    Result := IncludeLastPathDelim(Result)
  else
    Result := ExcludeLastPathDelim(Result);

  {↓ドライブパスでもなく、UNCパスでもない}
  if (not CheckDrivePath(Result)) and (not CheckUNCPath(Result)) then
  begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
    Result := '';
    Exit;
  end;
  finally WordSplited.Free; end;
end;
//----------------------------------------

{----------------------------------------
//      相対パスを取り出す関数
		GetRelativePathFromAbsolutePath
機能:       BasePath:基準のパス(ドライブorUNC)
			AbsolutePath:絶対パス(ドライブorUNC)
			戻り値はRelativePath
備考:       
ExtractRelativePathでいいようだ。
二つのパスから相対パスを生成する
http://www.wwlnk.com/boheme/delphi/tips/tec1590.htm
履歴:       
//----------------------------------------}
function GetRelativePathFromAbsolutePath(BasePath, AbsolutePath: WideString): WideString;
var
  UpPathString: WideString;
  i: Integer;
  BasePathLevel: Integer;
begin
  Result := '';
  if (not CheckDrivePath(BasePath)) and (not CheckUNCPath(BasePath)) then
  begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
    Exit;
  end;
  if (not CheckDrivePath(AbsolutePath)) and (not CheckUNCPath(AbsolutePath)) then 
  begin
//    Exception.Create('指定された文字列はファイルパスに適合しません:'+BasePath);
    Exit;
  end;

  BasePath := ExcludeLastPathDelim(BasePath);
  if IsFirstStr(AbsolutePath, BasePath, ccIgnoreCase) then
  begin
    {先頭だけ文字列置き換え>>[.]にする}
    Result := StringsReplace(AbsolutePath, [BasePath], ['.'], ccIgnoreCase, False);
    Exit;
  end;

  UpPathString := '';
  BasePathLevel := PathLevel(BasePath);
  for i := 1 to BasePathLevel do
  begin
    UpPathString := UpPathString + '..' + PathDelim;
    BasePath := ExcludeLastPathDelim(CutPathLevel(BasePath, 1));
    if IsFirstStr(AbsolutePath, BasePath, ccIgnoreCase) then
    begin
      Result := StringsReplace(AbsolutePath, [BasePath],
        [ExcludeLastPathDelim(UpPathString)], ccIgnoreCase, False);
      Exit;
    end;
  end;

{TODO:
・  BasePathのパスを上の方にたどる
・  先頭がAbsolutePathと一致するまで
    Resultに[..\]を追加する。
・  一致すれば、一致した部分をResultと置き換える
}

//  BasePath := IncludeLastPathDelim(BasePath);
//  Result := ExtractRelativePath(BasePath, AbsolutePath);
end;
//----------------------------------------
//没
//function GetRelativePathFromAbsolutePath(BasePath, AbsolutePath: WideString): WideString;
//begin
//  Result := '';
//  if (not CheckDrivePath(BasePath)) and (not CheckUNCPath(BasePath)) then Exit;
//  if (not CheckDrivePath(AbsolutePath)) and (not CheckUNCPath(AbsolutePath)) then Exit;
//
//  BasePath := IncludeLastPathDelim(BasePath);
//  Result := ExtractRelativePath(BasePath, AbsolutePath);
//end;

////    BaseName : ベースとなる絶対パス
////    SrcName  : 変換したい相対パス
//function ExtractFullPath(const BaseName, SrcName: string): string;

//  function AddFirstSlash(const Str: string): string;
//  begin
//    Result := Str;
//    if Copy(Result, 1, 1)<>'\' then
//      Result := '\' + Result;
//  end;

//  function DeleteFirstSlash(const Str: string): string;
//  begin
//    Result := Str;
//    if Copy(Result, 1, 1)='\' then
//      Result := Copy(Result, 2, MaxInt);
//  end;

//  function AddBackSlash(const Str: string): string;
//  begin
//    Result := Str;
//    if (Result<>'') and (Result[Length(Result)]<>'\') then
//      Result := Result + '\';
//  end;

//  function ExtractFullPath2(const RelativePath: string): string;
//  var
//    SrcPath: string;
//    SrcDirs: array[0..129] of PChar;
//    SrcDirCount: Integer;
//
//    procedure SplitDirs(var Path: string; var Dirs: array of PChar;
//      var DirCount: Integer);
//    var
//      I, J: Integer;
//    begin
//      I := 1;
//      J := 0;
//      while I <= Length(Path) do
//      begin
//        if Path[I] in LeadBytes then Inc(I)
//        else if Path[I] = '\' then             { Do not localize }
//        begin
//          Path[I] := #0;
//          Dirs[J] := @Path[I + 1];
//          Inc(J);
//        end;
//        Inc(I);
//      end;
//      DirCount := J - 1;
//    end;
//
//  var
//    i: Integer;
//    DriveName: string;
//  begin
//    Result := '';
//
//    DriveName := ExtractFileDrive(RelativePath);
//    SrcPath := Copy(RelativePath, Length(DriveName)+1, MaxInt);
//    SplitDirs(SrcPath, SrcDirs, SrcDirCount);
//
//    for i:=0 to SrcDirCount do
//    begin
//      if SrcDirs[i]='.' then
//      else if SrcDirs[i]='..' then
//        Result := ExtractFileDir(Result)
//      else begin 
//        Result := AddBackSlash(Result) + SrcDirs[i];
//      end;
//    end;
//
//    Result := DriveName + AddFirstSlash(Result);
//  end;
//
//begin
//  if ExtractFileDrive(SrcName)<>'' then
//  begin
//    {  SrcName にドライブ名が指定されているので BaseName は無視  }
//    Result := AddBackSlash(ExtractFullPath2(SrcName));
//  end else
//  begin
//    Result := ExtractFullPath2(AddBackSlash(BaseName)+
//      DeleteFirstSlash(SrcName));
//  end;
//end;

//end;

{---------------------------------------
    環境変数を取得する関数
機能:   
備考:   コマンドラインで
		set appdata とか
		set homepat などと入力すると環境変数が確認できる
履歴:   2012/01/15(日)
        ・  作成済み
}//(*-----------------------------------
//環境変数の一覧を'APPDATA=C:…AAA'という形式の文字列配列で取得します。
function GetEnvironmentStringArray: TStringDynArray;
var
  Pointer: PChar;
  s: String;
begin
  SetLength(Result, 0);
  s := '';
  Pointer := GetEnvironmentStrings;
  while True do
  begin
    if (Pointer^ = #0) and ((Pointer+1)^ = #0) then
    begin
      break;
    end else
    if Pointer^ = #0 then
    begin
      SetLength(Result, Length(Result)+1);
      Result[Length(Result)-1] := s;
      Inc(Pointer);
      s := '';
    end else
    begin
      s := s + Pointer^;
      Inc(Pointer);
    end;
  end;
end;

//環境変数に対して%APPDATA%XXX、と指定したパスを渡すと
//フルパスが返る関数
function ReplaceEnvironmentVariableNameToValue(Source: string): String;
var
  EnvironmentStringArray: TStringDynArray;
  I, J: Integer;
  OldStrs, NewStrs: TStringDynArray;
begin
  EnvironmentStringArray := GetEnvironmentStringArray;
  SetLength(OldStrs, Length(EnvironmentStringArray));
  SetLength(NewStrs, Length(EnvironmentStringArray));
  J := 0;
  for I := 0 to Length(EnvironmentStringArray) - 1 do
  begin
    if IsFirstStr(EnvironmentStringArray[I], '=') then Continue;
    OldStrs[J] := IncludeBothEndsStr(
      FirstString(EnvironmentStringArray[I], '='), '%');
    NewStrs[J] := LastString(EnvironmentStringArray[I], '=');
    Inc(J);
  end;
  SetLength(OldStrs, J + 1);
  SetLength(NewStrs, J + 1);

  Result := StringsReplace(Source, OldStrs, NewStrs, ccIgnoreCase);
end;
//------------------------------------*)

end.
