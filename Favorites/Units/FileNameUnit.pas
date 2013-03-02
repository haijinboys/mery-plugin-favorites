{ --��---------------------------��--
�����t�@�C�����A�Z���t�@�C�������ݕϊ��֐�
2002/06/05
  �쐬����Ă���
2003/03/17
  �֐��̐������L�q�A�S�̓I�ɃR�����g�Ȃǂ��C��
  GetShortFullPathName/GetLongFullPathName��
  �s�v�Ȃ̂�implementation�����Ɏ�������
  interface�ɂ�GetLongFileName/GetShortFileName��
  ���������B
2005/12/19
  LongShortFileName.pas����FileNameUnit.pas�ɖ��O�ύX
2005/12/29
  CheckUNCPath��CheckDrivePath�𐳂�������
2006/12/24(��)
�ECheckDrivePath���f�o�b�O
�EPathLevel/CutPathLevel������
�EtestTrailingPathDelimiter��ǉ�
2007/07/22(��) 00:04
�EFileUnit.pas����ExtractFolderFileName�֐����ړ�
2010/03/10(��)
�EGetAbsolutePathFromRelativePath/GetRelativePathFromAbsolutePath
  �����������B
2010/11/10
�ERepaceFlag��ccIgnoreCase/ccCaseSensitive�ɕύX�����B
2011/05/23(��)
�EGetEnvironmentStringArray��ǉ������B
2012/01/15(��)
�ECheckCanRelativePath��ǉ�
  ���̂��߂ɁACheckUNCPathHeader��CheckDrivePathHeader���쐬����
  CheckUNCPath��CheckDrivePath�����t�@�N�^�����O
//--��---------------------------��-- }
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
//�����O�t�@�C����/�V���[�g�t�@�C�������擾����
    GetLongFileName
    GetShortFileName
���l:       DelphiML�������[�h�u�����O�t�@�C���v
            2ch�̏��
            ���ǁAFDelphi�̃R�[�h��肱����̕���
            �Z���Ă悢���낤
����:       2003/03/17
              ���݂��Ȃ��t�@�C����ϊ������ꍇ
              �󕶎���Ԃ��悤�ɏC��
//--��----------------------��--}
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
//--��----------------------��--


{-------------------------------
//�t�@�C�������r����
�@�\:       �����񂪓���̃t�@�C���������Ă���̂��ǂ����𔻒f���܂��B
            ���݂���t�@�C���������ꂩ�ǂ������f�ł��܂���B
            A,B�͒Z���t�@�C�����ł������t�@�C�����ł�OK�ł��B
�߂�l:     true:�����t�@�C�� false:�قȂ�t�@�C��
���l:       ���j�b�gSysUtils��SameFileName�֐��ő�p�ł��邩������܂���
����:       2002/02/11
//--��----------------------��--}
function CheckSameFileName(A, B: String): Boolean;
var
  OldMode: UINT;// �G���[���[�h�ێ��p
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
//--��----------------------��--



{-------------------------------
//  UNC(=Universal Naming Convention)�p�X���ǂ����𔻒f����֐�
�@�\:       ������3�����ȏ�ł�����
            �擪���w\\?�x���Č`�ɂȂ��Ă���
            �r���Ńt�@�C���֎~������w\\�x���Ȃ�
            ���ǂ����𔻒f����֐�
���l:
����:       2005/12/28
2012/01/11(��)
�E  �C��
//--��----------------------��--}
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

//--��----------------------��--

{-------------------------------
//  �ʏ�̃h���C�u�p�X���ǂ����𔻒f����֐�
�@�\:       ������2�����ȏ�ł�����
            �擪���wX:�x���Č`�ɂȂ��Ă���
            �r���Ńt�@�C���֎~������w\\�x���Ȃ�
            ���ǂ����𔻒f����֐�
���l:
����:       2005/12/28
            2006/12/24(��) �듮����኱�C��
//--��----------------------��--}

  function CheckDrivePathHeader(Path: string): Boolean;
  begin
    Result := False;
    if Length(Path)<=1 then Exit;
    {���ꕶ���̏ꍇ�͂ǂ�����Ă��p�X�ɂȂ�Ȃ�}

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
//--��----------------------��--


{---------------------------------------
    ���΃p�X�ɂȂ�镶���񂩂ǂ����m�F����֐�
�@�\:   
���l:   
����:   2012/01/15(��)
        �E  �쐬
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
//  CheckDrivePath��CheckUNCPath�Ŏg��������֐�
�@�\:       ������1���������Ȃ�������
            UNC�w��ȊO�̏ꏊ(�܂�1�����ڈȊO)��"\\"���g���Ă�����
            ���̑��֎~�������g���Ă����肵��
            Path���t�@�C���p�X�Ƃ��Đ��藧���ǂ����𔻒f����֐�
�d�l:
�E�p�X�Ƃ��Đ��������[���Ɋ܂܂�Ă��鎞��True
  �������Ȃ��Ȃ�False

  �E\�L���ɂ���ċ󕶎����܂ޕ������s���B�����L���͊܂܂Ȃ��B
  �E�擪��\\�ƘA�����Ă���ꍇ
    �܂�A�󕶎��v�f���擪��2�A���ő����̂�OK
    1�����̋󕶎��v�f�Ȃ�NotOK
  �E[:]���擪�v�f��2�����ڈȊO�ɂ���ƃG���[
    >>�ł��u..\D:\�Ƃ����ꍇ������ȁc�v>>�Ή�����
  �E��؂�ꂽ�v�f�ɋ󕶎�������ꍇ�̓G���[
    (�܂�\\�L�����������ꍇ�̓G���[)
  �E��؂�ꂽ�v�f��[...]�Ŏn�܂���e������ƃG���[
  �E��؂�ꂽ�v�f��[/*?"<>|]���܂܂�Ă���ƃG���[
���l:
����:       2005/12/29
            2011/05/23(��)
            �E  �d�l���܂Ƃ߂��BUnicode�łœ���m�F
//--��----------------------��--}

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

  {��\�ŋ�؂����Ƃ��̋󍀖ځA�܂�\���A���ŕ���ł���ꍇ��
     ���e�����̂�[\\abc]�ƂȂ��Ă��鎞����
     ����ȊO��\���A���Ƃ݂Ȃ����̂�
     �t�@�C���p�X�ł͂Ȃ�}
  SearchStartIndex := 0;
  if (2 <= Split.Count) and (Split.Words[0] = EmptyStr) then
  begin
    {����Words[0]��EmptyStr��Words[1]��EmptyStr�ł͂Ȃ��Ȃ�
       [\]���擪��1�����̏ꍇ������
       �t�@�C���p�X�ł͂Ȃ�}
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
        {���擪���ڂ���Ȃ��̂�[.]��[..]�Ȃ�
           �t�@�C���p�X�ł͂Ȃ�(���Ƃɂ���)}
        Result := False;
        Exit;
      end;
    end;

    if FirstItemFlag then
    begin
      FirstItemFlag := False;
      {���ŏ��̍��ڂ�[:]���܂܂�Ă���
         ���ꂪ2�����ڂł͂Ȃ��ꍇ�̓t�@�C���p�X�ł͂Ȃ�}
      if InStr(DriveDelim, Split.Words[i]) then
        if not InStr(DriveDelim, Split.Words[i], 2, 1) then
        begin
          Result := False;
          Exit;
        end else
        begin
          {��2�����ڂ�[:]�ł�1�����ڂ��A���t�@�x�b�g�ł͂Ȃ��Ȃ�
             �t�@�C���p�X�ł͂Ȃ�}
          if not (CheckStrInTable(Split.Words[i][1], hanAlphaTbl)=itAllInclude) then
          begin
            Result := False;
            Exit;
          end;
        end;
    end else
    begin
      {���ŏ��̍��ڈȊO��[:]���܂܂�Ă���Ȃ�
         �t�@�C���p�X�ł͂Ȃ�}
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

    {��[.]��[..]�ȊO�Ő擪��[.]������ꍇ��
       �t�@�C���p�X�ł͂Ȃ�
       [A...]�Ƃ������ڂ�OK}
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

//--��----------------------��--

{-------------------------------
//  ExtractFileNameExcludeExt
�@�\:       �g���q���f�B���N�g���p�X���Ȃ�
            �t�@�C�������擾����
���l:       �t�@�C�������������Ȃ��ꍇ
            �߂�l��ChangeFileExt��ExtractFileName�̏����Ɉˑ�����
����:       2006/02/18
//--��----------------------��--}
function ExtractFileNameExcludeExt(const FileName: String): String;
begin
  Result := ChangeFileExt(ExtractFileName(FileName), '');
end;
//--��----------------------��--

{-------------------------------
//  �p�X�̖�������"AAA\BBB.txt"�Ƃ����`����
//  �t�H���_��\�t�@�C�����A�Ƃ����`���ŕ���������o��
���l:
����:       2005/11/08
//--��----------------------��--}
function ExtractFolderFileName(Path: String): String;
begin
  Result :=
    IncludeTrailingPathDelimiter(ExtractFileName(ExtractFileDir(Path))) +
    ExtractFileName(Path);
end;
//--��----------------------��--

{-------------------------------
//  �t�@�C���p�X�̊K�w�𒲂ׂ�֐�
�@�\:
���l:
����:       2006/12/24(��) 00:33
//--��----------------------��--}
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
//--��----------------------��--

{-------------------------------
//  �t�@�C���p�X�̊K�w��؂���֐�
�@�\:       �w�肵���K�w(CutCount)���Ƀt�H���_����ɏオ��֐�
            �߂�l�̍Ō�ɂ�[\]�L������
���l:
����:       2006/12/24(��) 00:33
//--��----------------------��--}
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
//--��----------------------��--

{---------------------------------------
    �ЂƂ�̃t�H���_�̓����t�@�C���p�X�𓾂�
�@�\:   
���l:   
����:   2011/07/15(��)
        �E  �쐬
}//(*-----------------------------------
function GetUpFolderPath(FilePath: string): String;
begin
  Result := ExtractFilePath( ExtractFileDir(FilePath) ) +
            ExtractFileName(FilePath)
end;
//------------------------------------*)



{-------------------------------
//  �g���q�̈�v�𒲂ׂ�֐�
�@�\:       Ext�̓s���I�h��擪�ɕt���Ă��t���Ȃ��Ă��悢
���l:       ".pas"��".pas.bak"�Ƃ����悤�Ȋg���q�ň�v�𒲂ׂ鎖���ł���
����:       2007/08/30(��) 15:09
//--��----------------------��--}
function CheckSameExt(const Filename, Ext: String): Boolean;
begin
//  Result := StringLastCompareCase(
//       IncludeFirstStr(Ext, '.'), Filename, ccIgnoreCase);
  Result := IsLastStr(Filename, IncludeFirstStr(Ext, '.'), ccIgnoreCase);
end;
//--��----------------------��--

{----------------------------------------
//      ��΃p�X�����o���֐�
        GetAbsolutePathFromRelativePath
�@�\:       BasePath:��̃p�X(�h���C�uorUNC)
            RelativePath:���΃p�X(..\��.\��ABC��)
            �߂�l��FullPath
���l:
BasePath��C:\��\\�ł���K�v����
RelativePath�̓p�X�ł���K�v����
(�s���I�h�ŃX�^�[�g���ĂȂ��Ă�����)

SetCurrentDir��ExpandFileName�̏ꍇ�A
���݂��Ȃ��t�H���_�̏ꍇ����B
���΃p�X���΃p�X�ɕϊ�����
http://www.wwlnk.com/boheme/delphi/tips/tec1600.htm

���샂�m������B
�T���v��: "���΃p�X�����΃p�X�֕ϊ�"
C:/Software/FirefoxPortable/Data/profile/ScrapBook/data/20100218100202/index.html

      �E\�L���ŕ�������(�󕶎��܂ޕ���)
      �E�v�f�ɋ󕶎���������G���[
        >>  CheckPathFollowRule�Ń`�F�b�N�ς�
      �E�v�f��
        [..]�P�ƂȂ���
        [.]�P�ƂȂ珈���Ȃ�
        [...]�ȏ�Ȃ�G���[
        [.x][..x][...x]�ȂǂƂ����t�@�C���̓G���[(���݂��Ȃ�)
        [A..]���̂悤�ȃt�@�C���͑��݂���
        >>  CheckPathFollowRule�Ń`�F�b�N�ς�
      �E���̃p�X��[C:\TEST1\TEST2\]��[C:\TEST1\TEST2]��
        �Ō��[\]�̗L�������Ȃ̂ŋ�ʂ͂��Ȃ�
      �E�Ō��[\]�����O���āA[\]�ŕ���
        [C:][TEST1][TEST2]�ƂȂ�̂ŁA
        ����𑊑΃p�X��[..]�������ꍇ����Ă���
����:
        2010/03/09(��)
        �E  �t�H���_�쐬����SetCurrentDir�ňړ���
            ExpandFileName������Ƃ����d�l��ύX����
            �S�Ď��쉻�����B�e�X�g���ʉ߂��Ă���B
//----------------------------------------}
function GetAbsolutePathFromRelativePath(BasePath, RelativePath: WideString): WideString;
var
  WordSplited: TStringSplitter;
  i, j: Integer;
  LastPathDelimFlag: Boolean;
begin
  Result := '';
  {���h���C�u�p�X�ł��Ȃ��AUNC�p�X�ł��Ȃ�}
  if (not CheckDrivePath(BasePath)) and (not CheckUNCPath(BasePath)) then
  begin
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
    Exit;
  end;
  if not CheckPathFollowRule(RelativePath) then
  begin
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
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
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
    Exit;
  end;

  j := 0;
  for i := 0 to WordSplited.Count - 1 do
  begin
    {�����̃`�F�b�N��BasePath�ɑ΂��ď�L�ł���Ă���̂ŕs�v
    if WordSplited.Words[i] = EmptyStr then
    begin
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
      Exit;
    end;}

    if WordSplited.Words[i] = '..' then
    begin
      {�����΃p�X����ɂ��ǂ肷���Ă���ꍇ�͕s�K��}
      if BasePath = '' then
      begin
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
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

  {���h���C�u�p�X�ł��Ȃ��AUNC�p�X�ł��Ȃ�}
  if (not CheckDrivePath(Result)) and (not CheckUNCPath(Result)) then
  begin
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
    Result := '';
    Exit;
  end;
  finally WordSplited.Free; end;
end;
//----------------------------------------

{----------------------------------------
//      ���΃p�X�����o���֐�
		GetRelativePathFromAbsolutePath
�@�\:       BasePath:��̃p�X(�h���C�uorUNC)
			AbsolutePath:��΃p�X(�h���C�uorUNC)
			�߂�l��RelativePath
���l:       
ExtractRelativePath�ł����悤���B
��̃p�X���瑊�΃p�X�𐶐�����
http://www.wwlnk.com/boheme/delphi/tips/tec1590.htm
����:       
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
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
    Exit;
  end;
  if (not CheckDrivePath(AbsolutePath)) and (not CheckUNCPath(AbsolutePath)) then 
  begin
//    Exception.Create('�w�肳�ꂽ������̓t�@�C���p�X�ɓK�����܂���:'+BasePath);
    Exit;
  end;

  BasePath := ExcludeLastPathDelim(BasePath);
  if IsFirstStr(AbsolutePath, BasePath, ccIgnoreCase) then
  begin
    {�擪����������u������>>[.]�ɂ���}
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
�E  BasePath�̃p�X����̕��ɂ��ǂ�
�E  �擪��AbsolutePath�ƈ�v����܂�
    Result��[..\]��ǉ�����B
�E  ��v����΁A��v����������Result�ƒu��������
}

//  BasePath := IncludeLastPathDelim(BasePath);
//  Result := ExtractRelativePath(BasePath, AbsolutePath);
end;
//----------------------------------------
//�v
//function GetRelativePathFromAbsolutePath(BasePath, AbsolutePath: WideString): WideString;
//begin
//  Result := '';
//  if (not CheckDrivePath(BasePath)) and (not CheckUNCPath(BasePath)) then Exit;
//  if (not CheckDrivePath(AbsolutePath)) and (not CheckUNCPath(AbsolutePath)) then Exit;
//
//  BasePath := IncludeLastPathDelim(BasePath);
//  Result := ExtractRelativePath(BasePath, AbsolutePath);
//end;

////    BaseName : �x�[�X�ƂȂ��΃p�X
////    SrcName  : �ϊ����������΃p�X
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
//    {  SrcName �Ƀh���C�u�����w�肳��Ă���̂� BaseName �͖���  }
//    Result := AddBackSlash(ExtractFullPath2(SrcName));
//  end else
//  begin
//    Result := ExtractFullPath2(AddBackSlash(BaseName)+
//      DeleteFirstSlash(SrcName));
//  end;
//end;

//end;

{---------------------------------------
    ���ϐ����擾����֐�
�@�\:   
���l:   �R�}���h���C����
		set appdata �Ƃ�
		set homepat �ȂǂƓ��͂���Ɗ��ϐ����m�F�ł���
����:   2012/01/15(��)
        �E  �쐬�ς�
}//(*-----------------------------------
//���ϐ��̈ꗗ��'APPDATA=C:�cAAA'�Ƃ����`���̕�����z��Ŏ擾���܂��B
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

//���ϐ��ɑ΂���%APPDATA%XXX�A�Ǝw�肵���p�X��n����
//�t���p�X���Ԃ�֐�
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
