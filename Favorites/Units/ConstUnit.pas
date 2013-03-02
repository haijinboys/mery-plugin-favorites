unit ConstUnit;

interface

const
  EmptyStr = '';
  CR = #$D;     //#13
  LF = #$A;     //#10
  TAB= #$9;     //#9
  NullStr = #$0;//#0
  EN = '\';
  CRLF = #$D#$A;//#13#10
  Space = ' ';
  ZenkakuSpace = '　';

  SingleQuote = '''';
  DoubleQuote = '"';

  Comma = ',';
  Colon = ':';
  SemiColon = ';';
  Period = '.';

const
  VK_ALT = 18;      //VK_MENU = 18;
  VK_RALT = 165;    //VK_RMENU = 165;
  VK_WIN = 91;      //VK_LWIN;
  VK_PAGEUP = 33;   //VK_PRIOR = 33;
  VK_PAGEDOWN = 34; //VK_NEXT = 34;

const zenHiraTbl: String =      // 全角ひらかな
    'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほ'+
    'まみむめもやゆよらりるれろわゐゑをんがぎぐげござじずぜぞだぢづでど'+
    'ばびぶべぼぱぴぷぺぽぁぃぅぇぉゃゅょゎっ';
const zenKataTbl: String =      // 全角カタカナ
    'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホ'+
    'マミムメモヤユヨラリルレロワンヲガギグゲゴザジズゼゾダヂヅデド'+
    'バビブベボパピプペポァィゥェォャュョヮッー';
const hanKanaTbl: String =      // 半角カタカナ
    'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｬｭｮｯﾟｰ･､｡｢｣';
    //※メーラーだと半角カタカナは全角に変換されてしまいます
const zenNumberTbl: String =    // 全角数字
    '１２３４５６７８９０';
const hanNumberTbl: String =    // 半角数字
    '1234567890';
const zenAlphaTbl: String =     // 全角英文字
    'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'+
    'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ';
const hanAlphaTbl: String =     // 半角英文字
    'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
const zenMarkTbl: String =      // 全角記号(もっと沢山あるけど一部のみ)
    '－．＆％＄＃＊＠？／，＜＞（）｛｝｜；：～＋｀！＇＝＿＾・';
const hanMarkTbl: String =      // 半角記号
    '!"#$%&''()*+,-./:;<=>?@[\]^_`{|}~';
const zenSpaceTbl: String = '　';   // 全角スペース
const hanSpaceTbl: String = ' ';    // 半角スペース

const CtrlCharTbl: String =     //制御文字
    #$01#$02#$03#$04#$05#$06#$07#$08#$09#$0A#$0B#$0C#$0D#$0E#$0F+
    #$10#$11#$12#$13#$14#$15#$16#$17#$18#$19#$1A#$1B#$1C#$1D#$1E#$1F;


implementation

end.
