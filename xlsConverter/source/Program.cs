// テスト用
// #define USER_NEW_INIT_DATA


using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Diagnostics;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;



namespace xlsConverter
{
    /*!
     * 解析関係のデータ.
     */
    public class AnalyseData
    {
    #region メンバ
        // 基本.
        private string _eod;
        private int _line;
        private int _key;

        private string _ext_column;
    #endregion

    #region プロパティ

        [XmlElement("eod")]
        public string eod
        {
            get { return _eod; }
            set { _eod = value; }
        }

        [XmlElement("line")]
        public int line
        {
            get { return _line; }
            set { _line = value; }
        }

        [XmlElement("key")]
        public int key
        {
            get { return _key; }
            set { _key = value; }
        }

        [XmlElement("ext_column")]
        public string ext_column
        {
            get { return _ext_column; }
            set { _ext_column = value; }
        }
    #endregion
    };

    /*!
     * C/C++ Headerファイル出力用定義
     */
    public class HeaderData
    {
        #region メンバ
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "_header";
        [XmlElement("ext")]
        public string ext = ".h";
        [XmlElement("type")]
        public string type = "c";
        #endregion
    };

    /*!
     * Xml ファイル出力用定義
     */
    public class XmlData
    {
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "";
        [XmlElement("ext")]
        public string ext = ".xml";
    }

    /*!
     * Json ファイル出力用定義
     */
    public class JsonData
    {
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "";
        [XmlElement("ext")]
        public string ext = ".json";
    }

    /*!
     * Csv ファイル出力用定義
     */
    public class CsvData
    {
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "";
        [XmlElement("ext")]
        public string ext = ".csv";
    }

    /*!
     * MessagePack ファイル出力用定義
     */
    public class MessagePackData
    {
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "";
        [XmlElement("ext")]
        public string ext = ".messagepack";
    }

    /*!
     * Valb ファイル出力用定義
     */
    public class ValbData
    {
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "";
        [XmlElement("ext")]
        public string ext = ".valb";
    }

    /*!
     * Strb ファイル出力用定義
     */
    public class StrbData
    {
        [XmlElement("code")]
        public string code = "UTF-8";
        [XmlElement("language")]
        public string language;
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "";
        [XmlElement("ext")]
        public string ext = ".strb";
    }
    
    /*!
     * StringXml ファイル出力用定義
     */
    public class StringXmlData
    {
        [XmlElement("hash")]
        public string hash;
        [XmlElement("language")]
        public string language;
        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "_string";
        [XmlElement("ext")]
        public string ext = ".xml";
    }

    /*!
     * Htb ファイル出力用定義
     */
    public class HtbData
    {
        [XmlElement("hash")]
        public string hash;
        [XmlElement("comment")]
        public string comment;
        [XmlElement("bit")]
        public int bit = 32;
        [XmlElement("empty")]
        public bool empty = false;

        [XmlElement("binaly_header")]
        public string binaly_header = "";
        [XmlElement("binaly_footer")]
        public string binaly_footer = "";
        [XmlElement("binaly_ext")]
        public string binaly_ext = ".htb";

        [XmlElement("source_header")]
        public string source_header = "";
        [XmlElement("source_footer")]
        public string source_footer = "_Hash";
        [XmlElement("source_ext")]
        public string source_ext = ".h";
    }

    /*!
     * Enum ファイル出力用定義
     */
    public class EnumData
    {
        [XmlElement("type")]
        public string type;
        [XmlElement("define")]
        public string define;
        [XmlElement("comment")]
        public string comment;
        [XmlElement("id")]
        public string id;

        [XmlElement("header")]
        public string header = "";
        [XmlElement("footer")]
        public string footer = "";
        [XmlElement("ext")]
        public string ext = ".h";
    };

    /*
     * ファイル出力用定義
     */
    public class PublishData
    {
        public PublishData()
        {
            header = new HeaderData();
            json = new JsonData();
            csv = new CsvData();
            messagepack = new MessagePackData();
            valb = new ValbData();
            strb = new StrbData();
            htb = new HtbData();
            stringxml = new StringXmlData();
            enumdata = new EnumData();
        }
        #region メンバ
        [XmlElement("alignment")]
        public int alignment = 4;
        [XmlElement("comment_tab")]
        public int commentTab = 8;
        [XmlElement("Header")]
        public HeaderData header;
        [XmlElement("Xml")]
        public XmlData xml;
        [XmlElement("Json")]
        public JsonData json;
        [XmlElement("Csv")]
        public CsvData csv;
        [XmlElement("MessagePack")]
        public MessagePackData messagepack;
        [XmlElement("Valb")]
        public ValbData valb;
        [XmlElement("Strb")]
        public StrbData strb;
        [XmlElement("Htb")]
        public HtbData htb;
        [XmlElement("StringXml")]
        public StringXmlData stringxml;
        [XmlElement("Enum")]
        public EnumData enumdata;
        #endregion

    };


    /*!
     * 初期化用xml読み込みクラス
     */
    [XmlRoot("InitData")]
    public class InitData
    {
        public InitData()
        {
            _analyse = new AnalyseData();
            _publish = new PublishData();
        }

        #region メンバ
        // 管理
        // 入力用
        private AnalyseData _analyse;   // xlsの解析用
        // 出力用
        private PublishData _publish;       // c/c++ Header 用
        #endregion

        #region メンバ
        [XmlElement("Analyse")]
        public AnalyseData analyse
        {
            get { return _analyse; }
            set { _analyse = value; }
        }
        [XmlElement("Publish")]
        public PublishData publish
        {
            get { return _publish; }
            set { _publish = value; }
        }
        #endregion
    };

    static class Program
    {
        enum Option : int 
        {
            OPTION_HEADER = 0,
//            OPTION_HEADER_XML,
            OPTION_XML,
            OPTION_JSON,
            OPTION_CSV,
            OPTION_MSGP,
            OPTION_VALB,
            OPTION_STRB,
            OPTION_STRING_XML,
            OPTION_HTB,
            OPTION_ENUM,

            OPTION_NUM,

            OPTION_FLAG_HEADER = (0x01 << OPTION_HEADER),
//            OPTION_FLAG_HEADER_XML = (0x01 << OPTION_HEADER_XML), // headerをc/c++用と分けたい場合はリマークを外す.
            OPTION_FLAG_XML = (0x01 << OPTION_XML),
            OPTION_FLAG_JSON = (0x01 << OPTION_JSON),
            OPTION_FLAG_CSV = (0x01 << OPTION_CSV),
            OPTION_FLAG_MSGP = (0x01 << OPTION_MSGP),
            OPTION_FLAG_VALB = (0x01 << OPTION_VALB),
            OPTION_FLAG_STRB = (0x01 << OPTION_STRB),
            OPTION_FLAG_STRING_XML = (0x01 << OPTION_STRING_XML),
            OPTION_FLAG_HTB = (0x01 << OPTION_HTB),
            OPTION_FLAG_ENUM = (0x01 << OPTION_ENUM),

            OPTION_FLAG_ALL = (
                OPTION_FLAG_HEADER | 
//                OPTION_FLAG_HEADER_XML | 
                OPTION_FLAG_XML | 
                OPTION_FLAG_JSON | 
                OPTION_CSV | 
                OPTION_FLAG_MSGP |
                OPTION_FLAG_VALB | 
                OPTION_FLAG_STRB | 
                OPTION_FLAG_STRING_XML | 
                OPTION_FLAG_HTB |
                OPTION_FLAG_ENUM ),
        };

        enum Argument : int
        {
            ARG_XLSX = 0,
            ARG_INI,
            ARG_SHEET,
            ARG_OUTPUT,

            ARG_MIN_NUM,

            ARG_OPTION = ARG_MIN_NUM,

            ARG_NUM,
            ARG_OPTION_INI = 0,
        };

        static InitData m_initData;
        static string VERSION = "1.0.0";

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static int Main( string[] args)
        {
            // コンバータの説明
            Console.WriteLine("xlsConverter : version " + VERSION);

            if ((int)Argument.ARG_MIN_NUM > args.Length)
            {
                if (1 == args.Length && 0 == args[(int)Argument.ARG_OPTION_INI].CompareTo("/INI"))
                {   // templateのiniを出力する.
                    Console.WriteLine("\"templateIni.xml\"を作成します");
                    m_initData = new InitData();
                    createTmpIni();
                    FileStream fs = new FileStream("templateIni.xml", FileMode.Create);
                    if (null != fs)
                    {
                        XmlSerializer serializer = new XmlSerializer(typeof(InitData));
                        serializer.Serialize(fs, m_initData);
                        fs.Close();
                    }
                    // iniを出力して終了
                    return 0;
                }
                else
                {
                    writeHelp();
                    return 0; // 引数が足りなければヘルプ表示して終了
                }
            }

            // xmlファイルから初期化
            {
                m_initData = new InitData();
                string iniPath = getIniPath( ref args);
                iniPath = Path.GetFullPath(iniPath);
                Console.WriteLine(iniPath + "を読み込みます");

                if (System.IO.File.Exists(iniPath))
                {
                    FileStream fs = new FileStream(iniPath, FileMode.Open);
                    if (null != fs)
                    {
                        Console.WriteLine(iniPath + "を読み込みました");
                        XmlSerializer serializer = new XmlSerializer(typeof(InitData));
                        m_initData = (InitData)serializer.Deserialize(fs);
                        fs.Close();
                        fs.Dispose();
                    }
                    else
                    {
                        Console.WriteLine(iniPath + "を読み込めません。ファイルがロックされている可能性があります。");
                        return -1; // 解析不能なので終了
                    }
                }
                else
                {
                    Console.WriteLine(iniPath + "が存在しません");
                    return -1; // 解析不能なので終了
                }

                // iniのチェック
                if (!Analyse.checkPadding(m_initData.publish.alignment))
                {   // 指定が無い場合はデファオトが4指定なのでここはスルーされる.
                    Console.Write(
                        "alignmentが正しくありません.\n" +
                        "alignmentは2の冪で且つ4以上である必要があります\n");
                    return -1; // 指定の不正により終了
                }
            }

            bool result = true;
            try
            {
                string filePath = getTargetFilePath(ref args);
                string sheetName = getSheetName(ref args);

                // excel 読み込み
                Analyse analyse = new Analyse(filePath, sheetName);

                // ヘッダー情報(パラメータ情報)の解析.
                analyse.analyseExcelHeader(m_initData.analyse.key, m_initData.analyse.ext_column, m_initData.publish.alignment);
                // 各列のパラメータを保存する.
                analyse.analyseExcelValues(m_initData.analyse.line, m_initData.analyse.eod);

                // データの出力
                {
                    int option = getOption(ref args);

                    // publishにデータを設定する
                    Publish.setCommentTabNum(m_initData.publish.commentTab);


                    // ヘッダー
                    // c/c++ 
                    if (0 != (option & (int)Option.OPTION_FLAG_HEADER))
                    {
                        switch (m_initData.publish.header.type.ToLower())
                        {
                            case "c":
                                Publish.publishHeaderC(
                                       getOutputPath(ref args,ref m_initData.publish.header.header, ref m_initData.publish.header.footer, ref m_initData.publish.header.ext), 
                                       getOutputFileName(ref args), 
                                       ref analyse );
                                break;
                            case "xml":
                                Publish.publishHeaderXml(
                                    getOutputPath(ref args, ref m_initData.publish.header.header, ref m_initData.publish.header.footer, ref m_initData.publish.header.ext),
                                    ref analyse);
                                break;
                            default:
                                throw new Exception("Header type error!" + m_initData.publish.header.type);
                        }
                    }

                    // データ
                    if (0 != (option & (int)Option.OPTION_FLAG_XML))
                    {
                        // xml
                        Publish.publishXml(
                            getOutputPath(ref args, ref m_initData.publish.xml.header, ref m_initData.publish.xml.footer, ref m_initData.publish.xml.ext),
                            ref analyse);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_JSON))
                    {
                        // json
                        Publish.publishJson(
                            getOutputPath(ref args, ref m_initData.publish.json.header, ref m_initData.publish.json.footer, ref m_initData.publish.json.ext),
                            ref analyse);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_CSV))
                    {
                        // csv
                        Publish.publishCSV(
                            getOutputPath(ref args, ref m_initData.publish.csv.header, ref m_initData.publish.csv.footer, ref m_initData.publish.csv.ext),
                            ref analyse);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_MSGP))
                    {
                        // messagepack
                        Publish.publishMessagepack(
                            getOutputPath(ref args, ref m_initData.publish.messagepack.header, ref m_initData.publish.messagepack.footer, ref m_initData.publish.messagepack.ext),
                            ref analyse);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_VALB))
                    {
                        // valb
                        Publish.publishValBinaly(
                            getOutputPath(ref args, ref m_initData.publish.valb.header, ref m_initData.publish.valb.footer, ref m_initData.publish.valb.ext),
                            ref analyse);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_STRB))
                    {
                        string tmp = "";
                        // strb
                        Publish.publishStrBinaly(
                            getOutputPath(ref args, ref m_initData.publish.strb.header, ref m_initData.publish.strb.footer, ref tmp),
                            m_initData.publish.strb.ext,
                            m_initData.publish.strb.language,
                            getOutputCode(m_initData.publish.strb.code),
                            ref analyse);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_HTB))
                    {
                        // htb
                        HashPublishIni hashIni = new HashPublishIni();
                        hashIni.hashBit = m_initData.publish.htb.bit;
                        hashIni.hash = m_initData.publish.htb.hash;
                        hashIni.comment = m_initData.publish.htb.comment;
                        hashIni.empty = m_initData.publish.htb.empty;
                        Publish.publishHTBinaly(
                            getOutputPath(ref args, ref m_initData.publish.htb.binaly_header, ref m_initData.publish.htb.binaly_footer, ref m_initData.publish.htb.binaly_ext),
                            getOutputPath(ref args, ref m_initData.publish.htb.source_header, ref m_initData.publish.htb.source_footer, ref m_initData.publish.htb.source_ext),
                            ref analyse,
                            ref hashIni);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_STRING_XML))
                    {
                        // string xml
                        Publish.publishStringXml(
                            getOutputPath(ref args, ref m_initData.publish.stringxml.header, ref m_initData.publish.stringxml.footer, ref m_initData.publish.stringxml.ext),
                            m_initData.publish.stringxml.language,
                            m_initData.publish.stringxml.hash, 
                            ref analyse);
                    }

                    if (0 != (option & (int)Option.OPTION_FLAG_ENUM))
                    {
                        // enum
                        Publish.publishEnum(
                            getOutputPath(ref args, ref m_initData.publish.enumdata.header, ref m_initData.publish.enumdata.footer, ref m_initData.publish.enumdata.ext),
                            m_initData.publish.enumdata.id,
                            m_initData.publish.enumdata.define,
                            m_initData.publish.enumdata.comment,
                            m_initData.publish.enumdata.type,
                            ref analyse);
                    }
                }
            }
            catch ( Exception e )
            {
                Console.WriteLine("Exception : {0}", e.Message);
                result = false;
            }
            finally
            {
                Console.WriteLine("convert finish");
            }

            return result ? 0 : -1;
        }

        /// <summary>
        /// ヘルプ内容の出力
        /// </summary>
        static void writeHelp()
        {
            Console.WriteLine(
                "↓↓↓↓ xlsConverter Help ↓↓↓↓");
            Console.Write(
                "書式: xlsConverter [xlsx] [xml] [sheet] [output] [option]\n" +
                "第1引数 : xlsxの指定\n" + 
                "第2引数 : 初期化用xmlの指定\n" +
                "第3引数 : Sheet名の指定\n" +
                "第4引数 : 出力ファイル名\n" + 
                "第5引数～ : コンバートモードの指定\n" + 
                "   コンバートモードオプション:\n" + 
                "       /H  : ヘッダーを出力します\n" +
//                "       /HX : xml形式でヘッダーを出力します\n" +
                "       /X  : xml形式で出力します\n" +
                "       /J  : json形式で出力します\n" +
                "       /C  : csv形式で出力します\n" +
                "       /M  : MessagePack形式で出力します\n" +
                "       /V  : valbを出力します\n" +
                "       /S  : strbを出力します\n" +
                "       /SX : String xmlを出力します\n" +
                "       /HT : htbを出力します\n" +
                "       /E  : enumリストを出力します\n" + 
                "---初期化用xmlについて---\n" +
                "初期化用xmlの記述は、以下になります\n" +
                "InitData : ルート\n" +
                "エレメント\n"+
                "Analyse : 解析機能に関するデータです\n" +
                "eod  : 必須 データの有無を検索する列を指定します\n" +
                "key  : 必須 データのKeyを宣言する行を指定します\n" + 
                "line : 必須 データの検索を開始する行を指定します\n" +
                "ext_column : オプション パラメータは出力しないが解析する列を指定します\n" +
                "Hash/Textに関するデータ\n" +
                "hash : 管理 ハッシュを作成する文字列のある列を指定します。\n" +
                "       /HT /SXを指定した場合は必須です。\n" + 
                "\n" + 
                "Publish : 出力機能に関するデータです\n" +
                "alignment : アラインメントを指定します。\n" +
                "            valb等のバイナリでは、この指定でpaddingを追加します\n" +
                "            alignmentは4以上の２の冪で指定する必要があります。\n" +
                "comment_tab : コメントを挿入する際に、そろえるタブの高さを指定します\n" + 
                "header : 出力する際、ファイル名の頭に追加する文字列を指定します\n" +
                "footer : 出力する際、ファイル名の最後に追加する文字列を指定します\n" +
                "ext : 出力する際、ファイルの拡張子を指定します\n" +
                "type : 出力するデータのタイプを指定します\n" +
                "       /Hの場合\n" +
                "       c   : c/c++ ソース形式\n" +
                "       xml : xml形式\n" +
                "       /Eの場合\n" +
                "       Define : #define で定数を定義します\n" +
                "       Enum   : Enumの定義を列挙します。\n" + 
                "empty: 管理 空の行があることを許容します\n" +
                "code : 管理 出力する文字列のエンコードを指定します\n" +
                "       未指定 : UTF-8形式で出力します\n" +
                "       SJIS : SHIFT-JIS形式で出力します\n" +
                "       UTF8 : UTF-8形式で出力します\n" +
                "language: 言語用 出力する言語のキーを指定します。\n" +
                "          複数の指定がある場合は , で区切って指定します。\n" +
                "          例えば、JP,EN,KOと指定すれば、そのkeyを指定してある列を出力します。\n" +
                "          strb,stringXml等でlanguageは参照されます\n" +
                "          strbのファイル名は、header + PATH + footer + _key + extになります。\n" +
                "          footerでkey名を指定する必要はありません\n" +
                "bit: 管理 Hash値のbit長を指定します。現在は(32,64)のみ指定できます。\n" +
                "     /HTを指定した場合は必須です。\n" +
                "id : 管理 enumを出力する際のIDを取得する列名を指定します。\n" +
                "     /Eを指定した場合は必須です。\n" +
                "define : 管理 enumを出力する際の定義を取得する列名を指定します。\n" +
                "         /Eを指定した場合は必須です。\n" +
                "comment : 管理 enumを出力する際のコメントを取得する列名を指定します。\n" +
                "その他の挙動について\n" +
                "引数が足りなかった場合はヘルプが表示されます。\n" +
                "第1引数に/INIオプションを指定した場合のみtemplateIni.xmlのを出力します。\n" +
                "必要に応じてテンプレートのtemplateIni.xmlを編集して使用してください。\n" +
                "\n"
                );
            Console.WriteLine(
                "↑↑↑↑ xlsConverter Help ↑↑↑↑");
        }

        /// <summary>
        /// 読み込むxlsxのファイルパスを取得する
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static string getTargetFilePath( ref string[] args )
        {
            const int index = (int)Argument.ARG_XLSX;
            // ファイルパスを取得するインデックス
            return args[index];
        }

        static string getIniPath( ref string[] args)
        {
            const int index = (int)Argument.ARG_INI;
            return args[index];
        }

        static string getSheetName( ref string[] args)
        {
            const int index = (int)Argument.ARG_SHEET;

            return args[index];
        }

        static string getOutputPath( ref string[] args, ref string header, ref string footer, ref string ext)
        {
            const int index = (int)Argument.ARG_OUTPUT;
            string directory = Path.GetDirectoryName(Path.GetFullPath(args[index]));
            string filename = Path.GetFileNameWithoutExtension(args[index]);

            return directory + @"\" + header + filename + footer + ext;
        }

        static string getOutputFileName(ref string[] args)
        {
            const int index = (int)Argument.ARG_OUTPUT;
            return Path.GetFileNameWithoutExtension(args[index]);
        }

        static STRBCodePage getOutputCode(string codeString)
        {
            STRBCodePage code = STRBCodePage.STRB_CODE_PAGE_UTF8;
            switch (codeString.ToUpper())
            {
                // STRB_CODE_PAGE_SJIS(0)
                case "SJIS":
                case "SHIFT-JIS":
                    code = STRBCodePage.STRB_CODE_PAGE_SJIS;
                    break;
                // STRB_CODE_PAGE_UTF8(1)
                case "UTF-8":
                case "UTF8":
                    code = STRBCodePage.STRB_CODE_PAGE_UTF8;
                    break;
                default:
                    break;
            }
            return code;
        }

        static int getOption( ref string[] args)
        {   // オプションを取得する
            string[] options = new string[(int)Option.OPTION_NUM]
            {
                "/H",
 //               "/HX",
                "/X",
                "/J",
                "/C",
                "/M",
                "/V",
                "/S",
                "/SX",
                "/HT",
                "/E"
            };
            const int startIndex = (int)Argument.ARG_OPTION;
            int option = 0;
            for (int i = startIndex; i < args.Length; ++i)
            {   // オプション分だけ検索
                bool enableOption = false;
                for (int j = 0; j < options.Length; ++j)
                {
                    // 大文字、小文字を区別しないで比較する.
                    if (0 == String.Compare(options[j], args[i], true ) )
                    {
                        option |= (0x01 << j);
                        enableOption = true;
                        break;
                    }
                }
                if (!enableOption)
                {
                    writeHelp();
                    Console.WriteLine("存在しないオプションが指定されています。 : " + args[i]);
                }
            }

            if (0 == (option & (int)Option.OPTION_FLAG_ALL))
            {
                writeHelp();
                throw new Exception("オプションが指定されていません");
            }

            return option;
        }

        static void createTmpIni()
        {
            m_initData.analyse.eod = "A";
            m_initData.analyse.line = 2;
            m_initData.analyse.key = 3;
            m_initData.analyse.ext_column = "B";
            m_initData.publish.alignment = 8;
            m_initData.publish.htb.bit = 32;
            m_initData.publish.htb.hash = "B";
            m_initData.publish.stringxml.hash = "B";
            m_initData.publish.strb.language = "jp,en,ko,zhs,zht";
            m_initData.publish.stringxml.language = "jp,en,ko,zhs,zht";
            m_initData.publish.strb.code = "UTF-8";
            m_initData.publish.header.ext = ".h";
            m_initData.publish.htb.empty = true;
            m_initData.publish.enumdata.id = "A";
            m_initData.publish.enumdata.define = "B";
            m_initData.publish.enumdata.comment = "C";
            m_initData.publish.enumdata.type = "Define";
        }

    }
}
