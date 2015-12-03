using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace xlsConverter
{
    /// <summary>
    /// バイナリのキー情報
    /// </summary>
    struct BinaryKey
    {
        public int hash;        // keyのハッシュ(32bitもいらないかも？16bitで十分？)
        public int type;        // キーのタイプ(数値、文字列、等)( 32bitもいらないかも？16bitで十分？)
        public int offset;      // データまでのオフセット.
        public int column;      // 列番号(出力はされない)
        public string columnName;   // 列名
        public int size;        // キーのサイズ.
        public string key;      // キーの文字列(出力はされない)

        public byte[] getByte()
        {
            int size = 0;
            // 必要なのはこの３つ
            byte[] tmp1 = BitConverter.GetBytes(hash);
            byte[] tmp2 = BitConverter.GetBytes(offset);
            byte[] tmp3 = BitConverter.GetBytes(type);
            size += tmp1.Length + tmp2.Length + tmp3.Length;

            byte[] bytes = new byte[size];
            tmp1.CopyTo(bytes, 0);
            tmp2.CopyTo(bytes, tmp1.Length);
            tmp3.CopyTo(bytes, tmp1.Length + tmp2.Length);

            return bytes;
        }
    };

    /// <summary>
    /// 読み込んだシートの情報
    /// </summary>
    struct ExcelSheetData
    {
        public int maxRow;
        public int maxColumn;
        public Excel.Worksheet sheet;
    };

    /// <summary>
    /// キータイプ情報
    /// </summary>
    enum KeyType : int
    {
        KEY_INVALID = 0,        // 不正なキータイプ.
        KEY_VALID = 1,          // 有効位置.
        KEY_S8 = 1,             // 整数1byte.
        KEY_U8 = 2,             // 整数1byte.
        KEY_S16 = 3,            // 整数2byte.
        KEY_U16 = 4,            // 整数2byte.
        KEY_S32 = 5,            // 整数4byte.
        KEY_U32 = 6,            // 整数4byte.
        KEY_S64 = 7,            // 整数8byte.
        KEY_U64 = 8,            // 整数8byte.
        KEY_FLOAT = 9,          // 実数4byte.
        KEY_DOUBLE = 10,        // 実数8byte.
        KEY_FIXED_STRING = 11,  // 文字列[指定文字数]
        KEY_STRING = 12,        // 文字列[文字数指定なし].
        KEY_PADDING = 13,       // パディング.
        KEY_EXTENSION = 14,     // 拡張キー.
    };
    struct KeyTypeData
    {
        public string key;
        public string dim;
        public int size;

        public KeyTypeData(string _key, string _dim, int _size)
        {
            key = _key;
            size = _size;
            dim = _dim;
        }
    };

    class Analyse
    {
        #region メンバ
        Excel.Application m_objApp = null;
        Excel.Workbook m_objBook = null;
        ExcelSheetData m_sheetData;

        KeyTypeData[] m_keyTypeList;
        List<BinaryKey> m_keyList;
        Dictionary<int, List<String>> m_paramList;
        #endregion

        #region プロパティ
        public KeyTypeData[] KeyTypeList
        {
            get
            {
                return m_keyTypeList;
            }
        }

        public List<BinaryKey> HeaderKeyList
        {
            get
            {
                return m_keyList;
            }
        }

        public Dictionary<int, List<String>> ParamList
        {
            get
            {
                return m_paramList;
            }
        }
        #endregion

        public Analyse(string filepath, string sheet)
        {
            // 初期化.
            initKeyTypeList();
            // 読み込み.
            loadExcel(filepath, sheet);
        }

        ~Analyse()
        {
            if (m_objBook != null)
            {
                m_objBook.Close(false, Type.Missing, Type.Missing);
            }
            if (m_objApp != null)
            {
                m_objApp.Quit();
            }
        }

        public void analyseExcelHeader(int keyRow, string ext_column, int alignment)
        {
            // 作成.
            m_keyList = new List<BinaryKey>();

            // 高速化.
            Excel.Range binRange = m_sheetData.sheet.get_Range("A" + keyRow, getColumnName(m_sheetData.maxColumn) + keyRow);
            Object[,] rangeValue = (Object[,])binRange.Value;                       // EXCEL10以上
            //                binRange = binRange.get_Resize(1, maxColumn); // rangeを取ってくる際、A+keyだけ取ってきてからresizeしても良い.
            //              Object[,] rangeValue = (Object[,])binRange.get_Value(Missing.Value);    // EXCEL9以下
            //valueの2次元配列として参照する.
            int offset = 0;
            int paddingIndex = 1;   // default:1
            string[] extension = ext_column.Split(',');
            bool isExtension = (0 < extension.Length);
            for (int i = 1; i <= m_sheetData.maxColumn; ++i)
            {
                // 高速化前(シートからCellsでrangeを指定して取得してくる→rangeを検索しに行くので低速)
                //                    Excel.Range current = objSheets.Cells[m_initData.key, i];
                //                    if (null == current.Value2)
                //                    {
                //                        continue;
                //                    }
                //                    string tmp = current.Value2;
                // 高速化後(Object[,]の２次元配列にアクセスしているだけなので、Rangeでアクセスするより高速)

                if (null == rangeValue[1, i])
                {
                    if (isExtension &&
                        -1 != Array.IndexOf(extension, getColumnName(i)))
                    { // 拡張キーの保存.
                        BinaryKey extKey = new BinaryKey();
                        extKey.key = "extension";
                        extKey.type = (int)KeyType.KEY_EXTENSION;
                        extKey.size = 0;
                        extKey.offset = 0;
                        extKey.hash = 0;
                        extKey.column = i;                // カラムを保存
                        extKey.columnName = getColumnName(extKey.column);
                        m_keyList.Add(extKey);
                    }
                    continue;
                }
                // 値を文字列に変換.
                String tmp = rangeValue[1, i].ToString();
                if (String.IsNullOrWhiteSpace(tmp) || '#' == tmp[0])
                {   // キーが空、null、空白文字の場合は無効.
                    // 0文字目が#の場合も無効
                    if (isExtension &&
                        -1 != Array.IndexOf(extension, getColumnName(i)) )
                    {
                        // 拡張キーの保存.
                        BinaryKey extKey = new BinaryKey();
                        extKey.key = "extension";
                        extKey.type = (int)KeyType.KEY_EXTENSION;
                        extKey.size = 0;
                        extKey.offset = 0;
                        extKey.hash = 0;
                        extKey.column = i;                // カラムを保存
                        extKey.columnName = getColumnName(extKey.column);
                        m_keyList.Add(extKey);
                    }
                    continue;
                }

                if (!System.Text.RegularExpressions.Regex.IsMatch(
                       tmp,
                       @"^[_:a-zA-Z0-9\*\[\]]+$"))
                {   // 変数に有効な文字列ではない( 日本語等が入っている可能性がある )
                    Console.WriteLine("Caution: 解析できないキーが見つかりました" + getColumnName(i) + " " + tmp + " ※この列のデータは出力されません");
                    Console.WriteLine("keyで指定している行に、key以外の文字列を入力することは禁止しています");
                    continue;
                }

                BinaryKey key = new BinaryKey();

                String[] type = tmp.Split(':');
                String keyValue = null;
                KeyType keyType = KeyType.KEY_INVALID;
                int size = 0;


                for (int j = 0; j < type.Length; ++j)
                {
                    if (!getKeyType( ref type[j], out keyType, out size ))
                    {   // keyの名前である.
                        keyValue = type[j];
                    }
                }

                if (KeyType.KEY_INVALID != keyType)
                {   // キーが無効でなければ大丈夫.
                    if (KeyType.KEY_FIXED_STRING != keyType ||
                        (0 == (m_keyTypeList[(int)keyType].size & 1)))
                    {
                        // charではない(間のpaddingのチェック)
                        int padding = offset & (m_keyTypeList[(int)keyType].size - 1);
                        if (0 != padding)
                        {
                            BinaryKey padKey = new BinaryKey();
                            padKey.key = "padding" + paddingIndex;
                            padKey.type = (int)KeyType.KEY_PADDING;
                            padKey.size = m_keyTypeList[(int)keyType].size - padding;   // 自身のサイズに依存する.
                            padKey.offset = offset;
                            padKey.hash = FnvHash.getFNV_1_32(padKey.key);
                            padKey.column = -1;    // 無効.
                            padKey.columnName = "padding";
                            m_keyList.Add(padKey);

                            offset += padKey.size;
                            ++paddingIndex;
                        }
                    }
                    // 結果を保存.
                    key.key = keyValue;
                    key.type = (int)(keyType);
                    key.size = size;
                    key.column = i;                // カラムを保存
                    key.columnName = getColumnName(key.column);
                    key.hash = FnvHash.getFNV_1_32(keyValue);   // ハッシュを保存
                    key.offset = offset;
                    m_keyList.Add(key);
                    // サイズを保存する.
                    if (0 < size)
                    {
                        offset += size;
                    }
                }
            }
            {
                // paddingのチェック( 構造体サイズは4の倍数であること )
                int padding = (offset & ( alignment -1 ) );
                if (0 != padding)
                {
                    BinaryKey padKey = new BinaryKey();
                    padKey.key = "padding" + paddingIndex;
                    padKey.type = (int)KeyType.KEY_PADDING;
                    padKey.size = alignment - padding;
                    padKey.offset = offset;
                    padKey.hash = FnvHash.getFNV_1_32(padKey.key);
                    padKey.column = -1;    // 無効.
                    padKey.columnName = "padding";
                    m_keyList.Add(padKey);

                    offset += padKey.size;
                    ++paddingIndex;
                }
            }
        }


        public void analyseExcelValues(int row, string eod)
        {
            // EODを検索する.
            m_sheetData.maxRow = analyseExcelValuesEOD(row, eod);


            // キーはカラム
            m_paramList = new Dictionary<int, List<String>>();

            for (int i = 0; i < m_keyList.Count; ++i)
            {
                if ((int)KeyType.KEY_PADDING == m_keyList[i].type)
                {
                    continue;
                }
                // 各列の情報を取得する。
                analyseExcelValuesColumn(m_keyList[i].column, row);
            }
        }

        private int analyseExcelValuesEOD(int row, string eod)
        {
            int line = m_sheetData.maxRow;

            Excel.Range binRange = m_sheetData.sheet.get_Range(eod + 1, eod + m_sheetData.maxRow);
            if (binRange == null)
            {   // EODのラインが不定.
                return m_sheetData.maxRow;
            }
            Object[,] rangeValue = (Object[,])binRange.Value;                       // EXCEL10以上

            for (int i = row; i < m_sheetData.maxRow; ++i)
            {
                if (null == rangeValue[i, 1])
                {   // null の時は空で.
                    continue;
                }
                if (0==rangeValue[i, 1].ToString().CompareTo("EOD"))
                {
                    line = i;
                    break;
                }
            }

            return line;
        }

        private void analyseExcelValuesColumn(int column, int row)
        {
            m_paramList[column] = new List<String>();
            Excel.Range binRange = m_sheetData.sheet.get_Range(getColumnName(column) + 1, getColumnName(column) + m_sheetData.maxRow);
            Object[,] rangeValue = (Object[,])binRange.Value;                       // EXCEL10以上

            // row の位置から開始.
            for (int i = row; i < m_sheetData.maxRow; ++i)
            {
                if (null == rangeValue[i, 1])
                {   // null の時は空で.
                    m_paramList[column].Add(String.Empty);
                }
                else
                {
                    // 文字列を取得する.
                    m_paramList[column].Add(rangeValue[i, 1].ToString());
                }
            }
        }

        /// <summary>
        /// キーの判定.
        /// </summary>
        private bool getKeyType(ref string value, out KeyType type, out int size)
        {
            bool result = false;
            type = KeyType.KEY_INVALID;
            size = 0;
            for (int k = (int)KeyType.KEY_VALID; k < m_keyTypeList.Count(); ++k)
            {
                if (k == (int)KeyType.KEY_FIXED_STRING)
                {   // 固定文字の判定.
                    if (-1 != value.ToLower().IndexOf(m_keyTypeList[k].key))
                    {

                        int startIndex = value.IndexOf("[") + 1;
                        int endIndex = value.IndexOf("]");
                        int stringSize = endIndex - startIndex;
                        if (-1 != startIndex && -1 != endIndex && 1 < stringSize)
                        {
                            string sizeStr = value.Substring(startIndex, stringSize);
                            int tmpSize = int.Parse(sizeStr);
                            if (0 < tmpSize)
                            {   // 構造体内に文字列を含める場合.
                                // データの値がキチンと取れればタイプを保存する.
                                type = KeyType.KEY_FIXED_STRING;
                                size = tmpSize;
                                result = true;
                                break;
                            }
                            else
                            {
                                Console.WriteLine(value + ": サイズ固定の文字列タイプを指定しましたが、文字数が設定されていません");
                                throw new Exception();
                            }
                        }
                    }
                }
                else
                    if (0 == String.Compare(value, m_keyTypeList[k].key, true))
                    {   // 型が見つかった.
                        type = (KeyType)k;
                        size = m_keyTypeList[k].size;
                        result = true;
                        break;
                    }
            }
            return result;
        }


#if false // そもそもいらない？
        /// <summary>
        /// 旧xlsConverterのiniファイルの出力.
        /// </summary>
        static bool outputHeaderIni(ref List<BinaryKey> keyList)
        {
            bool result = false;
            /// 
            // UTF-8でエレメントの情報をini出力
            System.IO.StreamWriter writer = new System.IO.StreamWriter(
                @"headerTest.ini",  // @todo ファイル名はどっかで受け取る.
                false,              // 追加する記述ではないため、新しく出力する.
                Encoding.UTF8);
            try
            {
                writer.WriteLine("//xlsConverter設定ファイル");
                writer.WriteLine();
                writer.WriteLine("// -----------------------------------------------------------------------------");
                writer.WriteLine("// コンバート対象行");
                writer.WriteLine("LINE:" + m_initData.line + ",EOD" + m_initData.eod);
                writer.WriteLine("// -----------------------------------------------------------------------------");
                writer.WriteLine("//TYPE:構造体名(指定しない場合はNoneになる");
                writer.WriteLine("TYPE:" + m_initData.bin);
                writer.WriteLine("//ENUM:enum列挙列(指定しない場合はB列になる");
                writer.WriteLine("ENUM:" + m_initData.hash);    // hash列がenumになるはず.
                writer.WriteLine();
                writer.WriteLine("// -----------------------------------------------------------------------------");
                // キーの出力.
                for (int i = 0; i < keyList.Count; ++i)
                {
                    string columnName = String.Empty;
                    string type = String.Empty;
                    if (keyList[i].type == (int)(KeyType.KEY_FIXED_STRING))
                    {
                        type = "char" + "[" + keyList[i].size + "]";
                        columnName = getColumnName(keyList[i].column);
                    }
                    else
                        if (keyList[i].type == (int)(KeyType.KEY_PADDING))
                        {
                            type = Convert.ToString(keyList[i].size);
                            columnName = "pad"; // paddingである定義.
                        }
                        else
                        {
                            type = m_keyTypeList[keyList[i].type].key;
                            columnName = getColumnName(keyList[i].column);
                        }
                    if (0 > keyList[i].column)
                    {
                        columnName = "pad"; // paddingである定義.
                    }
                    else
                    {
                        columnName = getColumnName(keyList[i].column);
                    }
                    writer.WriteLine(
                        "TAG:" + columnName + @"," +
                        type + "," +
                        keyList[i].key + ";" +
                        "\t\t\t\t//" + (keyList[i].offset + keyList[i].size) + "." + keyList[i].key.ToUpper()
                    );
                }
                // フラッシュして出力内容を書き出し.
                writer.Flush();

                result = true;
            }
            catch
            {
                result = false;
                Console.WriteLine("ini output error.");
                throw;
            }
            finally
            {
                // ファイル閉じ.
                writer.Close();
            }

            return result;
        }
#endif

        public KeyTypeData getKeyTypeData( KeyType type ) 
        {
            return m_keyTypeList[(int)type];
        }

        public string getParamFromKeyIndex(int key, int index)
        {
            return m_paramList[m_keyList[key].column][index];
        }

        private void loadExcel(string filePath, string sheetName)
        {
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine("ファイルが見つかりません filepath : " + filePath);
                throw new Exception();
            }
            // アプリケーション作成
            m_objApp = new Excel.Application();
            m_objApp.Visible = false;
            //            string currentPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            filePath = Path.GetFullPath(filePath);
            // ブック読み込み
            m_objBook = m_objApp.Workbooks.Open(
                    filePath,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // シート検索
            int sheetIndex = Analyse.getWorkSheetIndex(m_objBook, sheetName);
            if (0 > sheetIndex)
            {
                Console.WriteLine("Error!! sheet get error! sheetname = " + sheetName);
                throw new Exception();
            }
            // シートの読み込みに成功したら、情報を保存する.
            m_sheetData = new ExcelSheetData();
            m_sheetData.sheet = (Excel.Worksheet)m_objBook.Sheets[sheetIndex];
            Excel.Range maxCell = m_sheetData.sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            m_sheetData.maxRow = maxCell.Row;
            m_sheetData.maxColumn = maxCell.Column;
        }

        private void initKeyTypeList()
        {   // 値の初期化
            m_keyTypeList = new KeyTypeData[]
            {
                new KeyTypeData( "invalid", "invalid", -1 ),           // KEY_INVALID = 0,       // 無効.
                new KeyTypeData( "s8", "sint8", sizeof(SByte) ),       // KEY_S8 = 1,            // 整数1byte.
                new KeyTypeData( "u8", "uint8", sizeof(Byte) ),        // KEY_U8 = 2,            // 整数1byte.
                new KeyTypeData("s16", "sint16", sizeof(Int16) ),      // KEY_S16 = 3,           // 整数2byte.
                new KeyTypeData("u16", "uint16", sizeof(UInt16) ),     // KEY_U16 = 4,           // 整数2byte.
                new KeyTypeData("s32", "sint32", sizeof(Int32) ),      // KEY_S32 = 5,           // 整数4byte.
                new KeyTypeData("u32", "uint32", sizeof(UInt32) ),     // KEY_U32 = 6,           // 整数4byte.
                new KeyTypeData("s64", "sint64", sizeof(Int64) ),      // KEY_S64 = 7,           // 整数8byte.
                new KeyTypeData("u64", "uint64", sizeof(UInt64) ),     // KEY_U64 = 8,           // 整数8byte.
                new KeyTypeData("float", "float", sizeof(float) ),     // KEY_FLOAT = 9,         // 実数4byte.
                new KeyTypeData("double", "double", sizeof(double) ),  // KEY_DOUBLE = 10,       // 実数8byte.
                new KeyTypeData("char", "char", -1 ),                  // KEY_FIXED_STRING = 11, // 文字列[指定文字数]
                new KeyTypeData("string", "uint32", sizeof(UInt32) ),  // KEY_STRING = 12,       // 文字列[文字数指定なし].
                new KeyTypeData("padding", "sint8", -1),               // KEY_PADDING = 13,      // パディング.
                new KeyTypeData("extension", "", -1 ),                 // KEY_EXTENSION = 14,    // 拡張
            };
        }

        public static int getWorkSheetIndex(Excel.Workbook book, string SheetName)
        {
            int index = 1;

            foreach (Excel.Worksheet sh in book.Sheets)
            {
                if (SheetName == sh.Name)
                {
                    return index;
                }
                ++index;
            }

            return -1;
        }

        public static string getColumnName(int value)
        {
            const int ALPHABET_A = 'A'; // Aのコード
            const int NUM_ALPHABET = 26; // アルファベット数.
            string name = String.Empty;

            while (0 < value)
            {
                int per = value % NUM_ALPHABET;
                value /= NUM_ALPHABET;
                if (0 == per)
                {
                    per = NUM_ALPHABET;
                    --value;
                }
                name += Convert.ToChar(ALPHABET_A + (per - 1));

            }
            if (1 < name.Length)
            {   // 2文字以上なら逆順に
                char[] chars = name.ToCharArray();
                Array.Reverse(chars);
                name = new String(chars);
            }
            return name;
        }

        public static bool checkPadding(int value)
        {
            if (4 > value)
            {   // 最低を満たしていない.
                return false;
            }

            value = (unchecked(value & (int)0xaaaaaaaa) >> 0x01) + (value & (int)0x55555555);
            value = (unchecked(value & (int)0xcccccccc) >> 0x02) + (value & (int)0x33333333);
            value = (unchecked(value & (int)0xf0f0f0f0) >> 0x04) + (value & (int)0x0f0f0f0f);
            value = (unchecked(value & (int)0xff00ff00) >> 0x08) + (value & (int)0x00ff00ff);
            value = (unchecked(value & (int)0xffff0000) >> 0x10) + (value & (int)0x0000ffff);

            // 2の冪か？
            return (1 == value);
        }
    }

}
