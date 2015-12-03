using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace xlsConverter
{
    public class HTBOutputData
    {
        #region メンバ
        private int _index;
        private string _name;
        private string _comment;
        private Object _hash;
        #endregion

        #region プロパティ
        [XmlIgnoreAttribute]    // xmlには出力しない
        public int index
        {
            get { return _index; }
            set { _index = value; }
        }

        [XmlAttribute("name")]
        public string name
        {
            get { return _name; }
            set { _name = value; }
        }

        [XmlIgnoreAttribute]    // xmlには出力しない.
        public string comment
        {
            get { return _comment; }
            set { _comment = value; }
        }

        [XmlAttribute("hash")]
        public Object hash
        {
            get { return _hash; }
            set { _hash = value; }
        }
        #endregion
    };

    // Htbファイルのフォーマット
    struct HTBBinalyHeader
    {
        //	        public char[]	sign = {'H', 'T', 'B', '\0' };	//04.シグネチャ
        //          public ushort version = 0x0001;	                //06.ファイルバージョン
        //          public ushort endian = 0x7755;		            //08.ファイルエンディアン識別子
        public ushort bitType;                          //10.ビットタイプ( 32, 64 )
        public ushort hashnum;	                        //12.文字列のテーブル総数
        //          public uint address;	                        //16.アンパックアドレス

        /// <summary>
        /// バイト情報を取得する
        /// </summary>
        /// <returns></returns>
        public byte[] getByte()
        {
            int size = 0;
            const string sign = "HTB";          // ファイルタイプシグネチャ
            const ushort version = 0x0001;      // バージョンコード
            const ushort endian = 0x7755;       // エンディアン
            uint address = 0;

            byte[] tmp1 = System.Text.Encoding.UTF8.GetBytes(sign);
            byte[] tmp2 = BitConverter.GetBytes(version);
            byte[] tmp3 = BitConverter.GetBytes(endian);
            byte[] tmp4 = BitConverter.GetBytes(bitType);
            byte[] tmp5 = BitConverter.GetBytes(hashnum);
            byte[] tmp6 = BitConverter.GetBytes(address);
            size = (tmp1.Length + 1) + tmp2.Length + tmp3.Length + tmp4.Length + tmp5.Length + tmp6.Length;
            address = (uint)size;   // サイズを保存
            tmp6 = BitConverter.GetBytes(address);

            byte[] bytes = new byte[size];
            size = 0;
            // sign
            tmp1.CopyTo(bytes, 0);
            bytes[tmp1.Length] = 0;
            size += tmp1.Length + 1;
            // version
            tmp2.CopyTo(bytes, size);
            size += tmp2.Length;
            // endian
            tmp3.CopyTo(bytes, size);
            size += tmp3.Length;
            // bitType
            tmp4.CopyTo(bytes, size);
            size += tmp4.Length;
            // hashnum
            tmp5.CopyTo(bytes, size);
            size += tmp5.Length;
            // address
            tmp6.CopyTo(bytes, size);
            size += tmp6.Length;

            return bytes;
        }
    };

    /// <summary>
    /// 32bitハッシュ用ハッシュリスト
    /// </summary>
    struct HashIndexPair32
    {
        public int hash;
        public int index;

        public byte[] getByte()
        {
            int size = 0;
            byte[] tmp1 = BitConverter.GetBytes(hash);
            byte[] tmp2 = BitConverter.GetBytes(index);
            size += tmp1.Length + tmp2.Length;

            byte[] bytes = new byte[size];
            tmp1.CopyTo(bytes, 0);
            tmp2.CopyTo(bytes, tmp1.Length);

            return bytes;
        }
    };
    /// <summary>
    /// 64bitハッシュ用ハッシュリスト
    /// </summary>
    struct HashIndexPair64
    {
        public Int64 hash;
        public int index;

        public byte[] getByte()
        {
            int size = 0;
            byte[] tmp1 = BitConverter.GetBytes(hash);
            byte[] tmp2 = BitConverter.GetBytes(index);
            size += tmp1.Length + tmp2.Length;

            byte[] bytes = new byte[size];
            tmp1.CopyTo(bytes, 0);
            tmp2.CopyTo(bytes, tmp1.Length);

            return bytes;
        }
    };

    struct HashPublishIni
    {
        public int hashBit;
        public string hash;
        public string comment;
        public bool empty;
    };

    enum HashBit : int
    {
        BIT32 = 32,
        BIT64 = 64,
    };
    // valbファイルのフォーマット
    struct VALBBinalyHeader
    {
        public uint dataNum;	        // データテーブル総数
        public uint dataSize;           // データサイズ.
        /// <summary>
        /// バイト情報を取得する
        /// </summary>
        /// <returns></returns>
        public byte[] getByte()
        {
            int size = 0;
            const string sign = "VALB";          // ファイルタイプシグネチャ(valb)
            const ushort version = 0x0100;      // バージョンコード
            const ushort endian = 0x7755;       // エンディアン

            byte[] tmp1 = System.Text.Encoding.UTF8.GetBytes(sign);
            byte[] tmp2 = BitConverter.GetBytes(version);
            byte[] tmp3 = BitConverter.GetBytes(endian);
            byte[] tmp4 = BitConverter.GetBytes(dataNum);
            byte[] tmp5 = BitConverter.GetBytes(dataSize);
            size = tmp1.Length + tmp2.Length + tmp3.Length + tmp4.Length + tmp5.Length;

            byte[] bytes = new byte[size];
            size = 0;
            // sign
            tmp1.CopyTo(bytes, 0);
            bytes[tmp1.Length] = 0;
            size += tmp1.Length;
            // version
            tmp2.CopyTo(bytes, size);
            size += tmp2.Length;
            // endian
            tmp3.CopyTo(bytes, size);
            size += tmp3.Length;
            // datanum
            tmp4.CopyTo(bytes, size);
            size += tmp4.Length;
            // datasize
            tmp5.CopyTo(bytes, size);
            size += tmp5.Length;

            return bytes;
        }
    };

    struct STRBBinalyHeader
    {
        public UInt16 dataNum;	        // データテーブル総数
        public UInt16 codePage;         // 文字列のコンバートコード
        public UInt32 address;          // 文字列の先頭アドレスまでのオフセット.
        /// <summary>
        /// バイト情報を取得する
        /// </summary>
        /// <returns></returns>
        public byte[] getByte()
        {
            int size = 0;
            const string sign = "STRB";          // ファイルタイプシグネチャ(valb)
            const ushort version = 0x0100;      // バージョンコード
            const ushort endian = 0x7755;       // エンディアン

            byte[] tmp1 = System.Text.Encoding.UTF8.GetBytes(sign);
            byte[] tmp2 = BitConverter.GetBytes(version);
            byte[] tmp3 = BitConverter.GetBytes(endian);
            byte[] tmp4 = BitConverter.GetBytes(dataNum);
            byte[] tmp5 = BitConverter.GetBytes(codePage);
            byte[] tmp6 = BitConverter.GetBytes(address);
            size = tmp1.Length + tmp2.Length + tmp3.Length + tmp4.Length + tmp5.Length + tmp6.Length;
            address += (uint)size;
            tmp6 = BitConverter.GetBytes(address);

            byte[] bytes = new byte[size];
            size = 0;
            // sign
            tmp1.CopyTo(bytes, 0);
            bytes[tmp1.Length] = 0;
            size += tmp1.Length;
            // version
            tmp2.CopyTo(bytes, size);
            size += tmp2.Length;
            // endian
            tmp3.CopyTo(bytes, size);
            size += tmp3.Length;
            // datanum
            tmp4.CopyTo(bytes, size);
            size += tmp4.Length;
            // codePage
            tmp5.CopyTo(bytes, size);
            size += tmp5.Length;
            // address
            tmp6.CopyTo(bytes, size);
            size += tmp6.Length;

            return bytes;
        }
    };

    enum STRBCodePage : int
    {
        STRB_CODE_PAGE_SJIS = 0,
        STRB_CODE_PAGE_UTF8 = 1,

        STRB_CODE_PAGE_NUM,
    };

    enum EnumType : int
    {
        ENUM_TYPE_ENUM = 0,
        ENUM_TYPE_DEFINE = 1,
    };

    class Publish
    {
        static int m_commentTabNum = 8;

        public static void setCommentTabNum( int num )
        {
            m_commentTabNum = num;
        }
        private static void addComment( ref string baseString, string comment )
        {
            if ( 0 == comment.Length)
            {   // 入れるコメントが無い.
                return;
            }

            int tabNum = m_commentTabNum - (baseString.Length / 4);
            tabNum = (tabNum <= 0 ? 1 : tabNum);
            for (int tab = 0; tab < tabNum; ++tab)
            {
                baseString += "\t";
            }
            baseString += "// " + comment;
        }

        /// <summary>
        // c/c++用headerの出力
        /// </summary>
        public static bool publishHeaderC(String filePath, String structName, ref Analyse analyse)
        {
            bool result = false;

            // UTF-8でエレメントの情報をini出力
            // ヘッダーを出力
            System.IO.StreamWriter writer = new System.IO.StreamWriter(
                filePath,
                false,
                Encoding.UTF8);
            try
            {
                int size = 0;
                // 出力内容
                writer.WriteLine("// -----------------------------------------------------------------------------");
                writer.WriteLine("// コンバータによる自動出力(編集禁止!!!)");
                writer.WriteLine("struct " + structName + " {");

                List<BinaryKey> keyList = analyse.HeaderKeyList;
                KeyTypeData[] keyTypeList = analyse.KeyTypeList;

                // キーの出力.
                for (int i = 0; i < keyList.Count; ++i)
                {
                    if ((int)KeyType.KEY_EXTENSION == keyList[i].type)
                    {   // 拡張キーは無視.
                        continue;
                    }
                    string type = keyTypeList[keyList[i].type].dim;
                    string key = keyList[i].key;
                    switch (keyList[i].type)
                    {
                        case (int)(KeyType.KEY_FIXED_STRING):
                        case (int)(KeyType.KEY_PADDING):
                            key += "[" + keyList[i].size + "]";
                            break;
                        default:
                            break;
                    }
                    string tmpLine = type + " " + key + ";";
                    string comment = "" + (keyList[i].offset + keyList[i].size) + " " + keyList[i].key.ToUpper();
                    addComment(ref tmpLine, comment);
                    // 出力.
                    writer.WriteLine("\t" + tmpLine);
                    size += keyList[i].size;
                }
                writer.WriteLine("};" + " //" + size + "byte");
                // 書き込み中身をフラッシュ.
                writer.Flush();

                result = true;
            }
            catch
            {
                result = false;
                Console.WriteLine("C/C++ Header output error.");
                throw;
            }
            finally
            {
                writer.Close();
            }

            return result;

        }

        /// <summary>
        // xml形式でheaderを出力
        /// </summary>
        public static bool publishHeaderXml(String filePath, ref Analyse analyse)
        {
            bool result = false;
            // UTF-8でエレメントの情報をxml出力
            XmlTextWriter writer = new XmlTextWriter(
                filePath,  // ファイル名はどっかで受け取る.
                Encoding.UTF8);
            try
            {
                writer.Formatting = Formatting.Indented;
                // document開始
                writer.WriteStartDocument();
                writer.WriteStartElement("params");
                List<BinaryKey> keyList = analyse.HeaderKeyList;
                KeyTypeData[] keyTypeList = analyse.KeyTypeList;

                for (int i = 0; i < keyList.Count; ++i)
                {
                    if ((int)KeyType.KEY_EXTENSION == keyList[i].type ||
                        0 > keyList[i].column)
                    {
                        continue;   // 拡張、或いは無効のパラメータ(padding)
                    }
                    writer.WriteStartElement("keys");
                    writer.WriteAttributeString("name", keyList[i].key);
                    // 列を出力
                    writer.WriteAttributeString("column", keyList[i].columnName);

                    if (keyList[i].type != (int)(KeyType.KEY_FIXED_STRING))
                    {
                        writer.WriteAttributeString("type", keyTypeList[keyList[i].type].key);
                    }
                    else
                    {
                        writer.WriteAttributeString("type", "char" + "[" + keyList[i].size + "]");
                    }
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();   // params
                writer.WriteEndDocument();  // document終了
                writer.Flush();

                result = true;
            }
            catch
            {
                // エラー
                Console.WriteLine("Header XML output Error.");
                throw;
            }
            finally
            {
                writer.Close(); /// 閉じ.
            }

            return result;
        }

        // jsonで出力
        public static bool publishJson(string filePath, ref Analyse analyse)
        {
            bool result = false;
            // jsonを出力
            System.IO.StreamWriter writer = new System.IO.StreamWriter(
                filePath,
                false,
                Encoding.UTF8);

            try
            {
                Dictionary<int, List<String>> paramList = analyse.ParamList;
                List<BinaryKey> keyList = analyse.HeaderKeyList;
                // メンバーが動的に代わるものは使い難い.
                // DataContractJsonSerializer jsonSerializer; // .NET標準
                // 現状excelからの出力はこれで十分だと思うが、必要ならば別のライブラリを使用した方が良いと思う(miniJson等)
                String jsonString = String.Empty;

                int num = paramList.First().Value.Count;

                jsonString += '{';
                // 配列でアクセスさせる？
                jsonString += "\"" + "params" + "\":[";

                for (int i = 0; i < num; ++i)
                {
                    if (0 != i)
                    {
                        jsonString += ",";
                    }
                    jsonString += "{";

                    for (int j = 0; j < keyList.Count; ++j)
                    {
                        if ((int)KeyType.KEY_EXTENSION == keyList[i].type || 
                            0 > keyList[j].column ) 
                        {   // 拡張、或いは無効のキーである.
                            continue;
                        }
                        if ( 0 != j ) {
                            jsonString += ",";
                        }
                        jsonString += "\"" + keyList[j].key + "\":";
                        String param = paramList[keyList[j].column][i];

                        if ((int)KeyType.KEY_FIXED_STRING == keyList[j].type ||
                            (int)KeyType.KEY_STRING == keyList[j].type)
                        {
                            jsonString += "\"" + param + "\"";
                        }
                        else
                        {
                            jsonString += param;
                        }
                    }

                    jsonString += "}";
                }
                // 配列でアクセスさせる？
                jsonString += "]";

                jsonString += '}';

                // 生成したjsonを出力する.
                writer.Write(jsonString);
                // debug
//                Console.Write(jsonString);
                result = true;
            }
            catch
            {
                result = false;
                Console.WriteLine("C/C++ Header output error.");
                throw;
            }
            finally
            {
                writer.Close();
            }

            return result;
        }

        public static bool publishCSV(string filePath, ref Analyse analyse)
        {
            bool result = false;
            // jsonを出力
            System.IO.StreamWriter writer = null;

            try
            {
                // 出力先のファイルを開く.
                writer = new System.IO.StreamWriter(
                                filePath,
                                false,
                                Encoding.UTF8);

                Dictionary<int, List<String>> paramList = analyse.ParamList;
                List<BinaryKey> keyList = analyse.HeaderKeyList;
                // メンバーが動的に代わるものは使い難い.
                // DataContractJsonSerializer jsonSerializer; // .NET標準
                // 現状excelからの出力はこれで十分だと思うが、必要ならば別のライブラリを使用した方が良いと思う(miniJson等)
                String csvString = String.Empty;

                int num = paramList.First().Value.Count;

                for ( int i = 0; i < keyList.Count; ++ i )
                {
                    if ((int)KeyType.KEY_EXTENSION == keyList[i].type ||
                        0 > keyList[i].column)
                    {   // 拡張、或いは無効のキーである
                        continue;
                    }
                    csvString += keyList[i].key;
                    csvString += ",";   // @todo ,かtabか等の指定ができるようにする.
                }
                writer.WriteLine(csvString);
                csvString = String.Empty;
                for (int i = 0; i < num; ++i)
                {
                    for (int j = 0; j < keyList.Count; ++j)
                    {
                        if (0 > keyList[j].column)
                        {
                            continue;
                        }
                        csvString += paramList[keyList[j].column][i];
                        csvString += ",";
                    }
                    writer.WriteLine(csvString);
                    csvString = String.Empty;
                }

                // 生成したjsonを出力する.
//                writer.Write(csvString);
                result = true;
                Console.WriteLine("CSV : published." + filePath);
            }
            catch( Exception e)
            {
                result = false;
                Console.WriteLine("CSV Error : " + e.Message );
                throw;
            }
            finally
            {
                if (null != writer)
                {
                    writer.Close();
                }
            }

            return result;
        }

        // messagepack形式で出力
        public static bool publishMessagepack(string filePath, ref Analyse analyse)
        {
            bool result = false;
            // miniMessagePackを試してみる
            // messagepack for c#
            MiniMessagePack.MiniMessagePacker messagePacker = new MiniMessagePack.MiniMessagePacker();
            List<Dictionary<string, object>> packData = new List<Dictionary<string, object>>();

            FileStream fs = null;
            try
            {
                Dictionary<int, List<String>> paramList = analyse.ParamList;
                List<BinaryKey> keyList = analyse.HeaderKeyList;

                int num = paramList.First().Value.Count;

                for (int i = 0; i < num; ++i)
                {
                    Dictionary<string, object> data = new Dictionary<string, object>();
                    for (int j = 0; j < keyList.Count; ++j)
                    {
                        if ((int)KeyType.KEY_EXTENSION == keyList[i].type ||
                            0 > keyList[j].column)
                        {   // 拡張、或いは無効のキーである
                            continue;
                        }
                        switch ( keyList[j].type )
                        {
                        case (int)KeyType.KEY_S8:
                        case (int)KeyType.KEY_S16:
                        case (int)KeyType.KEY_S32:
                        case (int)KeyType.KEY_U8:
                        case (int)KeyType.KEY_U16:
                        case (int)KeyType.KEY_U32:
                        case (int)KeyType.KEY_S64:
                        case (int)KeyType.KEY_U64:
                            // miniMessagePackの仕様で、全てlong型になる.
                            data.Add(keyList[j].key, System.Convert.ToInt32( paramList[keyList[j].column][i]));
                            break;
                        case (int)KeyType.KEY_DOUBLE:
                        case (int)KeyType.KEY_FLOAT:
                            // miniMessagePackの仕様で、全てdouble型になる.
                            data.Add(keyList[j].key, System.Convert.ToDouble(paramList[keyList[j].column][i]));
                            break;
                        case (int)KeyType.KEY_FIXED_STRING:
                        case (int)KeyType.KEY_STRING:
                            data.Add(keyList[j].key, paramList[keyList[j].column][i]);
                            break;
                        }
                    }

                    packData.Add( data );
                }

                byte[] msgpack = messagePacker.Pack(packData);
                fs = new FileStream(filePath, FileMode.Create);
                fs.Write( msgpack, 0, msgpack.Length);

                // debug
                //Object unpack_test = messagePacker.Unpack(msgpack);
                //

                Console.WriteLine( "MessagePack : published." + filePath );
                result = true;
            } catch( Exception e )
            {
                Console.WriteLine( "MessagePack Error : " + e.Message );
                throw;
            }
            finally
            {
                if (null != fs)
                {
                    fs.Close();
                    fs.Dispose();
                    fs = null;
                }
            }

            return result;
        }

        // xmlで出力
        public static bool publishXml(string filePath, ref Analyse analyse)
        {
            bool result = false;
            // UTF-8でエレメントの情報をxml出力
            XmlTextWriter writer = new XmlTextWriter(
                filePath,  // ファイル名はどっかで受け取る.
                Encoding.UTF8);
            try
            {
                writer.Formatting = Formatting.Indented;
                // document開始
                writer.WriteStartDocument();
                writer.WriteStartElement("params");
                List<BinaryKey> keyList = analyse.HeaderKeyList;
                KeyTypeData[] keyTypeList = analyse.KeyTypeList;
                Dictionary<int, List<String>> paramList = analyse.ParamList;
                int num = paramList.First().Value.Count;


                for (int i = 0; i < num; ++i)
                {
                    writer.WriteStartElement("param");
                    for (int j = 0; j < keyList.Count; ++j)
                    {
                        if ((int)KeyType.KEY_EXTENSION == keyList[i].type ||
                            0 > keyList[j].column)
                        {
                            continue;   // 拡張、或いは無効のパラメータ(padding)
                        }
                        writer.WriteAttributeString(keyList[j].key, paramList[keyList[j].column][i]);
                    }
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();   // params
                writer.WriteEndDocument();  // document終了
                writer.Flush();

                result = true;
            }
            catch( Exception e )
            {
                // エラー
                Console.WriteLine("XML Error : " + e.Message);
                throw;
            }
            finally
            {
                writer.Close(); /// 閉じ.
            }

            return result;
        }

        // STRBで出力
        public static bool publishStrBinaly(string filepath, string ext, string langage, STRBCodePage code, ref Analyse analyse)
        {
            bool result = false;
            FileStream fs = null;

            Dictionary<int, List<String>> paramList = analyse.ParamList;
            List<BinaryKey> keyList = analyse.HeaderKeyList;

            try
            {
                string[] keys = langage.Split(',');
                for (int lang = 0; lang < keys.Length; ++lang)
                {
                    string key = keys[lang];
                    // 出力するカラムを検索する.
                    int column = keyList.Find(x => 0 == x.key.CompareTo(key)).column;
                    string outputPath = filepath + "_" + key + ext;
                    fs = new FileStream(outputPath, FileMode.Create);
                    // 対象のカラムの列を対象にデータを出力する.
                    var stringList = paramList[column];
                    int num = stringList.Count;

                    STRBBinalyHeader header = new STRBBinalyHeader();
                    header.dataNum = (ushort)num;
                    header.codePage = (ushort)code;
                    header.address = (UInt32)(num * sizeof(UInt32));
                    // ヘッダー書き込み.
                    byte[] headerByte = header.getByte();
                    fs.Write(headerByte, 0, headerByte.Length);

                    // サイズはuint32固定
                    UInt32 offset = 0;
                    ArrayList convertList = new ArrayList();
                    for (int i = 0; i < num; ++i)
                    {
                        byte[] offsetData = BitConverter.GetBytes(offset);
                        fs.Write(offsetData, 0, offsetData.Length);
                        byte[] data = null;//System.Text.Encoding.UTF8.GetBytes(stringList[i]);     // utf-8
                        switch (code)
                        {
                            //                    byte[] data = System.Text.Encoding.ASCII.GetBytes(paramList[column][i]);  // ascii
                            case STRBCodePage.STRB_CODE_PAGE_SJIS:
                                data = System.Text.Encoding.GetEncoding("SHIFT-JIS").GetBytes(stringList[i]);  // shift-jis
                                break;
                            case STRBCodePage.STRB_CODE_PAGE_UTF8:
                                data = System.Text.Encoding.UTF8.GetBytes(stringList[i]);     // utf-8
                                break;
                            default:
                                throw new Exception("STRB Code Page Error type : " + code);
                        }
                        convertList.Add(data);
                        offset += (UInt32)(data.Length + 1);
                    }

                    byte[] outData = new byte[offset];
                    Array.Clear(outData, 0, outData.Length);   // 0クリア
                    offset = 0;
                    for (int i = 0; i < num; ++i)
                    {
                        byte[] data = (byte[])convertList[i];
                        data.CopyTo(outData, offset);
                        offset += (UInt32)(data.Length + 1);
                    }
                    fs.Write(outData, 0, outData.Length);
                    Console.WriteLine("strb : published. " + outputPath);
                }

                result = true;
            }
            catch (Exception e)
            {
                result = false;
                throw e;
            }
            finally
            {
                if (null != fs)
                {
                    fs.Close();
                    fs.Dispose();
                    fs = null;
                }
            }

            return result;
        }

        // VALBで出力
        public static bool publishValBinaly(string filepath, ref Analyse analyse)
        {
            bool result = false;
            FileStream fs = null;
            VALBBinalyHeader header = new VALBBinalyHeader();

            Dictionary<int, List<String>> paramList = analyse.ParamList;
            List<BinaryKey> keyList = analyse.HeaderKeyList;

            uint size = 0;
            for (int i = 0; i < keyList.Count; ++i)
            {
                if ((int)KeyType.KEY_EXTENSION == keyList[i].type)
                {   // 拡張キーは保存しない.
                    continue;
                }
                size += (uint)keyList[i].size;
            }
            int num = paramList.First().Value.Count;
            header.dataSize = size;
            header.dataNum = (uint)num;

            // デバッグ用
            int currentIndex = 0;
            int currentKey = 0;
            //
            try
            {
                // ファイル作成.
                fs = new FileStream(filepath, FileMode.Create);
                byte[] headerArray = header.getByte();
                fs.Write(headerArray, 0, headerArray.Length);

                byte[] bytes = new byte[header.dataSize];
                int offset = 0;
                for (int i = 0; i < num; ++i)
                {
                    offset = 0;
                    currentIndex = i;
                    Array.Clear(bytes, 0, bytes.Length);    // 0クリアする.
                    for (int j = 0; j < keyList.Count; ++j)
                    {
                        currentKey = j;
                        string param = "";
                        if (0 <= keyList[j].column)
                        {   // paddingはparamを取らない.
                            param = paramList[keyList[j].column][i];
                        }
                        switch (keyList[j].type)
                        {
                            case (int)KeyType.KEY_S8: BitConverter.GetBytes(System.Convert.ToSByte(param)).CopyTo(bytes, offset); offset += sizeof(sbyte);  break;
                            case (int)KeyType.KEY_S16: BitConverter.GetBytes(System.Convert.ToInt16(param)).CopyTo(bytes, offset); offset += sizeof(Int16); break;
                            case (int)KeyType.KEY_S32: BitConverter.GetBytes(System.Convert.ToInt32(param)).CopyTo(bytes, offset); offset += sizeof(Int32); break;
                            case (int)KeyType.KEY_S64: BitConverter.GetBytes(System.Convert.ToInt64(param)).CopyTo(bytes, offset); offset += sizeof(Int64); break;
                            case (int)KeyType.KEY_U8: BitConverter.GetBytes(System.Convert.ToByte(param)).CopyTo(bytes, offset); offset += sizeof(byte); break;
                            case (int)KeyType.KEY_U16: BitConverter.GetBytes(System.Convert.ToUInt16(param)).CopyTo(bytes, offset); offset += sizeof(UInt16); break;
                            case (int)KeyType.KEY_U32: BitConverter.GetBytes(System.Convert.ToUInt32(param)).CopyTo(bytes, offset); offset += sizeof(UInt32); break;
                            case (int)KeyType.KEY_U64: BitConverter.GetBytes(System.Convert.ToUInt64(param)).CopyTo(bytes, offset); offset += sizeof(UInt64); break;
                            case (int)KeyType.KEY_DOUBLE: BitConverter.GetBytes(System.Convert.ToDouble(param)).CopyTo(bytes, offset); offset += sizeof(Double); break;
                            case (int)KeyType.KEY_FLOAT:
                                {   // 一旦floatにキャスト.
                                    float tmp = (float)System.Convert.ToDouble(param);
                                    BitConverter.GetBytes(tmp).CopyTo(bytes, offset);
                                    offset += sizeof(float);
                                    break;
                                }

                            case (int)KeyType.KEY_FIXED_STRING:
                                {
                                    //サイズ固定の為に別にバッファを確保する.
                                    byte[] txt = System.Text.Encoding.UTF8.GetBytes(param);
                                    txt.CopyTo(bytes, offset);   // バッファ以上の文字列が出力されようとした場合はここでエラーになる.
                                    offset += keyList[j].size;
                                }
                                break;
                            case (int)KeyType.KEY_STRING: 
                                throw new Exception("valb : The size of the string needs to be fixed." + "key:" + keyList[j].key);

                            case (int)KeyType.KEY_PADDING:
                                {
                                    // 書き込むデータはないのでオフセットだけずらす.
                                    offset += keyList[j].size;
                                }
                                break;
                            case (int)KeyType.KEY_EXTENSION:
                                // 拡張キーは無視.
                                break;
                            default:
                                throw new Exception("It can not output type is specified : " + keyList[j].key);
                        }
                    }

                    // 作成された構造体を書き込む.
                    fs.Write(bytes, 0, bytes.Length);
                }

                Console.WriteLine("valb : published. " + filepath);
                result = true;
            }
            catch (Exception e)
            {
                result = false;
                Console.WriteLine( "valb error : " + e.Message + "line : " + currentIndex + "key : " + keyList[currentKey].key );
                throw;
            }
            finally
            {
                if (null != fs)
                {
                    fs.Close();
                    fs.Dispose();
                    fs = null;
                }
            }

            return result;
        }

        /// <summary>
        // HTBの出力
        /// </summary>
        public static bool publishHTBinaly(string filePath, string headerPath, ref Analyse analyse, ref HashPublishIni ini)
        {
            bool result = true;
            Dictionary<int, List<String>> paramList = analyse.ParamList;
            List<BinaryKey> keyList = analyse.HeaderKeyList;
            try
            {
                Dictionary<Object, string> hashDictionary = new Dictionary<Object, string>();
                string hashKey = ini.hash;
                string commentKey = ini.comment;
                int hashColumn = keyList.Find(x => 0 == x.columnName.CompareTo(hashKey)).column;
                int commentColumn = keyList.Find(x => 0 == x.columnName.CompareTo(commentKey)).column;

                List<HTBOutputData> hashList = new List<HTBOutputData>();
                List<String> nameList = paramList[hashColumn];
                List<String> commentList = paramList[commentColumn];
                int num = paramList.First().Value.Count;
                for (int i = 0; i < num; ++i)
                {
                    String name = nameList[i];
                    if ((null == name) ||
                        (0 >= name.Length))
                    {
                        if (!ini.empty )
                        {   // hashのないアイテムは出力しない.
                            Console.WriteLine("emptyにfalseが指定されていますが、Hashに出力する文字列がみつかりません");
                            Console.WriteLine(ini.hash + ":" + i);
                            throw new Exception("hash item not found.");
                        }
                        // 0ハッシュが許可されている場合は次の要素へ.
                        continue;
                    }
                    Object param;
                    if ((int)HashBit.BIT32 == ini.hashBit)
                    {
                        param = FnvHash.getFNV_1_32(name);
                        if (0 == (Int32)param)
                        {
                            Console.WriteLine("0ハッシュが見つかりました!!!指定するデータ名を変更してください!!");
                            Console.WriteLine("name : " + name);
                            continue;
                        }
                    }
                    else
                    {
                        param = FnvHash.getFNV_1_64(name);
                        if (0 == (Int64)param)
                        {
                            Console.WriteLine("0ハッシュが見つかりました!!!指定するデータ名を変更してください!!");
                            Console.WriteLine("name : " + name);
                            continue;
                        }
                    }

                    if (hashDictionary.ContainsKey(param))
                    {
                        //dictionaly[data.hash].Add(data);
                        Console.WriteLine("Hashが重複しています。");
                        Console.WriteLine("key1 : " + name + " key2 : " + hashDictionary[param] + " hash :" + param);
                        result = false;
                        break;
                    }
                    else
                    {
                        hashDictionary.Add(param, name);
                    }

                    HTBOutputData hash = new HTBOutputData();
                    hash.hash = param;
                    hash.index = i;
                    hash.name = name;
                    hash.comment = commentList[i].Split('\n')[0];   // 複数行のコメントは最初の1行のみ

                    hashList.Add(hash);
                }

                // データ、ヘッダ、xmlを出力する.
                publishHTBBin(filePath, ini.hashBit, ref hashList);
                publishHTBHeader(headerPath, ini.hashBit, ref hashList);

                result = true;
//                Console.WriteLine( "HTB : published." + filePath);

            }
            catch (Exception e)
            {
                Console.WriteLine("HTB Error : " + e.Message);
                result = false;
                throw;
            }
            finally
            {
                // 特になし.
            }


            return result;
        }


        private static bool publishHTBBin(string filePath, int bit, ref List<HTBOutputData> outputList)
        {
            FileStream fs = null;

            // 必要なデータのみを抽出してソートをする
            List<HTBOutputData> binList = outputList.FindAll(x => bit == (int)HashBit.BIT32 ? 0 != (Int32)x.hash : 0 != (Int64)x.hash);
            binList.Sort(delegate(HTBOutputData a, HTBOutputData b)
            {
                if (bit == (int)HashBit.BIT32)
                {
                    if ((Int32)a.hash < (Int32)b.hash)
                    {
                        return -1;
                    }
                    else if ((Int32)a.hash > (Int32)b.hash)
                    {
                        return 1;
                    }
                }
                else
                {
                    if ( (Int64)a.hash < (Int64)b.hash )
                    {
                        return -1;
                    }
                    else if ( (Int64)a.hash > (Int64)b.hash )
                    {
                        return 1;
                    }
                }
                return 0;
            });

            try
            {
                fs = new FileStream(filePath, FileMode.Create);
                byte[] byteArry = null;

                HTBBinalyHeader header = new HTBBinalyHeader();
                // ヘッダー情報を保存
                header.hashnum = (ushort)binList.Count;
                header.bitType = (ushort)bit;
                byteArry = header.getByte();
                fs.Write(byteArry, 0, byteArry.Length);

                // ハッシュリスト保存
                for (int i = 0; i < binList.Count; ++i)
                {
                    if ((int)HashBit.BIT32 == bit)
                    {
                        HashIndexPair32 pair = new HashIndexPair32();
                        pair.hash = (Int32)binList[i].hash;
                        pair.index = binList[i].index;
                        byteArry = pair.getByte();

                    }
                    else if ((int)HashBit.BIT64 == bit)
                    {
                        HashIndexPair64 pair = new HashIndexPair64();
                        pair.hash = (Int64)binList[i].hash;
                        pair.index = binList[i].index;
                        byteArry = pair.getByte();
                    }
                    fs.Write(byteArry, 0, byteArry.Length);
                }
                Console.WriteLine("htb : {0}個のデータを出力しました", binList.Count);
                Console.WriteLine("htb : {0}を出力しました", filePath);
            }
            catch
            {
                throw;
            }
            finally
            {
                // 終了処理
                if (null != fs)
                {
                    fs.Close();
                    fs.Dispose();
                    fs = null;
                }
                binList.Clear();
            }

            return true;
        }

        private static bool publishHTBHeader(string filePath, int bit, ref List<HTBOutputData> outputList)
        {
            bool result = true;
            StreamWriter sw = null;
            try
            {
                // 必要なデータのみを抽出してソートをする
                List<HTBOutputData> binList = outputList.FindAll(x => (bit == (int)HashBit.BIT32 ? 0 != (Int32)x.hash : 0 != (Int64)x.hash));
                sw = new StreamWriter(filePath, false, Encoding.UTF8);
                for (int i = 0; i < outputList.Count; ++i)
                {
                    string tmpLine = "#define " + outputList[i].name + " " + "(" + outputList[i].hash + ")";
                    addComment(ref tmpLine, outputList[i].comment);
                    sw.WriteLine(tmpLine);
                }
                Console.WriteLine("header : {0}個のデータを出力しました", outputList.Count);
            }
            catch
            {
                result = false;
                throw;
            }
            finally
            {
                if (null != sw)
                {
                    sw.Close();
                    sw.Dispose();
                    sw = null;
                }
            }

            return result;
        }

        /// <summary>
        /// テキストファイルをxmlで出力します(全言語分+ハッシュ)
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="analyse"></param>
        /// <returns></returns>
        public static bool publishStringXml(string filePath, String langage, String nameTag, ref Analyse analyse)
        {
            bool result = false;
            XmlTextWriter writer = null;
            try
            {
                Dictionary<int, List<String>> paramList = analyse.ParamList;
                List<BinaryKey> keyList = analyse.HeaderKeyList;

                List<String> nameList = null;
                // 名前(hash用)のキーを選択
                BinaryKey nameKey = keyList.Find(x => 0 == x.columnName.CompareTo(nameTag));
                nameList = paramList[nameKey.column];
                // UTF-8でエレメントの情報をxml出力
                writer = new XmlTextWriter(
                    filePath,
                    Encoding.UTF8);
                writer.Formatting = Formatting.Indented;
                // document開始
                writer.WriteStartDocument();
                writer.WriteStartElement("MsgList");

                int num = nameList.Count;
                string[] keys = langage.Split(',');
                for (int i = 0; i < num; ++i)
                {
                    string name = nameList[i];
                    writer.WriteStartElement("Msg");

                    if (null != name && 0 < name.Length)
                    {

                        writer.WriteAttributeString("name", name);
                        writer.WriteAttributeString("hash", FnvHash.getFNV_1_32(name).ToString());
                        writer.WriteAttributeString("hash64", FnvHash.getFNV_1_64(name).ToString());
                    }

                    for (int j = 0; j < keys.Length; ++j)
                    {
                        BinaryKey key = keyList.Find(x => 0 == x.key.CompareTo(keys[j]));
                        string keyName = "msg_" + key.key;
                        writer.WriteAttributeString(keyName, paramList[key.column][i]);
                    }
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();   // MsgList
                writer.WriteEndDocument();  // document終了
                writer.Flush();

                result = true;
                Console.WriteLine("String Xml : published." + filePath);

            }
            catch (Exception e)
            {
                Console.WriteLine("String Xml : " + e.Source + e.Message);
                throw;
            }
            finally
            {
                if ( null != writer)
                {
                    writer.Close();
                    writer = null;
                }
            }

            return result;
        }
        public static bool publishEnum(string filePath, string id, string define, string comment, string type, ref Analyse analyse)
        {
            bool result = true;
            StreamWriter sw = null;
            try
            {
                EnumType enumType = EnumType.ENUM_TYPE_ENUM;
                if (null != type)
                {
                    switch (type.ToLower())
                    {
                        default:    // 存在しない場合もENUMタイプで出力
                        case "enum":
                            enumType = EnumType.ENUM_TYPE_ENUM;
                            break;
                        case "define":
                            enumType = EnumType.ENUM_TYPE_DEFINE;
                            break;
                    }
                }

                Dictionary<int, List<String>> paramList = analyse.ParamList;
                List<BinaryKey> keyList = analyse.HeaderKeyList;

                List<String> idList = null;
                List<String> defineList = null;
                List<String> commentList = null;
                // 名前(hash用)のキーを選択
                BinaryKey idKey = keyList.Find(x => 0 == x.columnName.CompareTo(id));
                BinaryKey defineKey = keyList.Find(x => 0 == x.columnName.CompareTo(define));
                BinaryKey commentKey = keyList.Find(x => 0 == x.columnName.CompareTo(comment));
                idList = paramList[idKey.column];
                defineList = paramList[defineKey.column];
                if ((int)KeyType.KEY_INVALID != commentKey.type)
                {
                    commentList = paramList[commentKey.column];
                }

                // 必要なデータのみを抽出してソートをする
                sw = new StreamWriter(filePath, false, Encoding.UTF8);
                int num = paramList.First().Value.Count;
                for (int i = 0; i < num; ++i)
                {
                    if ( 0 == defineList[i].Length)
                    {   // 0は出力しない.
                        continue;
                    }

                    if (0 == idList[i].Length)
                    {   // IDが無い.
                        throw new Exception("EnumのIDが設定されていません。" + "line : " + i + " define : " + defineList[i]);
                    }
                    if (!System.Text.RegularExpressions.Regex.IsMatch(
                       idList[i],
                       @"^[_:0-9\*\[\]]+$"))
                    {   // 数値以外が指定されている
                        throw new Exception("EnumのIDに数値以外を指定できません。" + "line : " + i + " define : " + defineList[i]);
                    }

                    string tmpLine = "";
                    switch( enumType)
                    {
                        case EnumType.ENUM_TYPE_ENUM:
                            tmpLine = defineList[i] + " = " + idList[i] + ",";
                            break;
                        case EnumType.ENUM_TYPE_DEFINE:
                            tmpLine = "#define " + defineList[i] + " " + "(" + idList[i] + ")";
                            break;
                        default:
                            throw new Exception("Publish Enum Error : Enum出力のタイプが不正です\n");
                    }
                    if (null != commentList)
                    {
                        addComment(ref tmpLine, commentList[i].Split('\n')[0]);
                    }
                    sw.WriteLine(tmpLine);
                }
                Console.WriteLine("enum : {0}個のデータを出力しました", num);
                Console.WriteLine("enum : published." + filePath);

                // 終了した.
                result = true;
            }
            catch
            {
                result = false;
                throw;
            }
            finally
            {
                if (null != sw)
                {
                    sw.Close();
                    sw.Dispose();
                    sw = null;
                }
            }

            return result;
        }
    }
}
