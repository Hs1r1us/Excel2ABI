using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel2ABI
{
    class Program
    {
        static List<string> sourceList, modeList, writeList;
        static string boardName;
        static bool is384Board, isSingle;
        static string outPutName;
        static List<string> errorMsg = new List<string>();
        static private bool NameErrorFlag;

        [STAThread]
        static void Main(string[] args)
        {
            string excelFile;
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\96.txt") && File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\384.txt"))
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("\nExcel not drop");
                    var openFile = new OpenFileDialog();
                    openFile.Filter = "Excel|*.xlsx";
                    openFile.RestoreDirectory = false;
                    openFile.Multiselect = false;
                    openFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    if (openFile.ShowDialog() == DialogResult.OK)
                    {
                        excelFile = openFile.FileName;
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    excelFile = args[0];
                }

                Console.Write("\nInput BoardName (Default: Y48-LG-1031):");
                string ReadName = Console.ReadLine();
                boardName = ReadName.Equals("") ? "Y48-LG-1031" : ReadName;
                
                Console.Write("\n384 ? 'y' or 'n' (Default: n):");
                string temp = Console.ReadLine();
                is384Board = isSingle = temp.ToLower().Contains("y");

                ExcelWorker(excelFile);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("\nMode file not found.");
                Console.Read();
            }
        }

        private static void ExcelWorker(string filePath)
        {
            string boardSuffix;
            ReadMode();
            ClearFolder();
            try {
                using (var fs = File.OpenRead(filePath))
                {
                    var wk = new XSSFWorkbook(fs);
                    var sheet = wk.GetSheetAt(0);
                    for (var i = 0; ; i++)
                    {
                        sourceList = new List<string>();

                        ICell tempVar = null;
                        var tempRow = sheet.GetRow(0);
                        if (tempRow != null)
                        {
                            tempVar = tempRow.GetCell(i);
                        }

                        if (tempVar != null)
                        {
                            if (tempVar.CellType == CellType.Formula)
                            {
                                tempVar.SetCellType(CellType.String);
                                boardSuffix = tempVar.StringCellValue;
                            }
                            else
                            {
                                boardSuffix = tempVar.ToString();
                            }
                        }
                        else
                        {
                            boardSuffix = (i + 1).ToString();
                        }

                        if (!CheckFileName(boardSuffix))
                        {
                            boardSuffix = (i + 1).ToString();
                            NameErrorFlag = true;
                        }

                        if (isSingle)
                        {
                            outPutName = boardSuffix;
                        }

                        for (var j = 0; j < 92; j++)
                        {
                            ICell cell = null;
                            var row = sheet.GetRow(j + 1);

                            if (row != null)
                            {
                                cell = row.GetCell(i);
                            }
                            string str;

                            if (cell == null)
                            {
                                //判断处于当列首行为空 && 下列首行为空 && 384完成输出
                                if (j == 0 && (sheet.GetRow(j + 1) == null || sheet.GetRow(j + 1).GetCell(i + 1) == null) && (!is384Board || isSingle))
                                {
                                    Console.WriteLine("");
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    foreach (string errorStr in errorMsg)
                                    {
                                        Console.WriteLine(errorStr);
                                    }
                                    Console.WriteLine((is384Board ? i / 2 : i) + " files output successful");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.ReadLine();
                                    return;
                                }

                                sourceList.Add("");
                                continue;
                            }

                            if (cell.CellType == CellType.Formula)
                            {
                                cell.SetCellType(CellType.String);
                                str = cell.StringCellValue;
                            }
                            else
                            {
                                str = cell.ToString();
                            }
                            sourceList.Add(str);
                        }

                        if (NameErrorFlag)
                        {
                            errorMsg.Add("第" + (i + 1) + "列命名非法，已更名为" + (i + 1));
                            NameErrorFlag = false;
                        }

                        TXTWorker(boardSuffix);
                    }
                }
            }catch(Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.Message);
                Console.ForegroundColor = ConsoleColor.White;
                Console.ReadLine();

            }
        }

        private static void ClearFolder()
        {
            var folder = AppDomain.CurrentDomain.BaseDirectory + "\\Create\\";
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
                while (Directory.Exists(folder)) ;
            }
            Directory.CreateDirectory(folder);
        }

        private static void TXTWorker(string suffix)
        {
            if (!is384Board || isSingle)
            {
                writeList = new List<string>();
                modeList.ForEach(i => writeList.Add(i));
            }

            if (is384Board)
            {
                if (isSingle)
                {
                    writeList[1] = writeList[1].Insert(1, boardName + "-" + suffix);
                    writeList[1] = writeList[1].Insert(0, boardName + "-" + suffix);

                    for (int i = 5; i < 18; i++)
                    {
                        int lineNum = i * 2 - 5;

                        if (!sourceList[i - 5].Equals(""))
                        {
                            writeList[lineNum] = writeList[lineNum].Insert(5, suffix);
                            writeList[lineNum] = writeList[lineNum].Insert(4, sourceList[i - 5]);
                        }
                        else
                        {
                            writeList[lineNum] = "NOTEXT";
                        }
                    }

                    for(int i = 18; i < 21; i++)
                    {
                        int lineNum = i * 2 - 5;

                        writeList[lineNum] = writeList[lineNum].Insert(writeList[lineNum].IndexOf('\t', writeList[lineNum].IndexOf('\t') + 1) + 1, suffix);
                    }

                    for (int i = 21; i < 100; i++)
                    {
                        int lineNum = i * 2 - 5;

                        if (!sourceList[i - 8].Equals(""))
                        {
                            writeList[lineNum] = writeList[lineNum].Insert(5, suffix);
                            writeList[lineNum] = writeList[lineNum].Insert(4, sourceList[i - 8]);
                        }
                        else
                        {
                            writeList[lineNum] = "NOTEXT";
                        }
                    }

                    writeList[195] = writeList[195].Insert(writeList[195].IndexOf('\t', writeList[195].IndexOf('\t') + 1) + 1, suffix);

                    isSingle = false;
                }
                else
                {
                    writeList[1] = writeList[1].Insert(writeList[1].IndexOf('\t', writeList[1].IndexOf('\t') + 1), "-" + suffix);
                    writeList[1] = writeList[1].Insert(writeList[1].IndexOf('\t'), "-" + suffix);

                    for (int i = 5; i < 18; i++)
                    {
                        int lineNum = i * 2 - 4;

                        if (!sourceList[i - 5].Equals(""))
                        {
                            writeList[lineNum] = writeList[lineNum].Insert(5, suffix);
                            writeList[lineNum] = writeList[lineNum].Insert(4, sourceList[i - 5]);
                        }
                        else
                        {
                            writeList[lineNum] = "NOTEXT";
                        }
                    }

                    for (int i = 18; i < 21; i++)
                    {
                        int lineNum = i * 2 - 4;

                        writeList[lineNum] = writeList[lineNum].Insert(writeList[lineNum].IndexOf('\t', writeList[lineNum].IndexOf('\t') + 1) + 1, suffix);
                    }

                    for (int i = 21; i < 100; i++)
                    {
                        int lineNum = i * 2 - 4;

                        if (!sourceList[i - 8].Equals(""))
                        {
                            writeList[lineNum] = writeList[lineNum].Insert(5, suffix);
                            writeList[lineNum] = writeList[lineNum].Insert(4, sourceList[i - 8]);
                        }
                        else
                        {
                            writeList[lineNum] = "NOTEXT";
                        }
                    }

                    writeList[196] = writeList[196].Insert(writeList[196].IndexOf('\t', writeList[196].IndexOf('\t') + 1) + 1, suffix);

                    isSingle = true;

                    OutPutTXT(outPutName);
                }
            }
            else
            {
                writeList[1] = writeList[1].Insert(1, boardName + "-" + suffix);
                writeList[1] = writeList[1].Insert(0, boardName + "-" + suffix);

                for (int i = 5; i < 18; i++)
                {
                    if (!sourceList[i - 5].Equals(""))
                    {
                        writeList[i] = writeList[i].Insert(5, suffix);
                        writeList[i] = writeList[i].Insert(4, sourceList[i - 5]);
                    }
                    else
                    {
                        writeList[i] = "NOTEXT";
                    }
                }

                for (int i = 18; i < 20; i++)
                {
                    writeList[i] = writeList[i].Insert(6, suffix);
                }

                writeList[20] = writeList[20].Insert(11, suffix);

                for (int i = 21; i < 100; i++)
                {
                    if (!sourceList[i - 8].Equals(""))
                    {
                        writeList[i] = writeList[i].Insert(5, suffix);
                        writeList[i] = writeList[i].Insert(4, sourceList[i - 8]);
                    }
                    else
                    {
                        writeList[i] = "NOTEXT";
                    }
                }

                writeList[100] = writeList[100].Insert(11, suffix);

                OutPutTXT(suffix);
            }
        }

        private static void OutPutTXT(string suffix)
        {
            string file = SameNameHandle(AppDomain.CurrentDomain.BaseDirectory + "\\Create" + "\\" + suffix.ToString() + ".txt", ".txt");

            using (var myStream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite))
            {
                using (var myWrite = new StreamWriter(myStream))
                {
                    for (var i = 0; i < writeList.Count; i++)
                    {
                        if (writeList[i].Equals("NOTEXT"))
                            continue;
                        myWrite.WriteLine(writeList[i]);
                    }
                }
            }
        }

        private static void ReadMode()
        {
            string path = is384Board ? AppDomain.CurrentDomain.BaseDirectory + "\\384.txt" : AppDomain.CurrentDomain.BaseDirectory + "\\96.txt";
            modeList = new List<string>();
            using (var sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    var lineContent = sr.ReadLine();
                    modeList.Add(lineContent);
                }
            }
        }

        /// <summary>
        /// 输出文件重名处理
        /// </summary>
        /// <param name="inputStr"></param>
        /// <returns></returns>
        private static string SameNameHandle(string inputStr, string extenStr)
        {
            string result = inputStr;

            int num = 0;
            while (File.Exists(result))
            {

                errorMsg.Add("文件" + inputStr + "重名，尝试更名为" + result);
                result = inputStr.Insert(inputStr.Length - extenStr.Length, "-" + ++num);
            }
            return result;
        }

        /// <summary>
        /// 文件名检查
        /// </summary>
        /// <returns></returns>
        public static Boolean CheckFileName(string fileName)

        {

            StringBuilder description = new StringBuilder();
            Boolean opResult = Regex.IsMatch(fileName, @"(?!((^(con)$)|^(con)\\..*|(^(prn)$)|^(prn)\\..*|(^(aux)$)|^(aux)\\..*|(^(nul)$)|^(nul)\\..*|(^(com)[1-9]$)|^(com)[1-9]\\..*|(^(lpt)[1-9]$)|^(lpt)[1-9]\\..*)|^\\s+|.*\\s$)(^[^\\\\\\/\\:\\<\\>\\*\\?\\\\\\""\\\\|]{1,255}$)");

            return opResult;
        }
    }
}