using System;
using ExcelDataReader;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml;
using PasswordFinder.Properties;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Concurrent;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Crypt;
using NPOI.HWPF.Extractor;
using NPOI.XWPF.UserModel;


namespace PasswordFinder
{
    public partial class Form1 : Form
    {
        private CancellationTokenSource _cts;

        public Form1()
        {
            InitializeComponent();
        }
        private List<string> RemoveDuplicate(List<string> sourceList)
        {
            if (sourceList == null) return new List<string>();

            return sourceList
                .Where(f => !f.Contains("~$"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(f => f)
                .ToList();
        }

        /// <summary>
        /// Дополняет базовую строку выделенными галками
        /// </summary>
        /// <param name="base_Text"></param>
        /// <returns></returns>
        private string Prepare_File_patterns(string base_Text)
        {
            string pat = base_Text;
            if (docBox.Checked)
                pat += ",*.doc,*.docx";
            if (xlsBox.Checked)
                pat += ",*.xls,*.xlsx";
            if (txtBox.Checked)
                pat += ",*.txt";
            if (rtfBox.Checked)
                pat += ",*.rtf";
            if (xmlBox.Checked)
                pat += ",*.xml";
            return pat;
        }

        private IEnumerable<string> GetFilesSafe(string root, System.Text.RegularExpressions.Regex patternRegex, CancellationToken ct)
        {
            var pending = new Stack<string>();
            pending.Push(root);
            while (pending.Count > 0)
            {
                string path = pending.Pop();

                ct.ThrowIfCancellationRequested();

                // Enumerate files in current directory
                IEnumerable<string> files = null;
                try { files = Directory.EnumerateFiles(path, "*", SearchOption.TopDirectoryOnly); }
                catch { }

                if (files != null)
                {
                    foreach (var file in files)
                    {
                        string fileName = Path.GetFileName(file);
                        if (fileName.StartsWith("~$")) continue;
                        if (patternRegex.IsMatch(fileName)) yield return file;
                    }
                }

                // Enumerate subdirectories
                IEnumerable<string> subDirs = null;
                try { subDirs = Directory.EnumerateDirectories(path, "*", SearchOption.TopDirectoryOnly); }
                catch { }

                if (subDirs != null)
                {
                    foreach (var subDir in subDirs) { pending.Push(subDir); }
                }
            }
        }

        private async void Button1_Click(object sender, EventArgs e)
        {
            if (_cts != null)
            {
                _cts.Cancel();
                return;
            }

            string[] Directories = richTextBox1.Lines;
            string[] File_patterns = Prepare_File_patterns(richTextBox2.Text).Split(',');
            string[] PWD_patterns = richTextBox3.Text.Split(',');

            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            string originalButtonText = button1.Text;
            button1.Text = (rusToolStripMenuItem.Checked || label3.Text.Contains("Папки")) ? "Остановить" : "Stop";
            
            listView1.Items.Clear();
            toolStripStatusLabel1.Text = "Инициализация поиска...";
            try
            {
            await Task.Run(() =>
            {
                List<string> Files = new List<string>();
                ConcurrentBag<string> Files_with_password = new ConcurrentBag<string>();

                var patterns = File_patterns
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Select(p => p.Trim())
                    .ToList();

                if (patterns.Count == 0) return;

                // Create a Regex for efficient single-pass matching of all patterns
                string regexString = "^(" + string.Join("|", patterns.Select(p =>
                    System.Text.RegularExpressions.Regex.Escape(p)
                        .Replace(@"\*", ".*")
                        .Replace(@"\?", ".")
                )) + ")$";

                var patternRegex = new System.Text.RegularExpressions.Regex(regexString,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Поиск файлов во всех указанных директориях
                for (int i = 0; i < Directories.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(Directories[i]) || !Directory.Exists(Directories[i])) continue;
                    Files.AddRange(GetFilesSafe(Directories[i], patternRegex, token));
                }

                Files = RemoveDuplicate(Files);//удалить дубликаты найденные по разным правилам

                int processedCount = 0;
                Parallel.ForEach(Files, new ParallelOptions 
                { 
                    CancellationToken = token,
                    MaxDegreeOfParallelism = Environment.ProcessorCount 
                }, currentFile =>
                {
                    token.ThrowIfCancellationRequested();

                    string ext = Path.GetExtension(currentFile).ToLower();
                    int currentProgress = Interlocked.Increment(ref processedCount);
                    this.Invoke((MethodInvoker)delegate { toolStripStatusLabel1.Text = string.Format("{0} / {1}", currentProgress, Files.Count); });

                    // Generate a list of potential passwords to try (null/empty + known search patterns)
                    var passwordsToTry = new List<string> { null };
                    passwordsToTry.AddRange(PWD_patterns.Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => p.Trim()));
                    var uniquePasswords = passwordsToTry.Distinct().ToList();

                    bool isexit = false;
                    if (currentFile.Contains(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            using (FileStream stream = File.Open(currentFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                foreach (var pass in uniquePasswords)
                                {
                                    try
                                    {
                                        stream.Position = 0;
                                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { Password = pass }))
                                        {
                                            DataSet result = reader.AsDataSet();
                                            for (int t = 0; t < result.Tables.Count; t++)
                                            {
                                                for (int p = 0; p < PWD_patterns.Length; p++)
                                                {
                                                    int columns_c = result.Tables[t].Columns.Count;
                                                    for (int c = 0; c < columns_c; c++)
                                                    {
                                                        DataRow[] range = result.Tables[t].Select(result.Tables[t].Columns[c].ColumnName + " like '%" + PWD_patterns[p] + "%'");
                                                        if (range != null && range.Length != 0)
                                                        {
                                                            for (int j = 0; j < range.Length; j++)
                                                            {
                                                                string lineData = String.Join(" ", range[j].ItemArray);
                                                                this.Invoke((MethodInvoker)delegate
                                                                {
                                                                    listView1.Items.Add(new ListViewItem(new string[] { currentFile, lineData, ext }));
                                                                });
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        break; // If we reached here, the password was correct (or not needed)
                                    }
                                    catch (ExcelDataReader.Exceptions.HeaderException ex) when (ex.Message.Contains("password")) { continue; }
                                    catch { break; }
                                }
                            }
                        }
                        catch { }
                        isexit = true;
                    }
                    if (currentFile.Contains(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            foreach (var pass in uniquePasswords)
                            {
                                try
                                {
                                    if (pass == null) // Try normal open first
                                    {
                                        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(currentFile, false))
                                        {
                                            var body = wordDoc.MainDocumentPart?.Document?.Body;
                                            if (body != null)
                                            {
                                                foreach (var p in PWD_patterns)
                                                {
                                                    var matches = body.Descendants<Paragraph>().Where(para => para.InnerText.Contains(p, StringComparison.OrdinalIgnoreCase));
                                                    foreach (var match in matches)
                                                    {
                                                        string context = match.InnerText;
                                                        this.Invoke((MethodInvoker)delegate { listView1.Items.Add(new ListViewItem(new string[] { currentFile, context.Trim(), ext })); });
                                                    }
                                                }
                                            }
                                        }
                                        break;
                                    }
                                    else // Try opening as an encrypted OOXML package using NPOI
                                    {
                                        using (FileStream fs = File.Open(currentFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                        {
                                            var nfs = new NPOI.POIFS.FileSystem.POIFSFileSystem(fs);
                                            var info = new NPOI.POIFS.Crypt.EncryptionInfo(nfs);
                                            var d = NPOI.POIFS.Crypt.Decryptor.GetInstance(info);
                                            if (d.VerifyPassword(pass))
                                            {
                                                using (var dataStream = d.GetDataStream(nfs))
                                                {
                                                    var xdoc = new NPOI.XWPF.UserModel.XWPFDocument(dataStream);
                                                    foreach (var p in PWD_patterns)
                                                    {
                                                        foreach (var para in xdoc.Paragraphs)
                                                            if (para.ParagraphText.Contains(p, StringComparison.OrdinalIgnoreCase))
                                                                this.Invoke((MethodInvoker)delegate { listView1.Items.Add(new ListViewItem(new string[] { currentFile, para.ParagraphText.Trim(), ext })); });
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                                catch { continue; }
                            }
                        }
                        catch { }
                        isexit = true;
                    }
                    if (currentFile.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            using (FileStream fs = new FileStream(currentFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                WordExtractor extractor = new WordExtractor(fs);
                                string[] paragraphs = extractor.ParagraphText;

                                if (paragraphs != null)
                                {
                                    foreach (var p in PWD_patterns)
                                    {
                                        foreach (var paraText in paragraphs)
                                        {
                                            if (paraText.Contains(p, StringComparison.OrdinalIgnoreCase))
                                            {
                                                this.Invoke((MethodInvoker)delegate
                                                {
                                                    listView1.Items.Add(new ListViewItem(new string[] { currentFile, paraText.Trim(), ext }));
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                        isexit = true;
                    }
                    if (currentFile.Contains(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            XmlDocument doc = new XmlDocument();
                            doc.Load(currentFile);
                            for (int p = 0; p < PWD_patterns.Length; p++)
                            {
                                XmlNodeList elemList = doc.GetElementsByTagName(PWD_patterns[p]);
                                for (int i = 0; i < elemList.Count; i++)
                                {
                                    string xmlString = elemList[i].OuterXml;
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        listView1.Items.Add(new ListViewItem(new string[] { currentFile, xmlString, ext }));
                                    });
                                }
                            }
                        }
                        catch { }
                        isexit = true;
                    }
                    if (!isexit)// Поиск по текстовому содержимому
                    {
                        try
                        {
                            string[] Lines = File.ReadAllLines(currentFile, Encoding.Default);
                            for (int l = 0; l < Lines.Length; l++)
                            {
                                for (int p = 0; p < PWD_patterns.Length; p++)
                                {
                                    if (Lines[l].Contains(PWD_patterns[p], StringComparison.OrdinalIgnoreCase))
                                    {
                                        Files_with_password.Add(currentFile);
                                        string[] Words = Lines[l].Split(' ');
                                        for (int w = 0; w < Words.Length; w++)
                                        {
                                            if (Words[w].Contains(PWD_patterns[p]))
                                            {
                                                string findedpwd = string.Empty;
                                                int w2 = w;
                                                if ((w2 - 2) >= 0) findedpwd += Words[w2 - 2] + " ";
                                                if ((w2 - 1) >= 0) findedpwd += Words[w2 - 1] + " ";
                                                if ((w2) >= 0) findedpwd += Words[w2] + " ";
                                                if ((w2 + 1) < Words.Length) findedpwd += Words[w2 + 1] + " ";
                                                if ((w2 + 2) < Words.Length) findedpwd += Words[w2 + 2] + " ";
                                                
                                                string resultText = findedpwd.Trim();
                                                this.Invoke((MethodInvoker)delegate
                                                {
                                                    listView1.Items.Add(new ListViewItem(new string[] { currentFile, resultText, ext }));
                                                });
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                });
            }, token);

                toolStripStatusLabel1.Text = "Выполнено";
            }
            catch (OperationCanceledException)
            {
                toolStripStatusLabel1.Text = "Поиск остановлен пользователем";
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = "Ошибка: " + ex.Message;
            }
            finally
            {
                _cts.Dispose();
                _cts = null;
                button1.Text = originalButtonText;
            }
        }

        private static IList<string> GetTablenames(DataTableCollection tables)
        {
            var tableList = new List<string>();
            foreach (var table in tables)
            {
                tableList.Add(table.ToString());
            }

            return tableList;
        }
        /// <summary>
        /// перевод порядкового номера столбца в его буквенный эквивалент.
        /// </summary>
        /// <param name="colNum"></param>
        /// <returns></returns>
        public static string ParseColNum(int colNum)
        {
            // тут конечно каждый сам по своему может контролировать (так как для офиса 2007 это не актуально)
            if (colNum > 256)
            {
                //MessageBox.Show(@"Кол-во колонок не должно быть более 256!");
                return "error";
            }

            char res = 'A';

            if (colNum < 27) return GetLetterByNum(colNum);

            while (colNum > 52)
            {
                colNum -= 26;
                res += (char)(res + 1);
            }

            colNum -= 26;
            return (res) + GetLetterByNum(colNum);
        }

        /// <summary>
        /// ф. получения буквы по номеру столбца (для Excel)
        /// </summary>
        /// <param name="colNum">номер столбца</param>
        /// <returns></returns>
        private static string GetLetterByNum(int colNum)
        {
            if (colNum <= 0) return "A";
            var book = new string[26];

            for (int i = 0; i < 26; i++)
            {
                book[i] = (((char)('A' + i)).ToString());
            }

            return book[colNum - 1];
        }

        private void RichTextBox2_TextChanged(object sender, EventArgs e)
        {
            richTextBox2.Text=richTextBox2.Text.Replace("*.doc", "");
            richTextBox2.Text = richTextBox2.Text.Replace("*.xls", "");
            richTextBox2.Text = richTextBox2.Text.Replace("*.txt", "");
            richTextBox2.Text = richTextBox2.Text.Replace("*.rtf", "");
            richTextBox2.Text = richTextBox2.Text.Replace("*.xml", "");
            richTextBox2.Text = richTextBox2.Text.Replace(",,", "");
            Settings.Default.findstring = richTextBox2.Text;
            Settings.Default.Save();
        }

        private void RusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label2.Text = "Список слов для поиска файлов (* - поиск по всем файлам):";
            label5.Text = "Список слов для поиска паролей:";
            button1.Text = "Найти пароли";
            label3.Text = "Папки где искать файлы:";
            listView1.Columns[0].Text = "Найденный файл";
            listView1.Columns[1].Text = "Найденный пароль";
            listView1.Columns[2].Text = "Тип";
        }

        private void EngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label2.Text = "Text to find files (* - analyze all files):";
            label5.Text = "Word for find passwords:";
            button1.Text = "Find passwords";
            label3.Text = "Folders where find files:";
            listView1.Columns[0].Text = "Finded file";
            listView1.Columns[1].Text = "Finded password";
            listView1.Columns[2].Text = "Type";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Icon = Resources.ico;
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.findwords = richTextBox3.Text;
            Settings.Default.Save(); Settings.Default.Reload(); 
        }

        private void docBox_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.doc = docBox.Checked;
            Settings.Default.Save(); 
        }

        private void xlsBox_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.xls = xlsBox.Checked;
            Settings.Default.Save(); Settings.Default.Upgrade(); 
        }

        private void txtBox_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.txt = txtBox.Checked;
            Settings.Default.Save();
        }

        private void rtfBox_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.rtf = rtfBox.Checked;
            Settings.Default.Save();
        }

        private void xmlBox_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.xml = xmlBox.Checked;
            Settings.Default.Save(); 
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_cts != null)
            {
                _cts.Cancel();
            }

            Settings.Default.Save();
            Settings.Default.Upgrade();
        }
    }
    public static class StringExtensions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }
    }
}
