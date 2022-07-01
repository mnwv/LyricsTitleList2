using System;
using System.Windows.Controls;
using System.Data;
using System.IO;

namespace LyricsTitleList2
{
    internal class Program
    {
        private const string LYRICS_ROOT2 = @"C:\Projects\Lyrics\Lyrics-LaTeX2.github";

        [STAThread] // 印刷処理を行う場合は STAThread 属性が必要です
        static void Main(string[] args)
        {
            // 歌詞曲名を印刷
            DataTable dt = collectLyricsTitles(LYRICS_ROOT2);
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                if (printDialog.PrintableAreaWidth < printDialog.PrintableAreaHeight)
                {
                    Console.WriteLine("用紙の向きを[横]に設定してください");
                }
                else
                {
                    LyricsTitlesPagenator pagenator = new LyricsTitlesPagenator(dt,
                        new System.Windows.Size(printDialog.PrintableAreaWidth, printDialog.PrintableAreaHeight));
                    printDialog.PrintDocument(pagenator, "曲名一覧");
                }
                Console.Write("[Enter]を押下するとプログラムを終了します。>");
                Console.ReadLine();
            }

        }

        private static DataTable collectLyricsTitles(string lyricsRooot)
        {
            Console.WriteLine($"曲名集積開始。ルート:{lyricsRooot}");
            DataTable dt = new DataTable();
            dt.TableName = "曲名一覧";
            dt.Columns.Add(new DataColumn("頭文字", typeof(string)));
            dt.Columns.Add(new DataColumn("親フォルダ", typeof(string)));
            dt.Columns.Add(new DataColumn("曲名", typeof(string)));

            foreach (string capitalPath in Directory.EnumerateDirectories(lyricsRooot))
            {
                string capital = Path.GetFileNameWithoutExtension(capitalPath);
                if (capital.Length != 1) continue;
                foreach (string kanaPath in Directory.EnumerateDirectories(capitalPath))
                {
                    string kana = Path.GetFileName(kanaPath);
                    foreach (string filePath in Directory.EnumerateFiles(kanaPath))
                    {
                        if (Path.GetExtension(filePath) != ".pdf") continue;
                        DataRow row = dt.NewRow();
                        row[dt.Columns["頭文字"]] = capital;
                        row[dt.Columns["親フォルダ"]] = kana;
                        row[dt.Columns["曲名"]] = Path.GetFileNameWithoutExtension(filePath);
                        dt.Rows.Add(row);
                    }
                }
            }
            // Create DataView
            DataView view = new DataView(dt);
            // Sort by 親フォルダ column in descending order
            view.Sort = "親フォルダ ASC";
            dt = view.ToTable();
            Console.WriteLine($"曲数:{dt.Rows.Count}");
            return dt;
        }
    }
}
