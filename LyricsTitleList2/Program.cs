using System;
using System.Windows.Controls;
using System.Data;
using System.IO;
using System.Windows;

namespace LyricsTitleList2
{
    internal class Program
    {
        [STAThread] // 印刷処理を行う場合は STAThread 属性が必要です
        static void Main(string[] args)
        {
            var musicFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
            string lyricsRoot = System.IO.Path.Combine(musicFolder, "Lyrics-LaTeX2.github");
            if (!Directory.Exists(lyricsRoot))
            {
                MessageBox.Show($"歌詞ファイルフォルダ{lyricsRoot}が見つかりません。");
                Application.Current.Shutdown();
                return;
            }

            // 歌詞曲名を印刷
            DataTable dt = collectLyricsTitles(lyricsRoot);
            PrintDialog printDialog = new PrintDialog();

            // 用紙を、A4, Landscapeにセット
            printDialog.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape;
            printDialog.PrintTicket.PageMediaSize = new System.Printing.PageMediaSize(System.Printing.PageMediaSizeName.ISOA4);

            if (printDialog.ShowDialog() == true)
            {
                if (printDialog.PrintableAreaWidth < printDialog.PrintableAreaHeight)
                {
                    Console.WriteLine("用紙の向きを[横]に設定してください");
                }
                else
                {
                    var printQueue = printDialog.PrintQueue;
                    Console.WriteLine($"printQueue.Name:{printQueue.Name}");
                    var printCapabilities = printQueue.GetPrintCapabilities();
                    Console.WriteLine($"ExtentWidth:{printCapabilities.PageImageableArea.ExtentWidth}");
                    Console.WriteLine($"ExtentHeight:{printCapabilities.PageImageableArea.ExtentHeight}");
                    Console.WriteLine($"OriginWidth:{printCapabilities.PageImageableArea.OriginWidth}");
                    Console.WriteLine($"OriginHeight:{printCapabilities.PageImageableArea.OriginHeight}");
                    Console.WriteLine($"OrientedPageMediaWidth:{printCapabilities.OrientedPageMediaWidth}");
                    Console.WriteLine($"OrientedPageMediaHeight{printCapabilities.OrientedPageMediaHeight}");

                    LyricsTitlesPagenator pagenator = new LyricsTitlesPagenator(dt, 
                        new System.Windows.Size(printDialog.PrintableAreaWidth, printDialog.PrintableAreaHeight),
                        printCapabilities);

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

/*
 * 
printQueue.Name:Brother DCP-J926N Printer (1 コピー)
ExtentWidth:771.299527559055
ExtentHeight:1100.11842519685
OriginWidth:11.1987401574803
OriginHeight:11.1987401574803
OrientedPageMediaWidth:793.700787401575
OrientedPageMediaHeight1122.51968503937
_itemFontSize:11.6
_margin_Y:19.3803937007873
_column_Width:157.159775028121
_gridTop:37.6537270341207
_gridBottom:774.320393700787
*/