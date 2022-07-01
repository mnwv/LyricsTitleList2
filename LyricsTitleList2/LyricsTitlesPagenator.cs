using System;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows;
using System.Data;
using System.Globalization;

namespace LyricsTitleList2
{
    internal class LyricsTitlesPagenator : DocumentPaginator
    {
        public override bool IsPageCountValid => true;

        public override int PageCount => 2;

        public override System.Windows.Size PageSize
        {
            get => _pageSize;
            set
            {
                _pageSize = value;
                PagenateData();
            }
        }

        public override IDocumentPaginatorSource Source => null;

        public override DocumentPage GetPage(int pageNumber)
        {
            Console.WriteLine($"GetPage() pageNumber:{pageNumber}");
            DrawingVisual visual = new DrawingVisual();

            using (DrawingContext dc = visual.RenderOpen())
            {
                DrawTitle(dc, pageNumber + 1);  // pageNumberは0オリジン
                //DrawOuterRectangle(dc);
                DrawItems(dc, pageNumber);
            }
            return new DocumentPage(visual, PageSize, new Rect(_pageSize), new Rect(_pageSize));
        }

        private DataTable _dt;

        private readonly Typeface _typeface = new Typeface("Meiryo UI");
        private readonly Typeface _numTypeface = new Typeface("Times New Roman");
        private Size _pageSize;
        private double _marginY;

        private FormattedText _titleText;
        private double _columnWidth;
        private double _capitalWidth;
        private double _gridTop;
        private double _gridBottom;

        private const int ROWS_PER_PAGE = 50;
        private const int COLS_PER_PAGE = 7;
        private const double MARGIN_X = 10.0;
        private const double TITLE_FONT_SIZE = 14.0;
        private const double DATE_FONT_SIZE = 10.0;
        private const double PAGE_NUM_FONT_SIZE = 10.0;
        private const double LYRICS_FONT_SIZE = 12.0;
        private const double THIN = 0.3;
        private const double THICK = 1.5;
        private const double CAPITAL_MARGIN_L = 2.0;
        private const double CAPITAL_MARGIN_R = 1.0;
        private const double NAME_MARGIN_L = 1.0;
        
        internal LyricsTitlesPagenator(DataTable dt, Size pageSize)
        {
            Console.WriteLine($"ページ幅:{pageSize.Width} ページ高:{pageSize.Height}");
            _dt = dt;
            _pageSize = pageSize;
            PagenateData();
        }

        /// <summary>
        /// 頁の諸要素を計算
        /// </summary>
        private void PagenateData()
        {
            // 縦のマージンを計算
            _titleText = GetFormattedText("曲名一覧",_typeface, TITLE_FONT_SIZE);
            // 頭文字欄の幅を計算
            FormattedText itemText = GetFormattedText("あ");
            _capitalWidth = itemText.Width + CAPITAL_MARGIN_L + CAPITAL_MARGIN_R;
            // 頁の上下のマージンを計算
            double gridHeight = itemText.Height * ROWS_PER_PAGE;
            _marginY = (_pageSize.Height - gridHeight - _titleText.Height - 1.0) / 2;
            Console.WriteLine($"_margin_Y:{_marginY}");
            // カラム幅(頭文字欄の幅を含む)を計算
            _columnWidth = (_pageSize.Width - MARGIN_X * 2) / COLS_PER_PAGE;
            Console.WriteLine($"_column_Width:{_columnWidth}");
            // 表の上端位置を計算
            _gridTop = _marginY + _titleText.Height + 1.0;
            Console.WriteLine($"_gridTop:{_gridTop}");
            // 表の下端位置を計算
            _gridBottom = _gridTop + gridHeight;
            Console.WriteLine($"_gridBottom:{_gridBottom}");
        }

        /// <summary>
        /// タイトルを描画
        /// </summary>
        /// <param name="dc"></param>
        /// <param name="pageNum">ページ番号(0オリジン)</param>
        private void DrawTitle(DrawingContext dc, int pageNum)
        {
            double titleX = (_pageSize.Width - _titleText.Width) / 2;
            dc.DrawText(_titleText, new Point(titleX, _marginY));
            FormattedText date =
                GetFormattedText(DateTime.Today.ToString("yyyy/MM/dd"), _numTypeface, DATE_FONT_SIZE);
            double date_Y = _marginY + (_titleText.Height - date.Height);
            dc.DrawText(date, new Point(MARGIN_X, date_Y));
            FormattedText page =
                GetFormattedText($"Psge: {pageNum}/{PageCount}", _numTypeface, PAGE_NUM_FONT_SIZE);
            double pagenumY = _marginY + (_titleText.Height - page.Height);
            double pagenumX = _pageSize.Width - MARGIN_X - page.Width;
            dc.DrawText(page, new Point(pagenumX, pagenumY));
        }

        private void DrawItems(DrawingContext dc, int pageNumber)
        {
            Point point = new Point(MARGIN_X, _gridTop);

            int start = ROWS_PER_PAGE * COLS_PER_PAGE * pageNumber;
            for (int i = 0; i < ROWS_PER_PAGE * COLS_PER_PAGE; i++)
            {
                double thickness = THIN;
                int rowIdx = start + i;
                if (rowIdx >= _dt.Rows.Count - 1)
                {   // データの最後
                    DrawHLine(dc, point, THICK);
                    EndOfDataDrawLine(dc, point.X);
                    break;
                }
                DataRow row = _dt.Rows[rowIdx];
                if (i % ROWS_PER_PAGE == 0)
                {   // 先頭行 縦線描画 頭文字描画
                    point.X = MARGIN_X + _columnWidth * (int)(i / ROWS_PER_PAGE);
                    point.Y = _gridTop;
                    DrawVLines(dc, point.X);
                    DrawCapital(dc, row, point);
                    thickness = THICK;
                }
                else if (row[_dt.Columns["頭文字"]].ToString() !=
                    _dt.Rows[rowIdx - 1][_dt.Columns["頭文字"]].ToString())
                {
                    DrawCapital(dc, row, point);
                    thickness = THICK;
                }
                DrawHLine(dc, point, thickness);
                FormattedText name = GetFormattedText(row[_dt.Columns["曲名"]].ToString());
                dc.DrawText(name, new Point(point.X + _capitalWidth + NAME_MARGIN_L, point.Y));
                point.Y += name.Height;
                if ((i + 1) % ROWS_PER_PAGE == 0)
                {   // 行の最後
                    DrawHLine(dc, point, THICK);
                    if (i + 1 == ROWS_PER_PAGE * COLS_PER_PAGE)
                    {
                        DrawVLines(dc, point.X + _columnWidth);
                    }
                }
            }
        }

        private void DrawHLine(DrawingContext dc, Point start, double thickness)
        {
            Point end = new Point(start.X + _columnWidth, start.Y);
            if (thickness == THIN)
            {
                start.X += _capitalWidth;
            }
            dc.DrawLine(new Pen(Brushes.Black, thickness), start, end);
        }

        private void DrawVLines(DrawingContext dc, double x)
        {
            Point pt1 = new Point(x, _gridTop);
            Point pt2 = new Point(x, _gridBottom);
            dc.DrawLine(new Pen(Brushes.Black, THICK), pt1, pt2);
            if (x < _pageSize.Width - MARGIN_X * 2)
            {   // 頭文字右側の縦線
                pt1.X += _capitalWidth;
                pt2.X += _capitalWidth;
                dc.DrawLine(new Pen(Brushes.Black, THIN), pt1, pt2);
            }
        }

        private void EndOfDataDrawLine(DrawingContext dc, double x)
        {
            // 水平線
            Point pt1 = new Point(x, _gridBottom);
            Point pt2 = new Point(x + _columnWidth, _gridBottom);
            dc.DrawLine(new Pen(Brushes.Black, THICK), pt1, pt2);
            // 垂直線
            pt1.X = x + _columnWidth;
            pt1.Y = _gridTop;
            pt2.X = pt1.X;
            pt2.Y = _gridBottom;
            dc.DrawLine(new Pen(Brushes.Black, THICK), pt1, pt2);
        }

        private void DrawCapital(DrawingContext dc, DataRow row, Point point)
        {
            FormattedText capital = GetFormattedText(row[_dt.Columns["頭文字"]].ToString());
            point.X += CAPITAL_MARGIN_L;
            dc.DrawText(capital, point);
        }

        private FormattedText GetFormattedText(string text)
        {
            return GetFormattedText(text, _typeface, LYRICS_FONT_SIZE);
        }

        private FormattedText GetFormattedText(string text, Typeface typeface, double fontSize)
        {
            DrawingVisual visual = new DrawingVisual();
            return new FormattedText(
                text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                typeface, fontSize, Brushes.Black,
                VisualTreeHelper.GetDpi(visual).PixelsPerDip);
        }

        //private void DrawOuterRectangle(DrawingContext dc)
        //{
        //    dc.DrawRectangle(null, new Pen(Brushes.Black, THICK),
        //        new Rect(MARGIN_X, _gridTop,
        //                 _pageSize.Width - MARGIN_X - 2.0,
        //                 _pageSize.Height - _marginY * 2 - _titleText.Height - 1.0));
        //}
    }
}
