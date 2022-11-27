using System.Drawing;
using ClosedXML.Excel;

#pragma warning disable

internal static partial class Program
{
  internal static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

  internal static void MainStream(string[] args)
  {
    if (Parse(args) == false)
    {
      logger.Warn("パラメータの解析に失敗しました。");
      return;
    }

    logger.Info($"画像ファイル({target_img_file})の解析を開始します。");

    // エクセル関連

    // Excelブックを取得する。
    if (File.Exists(target_img_file) == false)
    {
      logger.Info($"対象となるExcelファイル({target_img_file})が存在しません。");
      return;
    }
    XLWorkbook book = new(target_img_file);

    // ワークシート一覧を取得する。
    var worksheets = book.Worksheets;

    logger.Info($"ワークシート({worksheets.Count}件)の解析を開始します。");
    worksheets.ToList().ForEach(a => logger.Info($"{a}"));

    foreach (var worksheet in worksheets)
    {
      logger.Info($"ワークシート({worksheet.Name})の解析を開始します。 | {target_img_file}");

      // セル一覧を取得する。
      var cells = worksheet.CellsUsed(true);

      logger.Info($"セル({cells.Count()})の解析を開始します。");
      // cells.ToList().ForEach(a => logger.Info($"{a}"));

      // 最大値を取得する。
      var max_row = cells.Max(a => a.Address.RowNumber);
      var max_column = cells.Max(a => a.Address.ColumnNumber);

      Bitmap bitmap = new(max_row, max_column);

      foreach (var cell in cells)
      {
        var address = cell.Address;

        var x = address.ColumnNumber - 1;
        var y = address.RowNumber - 1;

        // logger.Info($"セル({cell.Address})の解析を開始します。");

        // セルの背景色を取得する。
        var color = cell.Style.Fill.BackgroundColor;
        // logger.Info($"セル({cell.Address})の背景色({color})を取得しました。");

        // セルの背景色をリストに追加する。
        bitmap.SetPixel(x, y, color.Color);
      }

      var outputfilename = $"{worksheet.Name}.png";

      bitmap.Save(outputfilename);
      logger.Info($"ワークシート({worksheet.Name})の解析が完了しました。 -> {outputfilename}");
    }
  }
}

#pragma warning restore
