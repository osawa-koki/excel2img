using Microsoft.Extensions.Configuration;

internal static partial class Program
{
  internal static string target_img_file = "";

  internal static bool Parse(string[] args)
  {
    try
    {
      // コマンドライン引数の解析
      var builder = new ConfigurationBuilder();
      builder.AddCommandLine(args);

      var config = builder.Build();

      var tmp_target_img_file = config["t"];

      if (tmp_target_img_file == null)
      {
        logger.Warn("「/t」オプションで対象Excelファイルパスを指定して下さい。");
        return false;
      }

      // コマンドラインが正常なら、パラメータを設定する
      
      target_img_file = tmp_target_img_file;

      return true;
    }
    catch (Exception ex)
    {
      logger.Error(ex);
      return false;
    }
  }
}
