using NLog;

string mutexName = "excel2img";

Mutex mutex = new(false, mutexName);
bool hasHandle = false;

#if !DEBUG
try
{
#endif
try
{
  // ミューテックスの所有権を要求する
  hasHandle = mutex.WaitOne(0, false);
}

catch (AbandonedMutexException)
{
  // 別のアプリケーションがミューテックスを解放しないで終了した時
  hasHandle = true;
}
// ミューテックスを得られたか調べる
if (hasHandle == false)
{
  // 得られなかった場合は、すでに起動していると判断して終了
  logger.Warn("既に別のインスタンスが実行されています。");
  return 1;
}

logger.Info("実行開始");
MainStream(args);
logger.Info("実行完了");
return 0;
#if !DEBUG
}
catch (Exception ex)
{
  logger.Error(ex);
  return 1;
}
finally
{
  if (hasHandle)
  {
    // ミューテックスを解放する
    mutex.ReleaseMutex();
  }
  mutex.Close();
}
#endif
