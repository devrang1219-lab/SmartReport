using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1.Comm
{
    public class SoborLog
    {
        public string PathBase = "C:/Soborsoft" + "/log/"; // 로그 폴더
        public int Interval = 2000; // 로그 쓰기 간격
        public int Period = 365; // 로그 보관 일자
        public bool DoSave = true;
        public bool DoDeleteOldLog = true;
        public bool UseLineHeader = true;

        public string Tag = "";

        private StreamWriter sw;
        private Mutex mtxLog = new Mutex(false, "LogMutex");
        private Thread thLog;
        private ConcurrentQueue<string> logList = new ConcurrentQueue<string>();
        private CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();

        private ListBox allLog = null;
        private ListBox errorLog = null;
        private ListBox warningLog = null;
        private ListBox infoLog = null;

        private string LogPath
        {
            get
            {
                return $"{PathBase}{DateTime.Now.ToString("yyyy")}/{DateTime.Now.ToString("MM")}/";
            }
        }

        private string LogFile
        {
            get
            {
                return $"{LogPath}{((Tag != "") ? Tag + "_" : "")}{LogDate}.log";
            }
        }

        public string LogDate
        {
            get
            {
                return DateTime.Now.ToString("yyyyMMdd");
            }
        }

        public string preDate = DateTime.Now.ToString("yyyyMMdd");

        public string LineHeader
        {
            get
            {
                return $"[{DateTime.Now.ToString("HH:mm:ss")}] ";
            }
        }


        public SoborLog()
        {
            //Start();
            StartLoop();
        }

        public SoborLog(ListBox allLog = null, ListBox infoLog = null, ListBox warningLog = null, ListBox errorLog = null)
        {
            this.allLog = allLog;
            this.infoLog = infoLog;
            this.warningLog = warningLog;
            this.errorLog = errorLog;
            //Start();
            StartLoop();
        }

        public SoborLog(string path, string tag = "")
        {
            path = path + "/";
            path = path.Replace("//", "/");
            path = path.Replace("./", "C:/Soborsoft/log/");
            path = path.Replace("/log/log/", "/log/");
            PathBase = path;
            Tag = tag;
            //Start();
            StartLoop();
        }

        // 객체 생성 (저장 폴더)
        public void Start()
        {
            thLog = new Thread(LoopLog);
            thLog.IsBackground = true;
            thLog.Start();
        }

        public void StartLoop()
        {
            if (_cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
            }

            _cancellationTokenSource = new CancellationTokenSource();
            _ = LoopLogAsync(); // 비동기 시작
        }

        public void StopLoop()
        {
            _cancellationTokenSource.Cancel(); // 루프 종료
        }

        // 자원 해제
        public void Dispose()
        {
            try
            {
                if (thLog != null)
                {
                    thLog.Join();
                    thLog = null;
                }

                if (mtxLog != null)
                {
                    mtxLog.Dispose();
                }
                
                if (sw != null) sw.Dispose();
            }
            catch
            { }
        }


        // 로그 쓰기 무한루프 (생성시 쓰레드로 실행)
        private void LoopLog()
        {
            while (true)
            {
                if (DoSave)
                {
                    WriteLogList();
                }
                if (DoDeleteOldLog)
                {
                    DeleteOldLog();
                }

                Task.Delay(Interval);
            }
        }

        private async Task LoopLogAsync()
        {
            var token = _cancellationTokenSource.Token;

            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (DoSave)
                    {
                        WriteLogList();
                    }

                    if (DoDeleteOldLog)
                    {
                        DeleteOldLog();
                    }

                    await Task.Delay(Interval, token); // CPU 점유율 최소화
                }
            }
            catch (TaskCanceledException)
            {
                // 루프가 종료된 경우 처리
                Console.WriteLine("Log loop was canceled.");
            }
            catch (Exception ex)
            {
                // 예외 처리
                Console.WriteLine($"Unexpected error in log loop: {ex.Message}");
            }
        }


        // 로그 목록에 메시지를 추가
        public void Add(string str, bool useHeader = true)
        {
            logList.Enqueue(((useHeader) ? LineHeader : "") + str);
        }


        // 로그 목록 파일에 쓰기
        private void WriteLogList()
        {

            List<string> logs = GetLogList();
            if (logs == null || logs.Count == 0)
            {
                return;
            }

            try
            {
                // 폴더가 없으면 폴더를 만든다.
                if (!Directory.Exists(LogPath))
                {
                    Directory.CreateDirectory(LogPath);
                }

                // 문자열을 쓴다. (파일이 없으면 파일을 만든다.)
                sw = new StreamWriter(LogFile, true, Encoding.Default);

                foreach (string log in logs)
                {
                    if (log == "")
                    {
                        continue;
                    }
                    sw.WriteLine(log);
                    WriteLogToListView(log);
#if DEBUG
                    System.Diagnostics.Debug.WriteLine(log);
#endif
                }
                sw.Close();
            }
            catch
            {
                //throw ex;
            }
        }

        private void WriteLogToListView(string log)
        {
            if (allLog == null) return; 
            // 추가, 최대 200줄 유지, 마지막 줄로 스크롤
            Action addAndTrim = () =>
            {
                allLog.Items.Add(log);

                // 200줄 초과하면 오래된 항목부터 제거
                while (allLog.Items.Count > 200)
                {
                    try
                    {
                        allLog.Items.RemoveAt(0);
                    }
                    catch { break; }
                }

                // 마지막 항목이 보이도록 스크롤
                if (allLog.Items.Count > 0)
                {
                    try
                    {
                        allLog.TopIndex = allLog.Items.Count - 1;
                    }
                    catch { }
                }
            };

            if (allLog.InvokeRequired)
            {
                allLog.BeginInvoke(addAndTrim); // 비동기(블로킹 안함)
            }
            else
            {
                addAndTrim();
            }
        }


        // 로그 목록 반환
        public List<string> GetLogList()
        {
            if (logList.Count <= 0)
            {
                return null;
            }

            List<string> logs = new List<string>();
            while (logList.TryDequeue(out string log))
            {
                logs.Add(log);
            }
            return logs;
        }


        // 오래된 로그 파일 삭제
        private void DeleteOldLog()
        {
            // 지난 로그 삭제
            if (LogDate != preDate)
            {
                for (int i = -7; i > -14; i--)
                {
                    DateTime preDate = DateTime.Now.AddDays(i);
                    string filePath = LogPath + preDate.ToString("yyyyMMdd") + ".log";

                    try
                    {
                        File.Delete(filePath);
                        filePath = LogPath + Tag + "_" + preDate.ToString("yyyyMMdd") + ".log";
                        File.Delete(filePath);
                    }
                    catch
                    { }
                }
            }
        }
    }
}
