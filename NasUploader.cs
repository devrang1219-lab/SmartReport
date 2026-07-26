using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    // 간단한 NAS(네트워크 공유) 업로드 도우미
    // 사용: NasUploader.UploadReports(localRoot, nasRoot, username, password, logAction);
    public static class NasUploader
    {
        public static void UploadReports(string localRoot, string nasRoot, string username, string password, Action<string> log = null)
        {
            if (string.IsNullOrEmpty(localRoot) || !Directory.Exists(localRoot))
            {
                log?.Invoke($"로컬 폴더를 찾을 수 없음: {localRoot}");
                return;
            }

            if (string.IsNullOrEmpty(nasRoot))
            {
                log?.Invoke("NAS 경로를 지정하세요.");
                return;
            }

            // 먼저 NAS 공유에 연결 시도 (UNC 루트로 연결)
            bool connected = false;
            string uncRoot = nasRoot;
            try
            {
                connected = NetworkShare.ConnectToShare(uncRoot, username, password, log);

                var localDirs = Directory.GetDirectories(localRoot);
                foreach (var localDir in localDirs)
                {
                    string localName = Path.GetFileName(localDir);
                    log?.Invoke($"처리 폴더: {localName}");

                    // NAS에서 동일한 폴더 이름을 검색
                    string[] matches = new string[0];
                    try
                    {
                        if (Directory.Exists(uncRoot))
                            matches = Directory.GetDirectories(uncRoot, localName, SearchOption.AllDirectories);
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke($"NAS 경로 탐색 중 오류: {ex.Message}");
                    }

                    if (matches == null || matches.Length == 0)
                    {
                        log?.Invoke($"NAS에서 동일한 폴더를 찾지 못함: {localName} (스킵)");
                        continue;
                    }

                    string targetFolder = matches[0];
                    log?.Invoke($"대상 NAS 폴더: {targetFolder}");

                    bool isAnnual = localName.IndexOf("연차", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (isAnnual)
                    {
                        // 연차 폴더인 경우 하위 '04 보고서' 폴더의 모든 파일 업로드
                        string reportSub = Path.Combine(localDir, "04 보고서");
                        if (!Directory.Exists(reportSub))
                        {
                            log?.Invoke($"연차 폴더지만 '04 보고서'를 찾을 수 없음: {reportSub} (스킵)");
                            continue;
                        }

                        var files = Directory.GetFiles(reportSub);
                        foreach (var f in files)
                        {
                            try
                            {
                                string dest = Path.Combine(targetFolder, Path.GetFileName(f));
                                File.Copy(f, dest, true);
                                log?.Invoke($"업로드 성공: {f} -> {dest}");
                            }
                            catch (Exception ex)
                            {
                                log?.Invoke($"업로드 실패: {f} -> {targetFolder} : {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        // 연차가 아닌 경우 폴더내의 .xlsx와 .pdf 파일만 업로드
                        var files = Directory.GetFiles(localDir)
                            .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));

                        foreach (var f in files)
                        {
                            try
                            {
                                string dest = Path.Combine(targetFolder, Path.GetFileName(f));
                                File.Copy(f, dest, true);
                                log?.Invoke($"업로드 성공: {f} -> {dest}");
                            }
                            catch (Exception ex)
                            {
                                log?.Invoke($"업로드 실패: {f} -> {targetFolder} : {ex.Message}");
                            }
                        }
                    }
                }
            }
            finally
            {
                if (connected)
                {
                    try { NetworkShare.DisconnectShare(uncRoot, true, log); } catch { }
                }
            }
        }
    }

    internal static class NetworkShare
    {
        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(ref NETRESOURCE netResource, string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);

        [StructLayout(LayoutKind.Sequential)]
        private struct NETRESOURCE
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpLocalName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpRemoteName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpComment;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpProvider;
        }

        public static bool ConnectToShare(string uncPath, string username, string password, Action<string> log = null)
        {
            try
            {
                var nr = new NETRESOURCE
                {
                    dwType = 1, // RESOURCETYPE_DISK
                    lpRemoteName = uncPath
                };

                int result = WNetAddConnection2(ref nr, password, username, 0);
                if (result != 0)
                {
                    log?.Invoke($"네트워크 연결 실패: 코드 {result}");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"네트워크 연결 중 예외: {ex.Message}");
                return false;
            }
        }

        public static void DisconnectShare(string uncPath, bool force, Action<string> log = null)
        {
            try
            {
                int result = WNetCancelConnection2(uncPath, 0, force);
                if (result != 0) log?.Invoke($"네트워크 연결 해제 실패: 코드 {result}");
            }
            catch (Exception ex)
            {
                log?.Invoke($"네트워크 연결 해제 예외: {ex.Message}");
            }
        }
    }
}
