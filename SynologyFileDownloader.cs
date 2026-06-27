using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace SynologyIntegration
{
    public class SynologyFileDownloaderConfig
    {
        public string BaseUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }

        // 탐색 시작 폴더 (예: /2_1전기직무고시점검보고서)
        public string SearchFolder { get; set; }

        // 파일명에 포함될 검색어
        public string Keyword { get; set; }

        // 허용 확장자 (예: .xlsx, .pdf)
        public IEnumerable<string> Extensions { get; set; }

        // 하위 폴더까지 재귀 탐색 여부
        public bool Recursive { get; set; }

        // 검색 결과 중 최신 파일만 다운로드할 때 사용
        public bool DownloadLatestOnly { get; set; }

        public SynologyFileDownloaderConfig()
        {
            BaseUrl = "";
            Username = "";
            Password = "";
            SearchFolder = "/";
            Keyword = "";
            Extensions = new string[0];
            Recursive = true;
            DownloadLatestOnly = false;
        }

        public string[] GetNormalizedExtensions()
        {
            if (Extensions == null)
                return new string[0];

            return Extensions
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x =>
                {
                    string ext = x.Trim();
                    if (!ext.StartsWith(".")) ext = "." + ext;
                    return ext.ToLowerInvariant();
                })
                .Distinct()
                .ToArray();
        }

        
    }

    

    //public class SynologyFileItem
    //{
    //    public string Name { get; set; }
    //    public string Path { get; set; }
    //    public bool IsDir { get; set; }
    //    public long Size { get; set; }
    //    public DateTime ModifiedTime { get; set; }

        //    public SynologyFileItem()
        //    {
        //        Name = "";
        //        Path = "";
        //        ModifiedTime = DateTime.MinValue;
        //    }

        //    public override string ToString()
        //    {
        //        return string.Format("{0} ({1:N0} bytes, {2:yyyy-MM-dd HH:mm:ss})", Path, Size, ModifiedTime);
        //    }
        //}

    public class SynologyFileItem
    {
        public bool Selected { get; set; }   // 체크박스용

        public string Name { get; set; }
        public string Path { get; set; }
        public long Size { get; set; }
        public DateTime ModifiedTime { get; set; }
        public bool IsDir { get; set; }

        public SynologyFileItem()
        {
            Name = "";
            Path = "";
            ModifiedTime = DateTime.MinValue;
        }
    }

    //public class SynologyFileDownloader : IDisposable
    //{
    //    private readonly SynologyFileDownloaderConfig _config;
    //    private readonly HttpClient _httpClient;
    //    private readonly CookieContainer _cookieContainer;
    //    private readonly JavaScriptSerializer _serializer;

    //    private string _sid;
    //    private bool _disposed;

    //    public SynologyFileDownloader(SynologyFileDownloaderConfig config)
    //    {
    //        if (config == null) throw new ArgumentNullException("config");

    //        _config = config;
    //        _cookieContainer = new CookieContainer();
    //        _serializer = new JavaScriptSerializer();

    //        var handler = new HttpClientHandler
    //        {
    //            CookieContainer = _cookieContainer,
    //            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
    //        };

    //        _httpClient = new HttpClient(handler);
    //        _httpClient.Timeout = TimeSpan.FromMinutes(10);
    //    }

    //    public bool IsLoggedIn
    //    {
    //        get { return !string.IsNullOrWhiteSpace(_sid); }
    //    }

    //    /// <summary>
    //    /// DSM 로그인
    //    /// </summary>
    //    public async Task LoginAsync(CancellationToken cancellationToken = default(CancellationToken))
    //    {
    //        ThrowIfDisposed();

    //        string url =
    //            _config.BaseUrl + "/webapi/auth.cgi" +
    //            "?api=SYNO.API.Auth" +
    //            "&version=6" +
    //            "&method=login" +
    //            "&account=" + Uri.EscapeDataString(_config.Username) +
    //            "&passwd=" + Uri.EscapeDataString(_config.Password) +
    //            "&session=FileStation" +
    //            "&format=sid";

    //        string json = await GetStringAsync(url, cancellationToken).ConfigureAwait(false);

    //        // 디버깅용: 로그인 응답 원문 확인
    //        // throw new Exception("로그인 응답 원문:\r\n" + json);

    //        var response = Deserialize<SynoLoginResponse>(json);

    //        if (response == null || !response.success || response.data == null || string.IsNullOrWhiteSpace(response.data.sid))
    //        {
    //            throw new InvalidOperationException("Synology 로그인 실패. 응답: " + json);
    //        }

    //        // ★ 이 줄이 반드시 있어야 함
    //        _sid = response.data.sid;
    //    }

    //    /// <summary>
    //    /// 설정값 기준으로 파일 검색
    //    /// </summary>
    //    public async Task<List<SynologyFileItem>> SearchFilesAsync(CancellationToken cancellationToken = default(CancellationToken))
    //    {
    //        return await SearchFilesAsync(
    //            _config.SearchFolder,
    //            _config.Keyword,
    //            _config.Extensions,
    //            _config.Recursive,
    //            cancellationToken).ConfigureAwait(false);
    //    }

    //    /// <summary>
    //    /// 지정 폴더/키워드/확장자로 파일 검색
    //    /// </summary>
    //    public async Task<List<SynologyFileItem>> SearchFilesAsync(
    //        string folderPath,
    //        string keyword,
    //        IEnumerable<string> allowedExtensions,
    //        bool recursive,
    //        CancellationToken cancellationToken = default(CancellationToken))
    //    {
    //        ThrowIfDisposed();
    //        EnsureLoggedIn();

    //        string taskId = await StartSearchAsync(folderPath, keyword, recursive, cancellationToken).ConfigureAwait(false);

    //        try
    //        {
    //            var results = await ListSearchResultsAsync(taskId, cancellationToken).ConfigureAwait(false);

    //            HashSet<string> extSet = null;

    //            if (allowedExtensions != null)
    //            {
    //                var extArray = allowedExtensions
    //                    .Where(x => !string.IsNullOrWhiteSpace(x))
    //                    .Select(NormalizeExtension)
    //                    .Distinct(StringComparer.OrdinalIgnoreCase)
    //                    .ToArray();

    //                if (extArray.Length > 0)
    //                {
    //                    extSet = new HashSet<string>(extArray, StringComparer.OrdinalIgnoreCase);
    //                }
    //            }

    //            var files = results
    //                .Where(x => !x.IsDir)
    //                .Where(x =>
    //                {
    //                    if (extSet == null || extSet.Count == 0)
    //                        return true;

    //                    string ext = Path.GetExtension(x.Name) ?? "";
    //                    return extSet.Contains(ext);
    //                })
    //                .ToList();

    //            return files;
    //        }
    //        finally
    //        {
    //            await CleanSearchAsync(taskId, cancellationToken).ConfigureAwait(false);
    //        }
    //    }

    //    /// <summary>
    //    /// 설정 기준으로 검색 후 다운로드
    //    /// </summary>
    //    public async Task<List<string>> DownloadByConfigAsync(
    //        string localFolder,
    //        CancellationToken cancellationToken = default(CancellationToken))
    //    {
    //        var files = await SearchFilesAsync(cancellationToken).ConfigureAwait(false);

    //        if (_config.DownloadLatestOnly && files.Count > 0)
    //        {
    //            files = files
    //                .OrderByDescending(x => x.ModifiedTime)
    //                .Take(1)
    //                .ToList();
    //        }

    //        return await DownloadFilesAsync(files, localFolder, cancellationToken).ConfigureAwait(false);
    //    }

    //    /// <summary>
    //    /// 파일 목록 다운로드
    //    /// 반환값: 저장된 로컬 파일 경로 목록
    //    /// </summary>
    //    public async Task<List<string>> DownloadFilesAsync(
    //        IEnumerable<SynologyFileItem> files,
    //        string localFolder,
    //        CancellationToken cancellationToken = default(CancellationToken))
    //    {
    //        ThrowIfDisposed();
    //        EnsureLoggedIn();

    //        if (files == null) throw new ArgumentNullException("files");
    //        if (string.IsNullOrWhiteSpace(localFolder)) throw new ArgumentNullException("localFolder");

    //        Directory.CreateDirectory(localFolder);

    //        var savedPaths = new List<string>();

    //        foreach (var file in files)
    //        {
    //            string saved = await DownloadFileAsync(file.Path, localFolder, cancellationToken).ConfigureAwait(false);
    //            savedPaths.Add(saved);
    //        }

    //        return savedPaths;
    //    }

    //    /// <summary>
    //    /// 원격 파일 1개 다운로드
    //    /// 반환값: 저장된 로컬 파일 경로
    //    /// </summary>
    //    public async Task<string> DownloadFileAsync(
    //        string remoteFilePath,
    //        string localFolder,
    //        CancellationToken cancellationToken = default(CancellationToken))
    //    {
    //        ThrowIfDisposed();
    //        EnsureLoggedIn();

    //        if (string.IsNullOrWhiteSpace(remoteFilePath))
    //            throw new ArgumentNullException("remoteFilePath");

    //        if (string.IsNullOrWhiteSpace(localFolder))
    //            throw new ArgumentNullException("localFolder");

    //        Directory.CreateDirectory(localFolder);

    //        string fileName = Path.GetFileName(remoteFilePath);
    //        fileName = GetSafeFileName(fileName);

    //        string localPath = Path.Combine(localFolder, fileName);
    //        localPath = GetUniqueLocalPath(localPath);

    //        string url =
    //            _config.BaseUrl + "/webapi/entry.cgi" +
    //            "?api=SYNO.FileStation.Download" +
    //            "&version=2" +
    //            "&method=download" +
    //            "&mode=download" +
    //            "&_sid=" + Uri.EscapeDataString(_sid) +
    //            "&path=" + Uri.EscapeDataString(remoteFilePath);

    //        using (var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
    //        {
    //            response.EnsureSuccessStatusCode();

    //            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
    //            using (var fs = new FileStream(localPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
    //            {
    //                await stream.CopyToAsync(fs).ConfigureAwait(false);
    //            }
    //        }

    //        return localPath;
    //    }

    //    /// <summary>
    //    /// DSM 로그아웃
    //    /// </summary>
    //    public async Task LogoutAsync(CancellationToken cancellationToken = default(CancellationToken))
    //    {
    //        ThrowIfDisposed();

    //        if (string.IsNullOrWhiteSpace(_sid))
    //            return;

    //        string url =
    //            _config.BaseUrl + "/webapi/auth.cgi" +
    //            "?api=SYNO.API.Auth" +
    //            "&version=6" +
    //            "&method=logout" +
    //            "&session=FileStation" +
    //            "&_sid=" + Uri.EscapeDataString(_sid);

    //        try
    //        {
    //            await GetStringAsync(url, cancellationToken).ConfigureAwait(false);
    //        }
    //        catch
    //        {
    //            // 로그아웃 실패 무시
    //        }
    //        finally
    //        {
    //            _sid = null;
    //        }
    //    }

    //    #region Search internals

    //    private async Task<string> StartSearchAsync(
    //        string folderPath,
    //        string keyword,
    //        bool recursive,
    //        CancellationToken cancellationToken)
    //    {
    //        var form = new Dictionary<string, string>
    //        {
    //            { "api", "SYNO.FileStation.Search" },
    //            { "version", "2" },
    //            { "method", "start" },
    //            { "folder_path", folderPath },
    //            { "pattern", keyword },
    //            { "recursive", recursive ? "true" : "false" },
    //            { "additional", "[\"time\",\"size\"]" },
    //            { "_sid", _sid }
    //        };

    //        string json = await PostFormAsync(_config.BaseUrl + "/webapi/entry.cgi", form, cancellationToken).ConfigureAwait(false);
    //        var response = Deserialize<SynoSearchStartResponse>(json);

    //        if (response == null || !response.success || response.data == null || string.IsNullOrWhiteSpace(response.data.taskid))
    //        {
    //            throw new InvalidOperationException("Synology 검색 시작 실패. 응답: " + json);
    //        }

    //        return response.data.taskid;
    //    }

    //    private async Task<List<SynologyFileItem>> ListSearchResultsAsync(
    //string taskId,
    //CancellationToken cancellationToken)
    //    {
    //        // 1) 먼저 검색이 끝날 때까지 기다림
    //        bool finished = false;

    //        for (int i = 0; i < 60; i++)   // 최대 60초 대기
    //        {
    //            var form = new Dictionary<string, string>
    //    {
    //        { "api", "SYNO.FileStation.Search" },
    //        { "version", "2" },
    //        { "method", "list" },
    //        { "taskid", taskId },
    //        { "offset", "0" },
    //        { "limit", "100" },
    //        { "_sid", _sid }
    //    };

    //            string json = await PostFormAsync(_config.BaseUrl + "/webapi/entry.cgi", form, cancellationToken).ConfigureAwait(false);
    //            var response = Deserialize<SynoSearchListResponse>(json);

    //            if (response == null)
    //                throw new InvalidOperationException("검색 결과 응답 파싱 실패");

    //            if (!response.success)
    //                throw new InvalidOperationException("검색 결과 조회 실패. 응답: " + json);

    //            if (response.data != null && response.data.finished)
    //            {
    //                finished = true;
    //                break;
    //            }

    //            await Task.Delay(1000, cancellationToken).ConfigureAwait(false);
    //        }

    //        if (!finished)
    //            throw new TimeoutException("검색이 완료되지 않았습니다.");

    //        // 2) 검색 완료 후 최종 결과를 페이지 단위로 모두 가져오기
    //        var result = new List<SynologyFileItem>();
    //        int offset = 0;
    //        const int pageSize = 500;

    //        while (true)
    //        {
    //            var form = new Dictionary<string, string>
    //    {
    //        { "api", "SYNO.FileStation.Search" },
    //        { "version", "2" },
    //        { "method", "list" },
    //        { "taskid", taskId },
    //        { "offset", offset.ToString() },
    //        { "limit", pageSize.ToString() },
    //        { "_sid", _sid }
    //    };

    //            string json = await PostFormAsync(_config.BaseUrl + "/webapi/entry.cgi", form, cancellationToken).ConfigureAwait(false);
    //            var response = Deserialize<SynoSearchListResponse>(json);

    //            if (response == null)
    //                throw new InvalidOperationException("검색 결과 응답 파싱 실패");

    //            if (!response.success)
    //                throw new InvalidOperationException("검색 결과 조회 실패. 응답: " + json);

    //            if (response.data == null || response.data.files == null || response.data.files.Count == 0)
    //                break;

    //            foreach (var file in response.data.files)
    //            {
    //                result.Add(ToSynologyFileItem(file));
    //            }

    //            if (response.data.files.Count < pageSize)
    //                break;

    //            offset += pageSize;
    //        }

    //        return result;
    //    }

    //    private async Task CleanSearchAsync(string taskId, CancellationToken cancellationToken)
    //    {
    //        if (string.IsNullOrWhiteSpace(_sid))
    //            return;

    //        try
    //        {
    //            var form = new Dictionary<string, string>
    //            {
    //                { "api", "SYNO.FileStation.Search" },
    //                { "version", "2" },
    //                { "method", "clean" },
    //                { "taskid", taskId },
    //                { "_sid", _sid }
    //            };

    //            await PostFormAsync(_config.BaseUrl + "/webapi/entry.cgi", form, cancellationToken).ConfigureAwait(false);
    //        }
    //        catch
    //        {
    //            // 정리 실패 무시
    //        }
    //    }

    //    #endregion

    //    #region HTTP helpers

    //    private async Task<string> GetStringAsync(string url, CancellationToken cancellationToken)
    //    {
    //        using (var response = await _httpClient.GetAsync(url, cancellationToken).ConfigureAwait(false))
    //        {
    //            response.EnsureSuccessStatusCode();

    //            byte[] bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
    //            return Encoding.UTF8.GetString(bytes);
    //        }
    //    }

    //    private async Task<string> PostFormAsync(
    //    string url,
    //    Dictionary<string, string> form,
    //    CancellationToken cancellationToken)
    //    {
    //        using (var content = new FormUrlEncodedContent(form))
    //        using (var response = await _httpClient.PostAsync(url, content, cancellationToken).ConfigureAwait(false))
    //        {
    //            response.EnsureSuccessStatusCode();

    //            byte[] bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
    //            return Encoding.UTF8.GetString(bytes);
    //        }
    //    }

    //    #endregion

    //    #region Utils

    //    private T Deserialize<T>(string json) where T : class
    //    {
    //        return _serializer.Deserialize<T>(json);
    //    }

    //    private void EnsureLoggedIn()
    //    {
    //        if (string.IsNullOrWhiteSpace(_sid))
    //            throw new InvalidOperationException("Synology DSM 로그인 후 사용해야 합니다.");
    //    }

    //    private void ThrowIfDisposed()
    //    {
    //        if (_disposed)
    //            throw new ObjectDisposedException("SynologyFileDownloader");
    //    }

    //    private static string NormalizeExtension(string ext)
    //    {
    //        if (string.IsNullOrWhiteSpace(ext))
    //            return ext;

    //        ext = ext.Trim();
    //        return ext.StartsWith(".") ? ext : "." + ext;
    //    }

    //    private static SynologyFileItem ToSynologyFileItem(SynoSearchFile file)
    //    {
    //        DateTime modified = DateTime.MinValue;

    //        if (file.additional != null &&
    //            file.additional.time != null &&
    //            file.additional.time.mtime.HasValue)
    //        {
    //            try
    //            {
    //                modified = FromUnixTimeSeconds(file.additional.time.mtime.Value).ToLocalTime();
    //            }
    //            catch
    //            {
    //                modified = DateTime.MinValue;
    //            }
    //        }

    //        long size = 0;
    //        if (file.additional != null)
    //        {
    //            size = file.additional.size;
    //        }

    //        return new SynologyFileItem
    //        {
    //            Name = file.name ?? "",
    //            Path = file.path ?? "",
    //            IsDir = file.isdir,
    //            Size = size,
    //            ModifiedTime = modified
    //        };
    //    }

    //    private static DateTime FromUnixTimeSeconds(long seconds)
    //    {
    //        var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    //        return epoch.AddSeconds(seconds);
    //    }

    //    private static string GetSafeFileName(string fileName)
    //    {
    //        foreach (char c in Path.GetInvalidFileNameChars())
    //        {
    //            fileName = fileName.Replace(c, '_');
    //        }
    //        return fileName;
    //    }

    //    private static string GetUniqueLocalPath(string path)
    //    {
    //        if (!File.Exists(path))
    //            return path;

    //        string dir = Path.GetDirectoryName(path);
    //        string name = Path.GetFileNameWithoutExtension(path);
    //        string ext = Path.GetExtension(path);

    //        int i = 1;
    //        while (true)
    //        {
    //            string candidate = Path.Combine(dir, name + "_" + i + ext);
    //            if (!File.Exists(candidate))
    //                return candidate;
    //            i++;
    //        }
    //    }

    //    #endregion

    //    public void Dispose()
    //    {
    //        if (_disposed) return;

    //        _httpClient.Dispose();
    //        _disposed = true;
    //    }
    //}

    public class SynologyFileDownloader : IDisposable
    {
        private readonly SynologyFileDownloaderConfig _config;
        private readonly CookieContainer _cookieContainer;
        private readonly HttpClientHandler _handler;
        private readonly HttpClient _httpClient;
        private readonly JavaScriptSerializer _serializer;

        private bool _disposed;
        private string _sid;

        public SynologyFileDownloader(SynologyFileDownloaderConfig config)
        {
            if (config == null) throw new ArgumentNullException("config");
            if (string.IsNullOrWhiteSpace(config.BaseUrl)) throw new ArgumentException("BaseUrl이 필요합니다.");
            if (string.IsNullOrWhiteSpace(config.Username)) throw new ArgumentException("Username이 필요합니다.");
            if (string.IsNullOrWhiteSpace(config.Password)) throw new ArgumentException("Password가 필요합니다.");
            if (string.IsNullOrWhiteSpace(config.SearchFolder)) throw new ArgumentException("SearchFolder가 필요합니다.");

            _config = config;
            _cookieContainer = new CookieContainer();
            _handler = new HttpClientHandler
            {
                CookieContainer = _cookieContainer,
                UseCookies = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            _httpClient = new HttpClient(_handler);
            _httpClient.Timeout = TimeSpan.FromMinutes(10);

            _serializer = new JavaScriptSerializer();
            _serializer.MaxJsonLength = int.MaxValue;
        }

        public bool IsLoggedIn
        {
            get { return !string.IsNullOrWhiteSpace(_sid); }
        }

        public string DebugSid
        {
            get { return _sid; }
        }

        public async Task LoginAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();

            string url =
                NormalizeBaseUrl(_config.BaseUrl) + "/webapi/entry.cgi" +
                "?api=SYNO.API.Auth" +
                "&version=6" +
                "&method=login" +
                "&account=" + Uri.EscapeDataString((_config.Username ?? "").Trim()) +
                "&passwd=" + Uri.EscapeDataString((_config.Password ?? "").Trim()) +
                "&session=FileStation" +
                "&format=sid";

            string json = await GetStringAsync(url, cancellationToken).ConfigureAwait(false);

            var response = Deserialize<SynoLoginResponse>(json);

            if (response == null || !response.success || response.data == null || string.IsNullOrWhiteSpace(response.data.sid))
            {
                throw new InvalidOperationException("Synology 로그인 실패. 응답: " + json);
            }

            _sid = response.data.sid;
        }

        public async Task LogoutAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();

            if (string.IsNullOrWhiteSpace(_sid))
                return;

            string url =
                NormalizeBaseUrl(_config.BaseUrl) + "/webapi/entry.cgi" +
                "?api=SYNO.API.Auth" +
                "&version=6" +
                "&method=logout" +
                "&session=FileStation";

            try
            {
                await GetStringAsync(url, cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                // 로그아웃 실패는 무시
            }
            finally
            {
                _sid = null;
            }
        }

        /// <summary>
        /// B안: Search API를 쓰지 않고, 지정 폴더를 재귀 탐색하여
        /// 파일명 키워드/확장자 조건에 맞는 파일을 찾는다.
        /// </summary>
        public async Task<List<SynologyFileItem>> SearchFilesAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();
            EnsureLoggedIn();

            var extSet = new HashSet<string>(_config.GetNormalizedExtensions(), StringComparer.OrdinalIgnoreCase);
            var result = new List<SynologyFileItem>();

            string rootPath = NormalizeNasPath(_config.SearchFolder);
            string keyword = (_config.Keyword ?? "").Trim();

            await ListFolderRecursiveAsync(
                rootPath,
                keyword,
                extSet,
                _config.Recursive,
                result,
                cancellationToken).ConfigureAwait(false);

            // 폴더 제외, 수정일 최신순 정렬
            var files = result
                .Where(x => !x.IsDir)
                .OrderByDescending(x => x.ModifiedTime)
                .ThenBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return files;
        }

        public async Task<List<string>> DownloadFilesAsync(
            IEnumerable<SynologyFileItem> files,
            string saveFolder,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();
            EnsureLoggedIn();

            if (files == null) throw new ArgumentNullException("files");
            if (string.IsNullOrWhiteSpace(saveFolder)) throw new ArgumentException("saveFolder가 필요합니다.");

            Directory.CreateDirectory(saveFolder);

            var targets = files.Where(x => x != null && !x.IsDir).ToList();

            if (_config.DownloadLatestOnly && targets.Count > 1)
            {
                targets = targets
                    .OrderByDescending(x => x.ModifiedTime)
                    .ThenByDescending(x => x.Size)
                    .Take(1)
                    .ToList();
            }

            var savedFiles = new List<string>();

            foreach (var file in targets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string localPath = GetUniqueLocalFilePath(saveFolder, file.Name);
                await DownloadSingleFileAsync(file.Path, localPath, cancellationToken).ConfigureAwait(false);
                savedFiles.Add(localPath);
            }

            return savedFiles;
        }

        private async Task ListFolderRecursiveAsync(
    string folderPath,
    string keyword,
    HashSet<string> extSet,
    bool recursive,
    List<SynologyFileItem> result,
    CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            List<SynologyFileItem> entries = await ListFolderAsync(folderPath, cancellationToken).ConfigureAwait(false);

            foreach (var entry in entries)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (entry.IsDir)
                {
                    if (IsRecycleFolder(entry))
                        continue;

                    if (recursive)
                    {
                        await ListFolderRecursiveAsync(
                            entry.Path,
                            keyword,
                            extSet,
                            true,
                            result,
                            cancellationToken).ConfigureAwait(false);
                    }
                }
                else
                {
                    if (IsMatch(entry, keyword, extSet))
                    {
                        result.Add(entry);
                    }
                }
            }
        }

        private bool IsRecycleFolder(SynologyFileItem item)
        {
            if (item == null || !item.IsDir)
                return false;

            string name = (item.Name ?? "").Trim();

            // Synology 공유폴더 휴지통
            if (name.Equals("#recycle", StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private bool IsMatch(SynologyFileItem file, string keyword, HashSet<string> extSet)
        {
            if (file == null || file.IsDir)
                return false;

            // 키워드 필터
            if (!string.IsNullOrWhiteSpace(keyword))
            {
                if (string.IsNullOrWhiteSpace(file.Name) ||
                    file.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            // 확장자 필터
            if (extSet != null && extSet.Count > 0)
            {
                string ext = Path.GetExtension(file.Name) ?? "";
                if (!extSet.Contains(ext))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// 특정 폴더의 바로 아래 항목들(파일/폴더)을 가져온다.
        /// 필요 시 offset 기반으로 페이지를 끝까지 조회한다.
        /// </summary>
        private async Task<List<SynologyFileItem>> ListFolderAsync(
            string folderPath,
            CancellationToken cancellationToken)
        {
            var result = new List<SynologyFileItem>();
            int offset = 0;
            const int pageSize = 500;

            while (true)
            {
                var form = new Dictionary<string, string>
                {
                    { "api", "SYNO.FileStation.List" },
                    { "version", "2" },
                    { "method", "list" },
                    { "folder_path", folderPath },
                    { "offset", offset.ToString() },
                    { "limit", pageSize.ToString() },
                    { "additional", "[\"time\",\"size\"]" },
                    { "_sid", _sid }
                };

                string json = await PostFormAsync(
                    NormalizeBaseUrl(_config.BaseUrl) + "/webapi/entry.cgi",
                    form,
                    cancellationToken).ConfigureAwait(false);

                var response = Deserialize<SynoListResponse>(json);

                if (response == null)
                    throw new InvalidOperationException("폴더 목록 응답 파싱 실패. 경로: " + folderPath);

                if (!response.success)
                {
                    if (response.error != null && response.error.code == 407)
                        return new List<SynologyFileItem>();

                    throw new InvalidOperationException("폴더 목록 조회 실패. 경로: " + folderPath + ", 응답: " + json);
                }

                if (response.data == null || response.data.files == null || response.data.files.Count == 0)
                    break;

                foreach (var item in response.data.files)
                {
                    result.Add(ToSynologyFileItem(item));
                }

                // 더 이상 다음 페이지가 없으면 종료
                if (response.data.files.Count < pageSize)
                    break;

                offset += pageSize;
            }

            return result;
        }

        public async Task<List<string>> DownloadSelectedFilesAsync(
    IEnumerable<SynologyFileItem> files,
    string baseSaveFolder,
    CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();
            EnsureLoggedIn();

            if (files == null) throw new ArgumentNullException("files");
            if (string.IsNullOrWhiteSpace(baseSaveFolder)) throw new ArgumentException("baseSaveFolder가 필요합니다.");

            var selectedFiles = files
                .Where(x => x != null && !x.IsDir && x.Selected)
                .ToList();

            if (selectedFiles.Count == 0)
                return new List<string>();

            string keywordFolderName = SanitizeFolderName(
                string.IsNullOrWhiteSpace(_config.Keyword) ? "검색결과" : _config.Keyword.Trim());

            string rootSaveFolder = Path.Combine(baseSaveFolder, keywordFolderName);
            Directory.CreateDirectory(rootSaveFolder);

            var savedFiles = new List<string>();

            foreach (var file in selectedFiles)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // NAS 경로의 마지막 폴더명
                string lastFolderName = GetLastFolderNameFromFilePath(file.Path);
                lastFolderName = SanitizeFolderName(lastFolderName);

                string targetFolder = rootSaveFolder;
                if (!string.IsNullOrWhiteSpace(lastFolderName))
                {
                    targetFolder = Path.Combine(rootSaveFolder, lastFolderName);
                }

                Directory.CreateDirectory(targetFolder);

                string localPath = GetUniqueLocalFilePath(targetFolder, file.Name);

                await DownloadSingleFileAsync(file.Path, localPath, cancellationToken).ConfigureAwait(false);
                savedFiles.Add(localPath);
            }

            return savedFiles;
        }

        private string GetLastFolderNameFromFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return "";

            // 예: /A/B/C/file.xlsx  ->  마지막 폴더 = C
            string normalized = filePath.Replace("\\", "/").TrimEnd('/');

            int lastSlash = normalized.LastIndexOf('/');
            if (lastSlash <= 0)
                return "";

            string parentPath = normalized.Substring(0, lastSlash);   // /A/B/C
            int parentLastSlash = parentPath.LastIndexOf('/');

            if (parentLastSlash < 0)
                return parentPath.Trim('/');

            return parentPath.Substring(parentLastSlash + 1);
        }

        private string SanitizeFolderName(string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName))
                return "";

            string name = folderName.Trim();

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name;
        }

        private async Task DownloadSingleFileAsync(
            string filePath,
            string localPath,
            CancellationToken cancellationToken)
        {
            string url =
                NormalizeBaseUrl(_config.BaseUrl) + "/webapi/entry.cgi" +
                "?api=SYNO.FileStation.Download" +
                "&version=2" +
                "&method=download" +
                "&mode=open" +
                "&path=" + Uri.EscapeDataString(filePath) +
                "&_sid=" + Uri.EscapeDataString(_sid);

            using (var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
            {
                response.EnsureSuccessStatusCode();

                using (var input = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                using (var output = new FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await input.CopyToAsync(output).ConfigureAwait(false);
                }
            }
        }

        private async Task<string> GetStringAsync(string url, CancellationToken cancellationToken)
        {
            using (var response = await _httpClient.GetAsync(url, cancellationToken).ConfigureAwait(false))
            {
                response.EnsureSuccessStatusCode();

                byte[] bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                return Encoding.UTF8.GetString(bytes);
            }
        }

        private async Task<string> PostFormAsync(
            string url,
            Dictionary<string, string> form,
            CancellationToken cancellationToken)
        {
            using (var content = new FormUrlEncodedContent(form))
            using (var response = await _httpClient.PostAsync(url, content, cancellationToken).ConfigureAwait(false))
            {
                response.EnsureSuccessStatusCode();

                byte[] bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                return Encoding.UTF8.GetString(bytes);
            }
        }

        private T Deserialize<T>(string json) where T : class
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;

            return _serializer.Deserialize<T>(json);
        }

        private SynologyFileItem ToSynologyFileItem(SynoListFile file)
        {
            return new SynologyFileItem
            {
                Name = file.name ?? "",
                Path = file.path ?? "",
                IsDir = file.isdir,
                Size = file.additional != null ? file.additional.size : 0L,
                ModifiedTime = UnixTimeToDateTime(
                    file.additional != null && file.additional.time != null
                        ? file.additional.time.mtime
                        : 0)
            };
        }

        private DateTime UnixTimeToDateTime(long unixTime)
        {
            if (unixTime <= 0) return DateTime.MinValue;

            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return epoch.AddSeconds(unixTime).ToLocalTime();
        }

        private string GetUniqueLocalFilePath(string folder, string fileName)
        {
            string path = Path.Combine(folder, fileName);

            if (!File.Exists(path))
                return path;

            string name = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);

            int index = 1;
            while (true)
            {
                string newName = string.Format("{0} ({1}){2}", name, index, ext);
                string newPath = Path.Combine(folder, newName);

                if (!File.Exists(newPath))
                    return newPath;

                index++;
            }
        }

        private string NormalizeBaseUrl(string baseUrl)
        {
            return (baseUrl ?? "").Trim().TrimEnd('/');
        }

        private string NormalizeNasPath(string path)
        {
            string p = (path ?? "").Trim();
            if (string.IsNullOrWhiteSpace(p))
                return "/";

            if (!p.StartsWith("/"))
                p = "/" + p;

            return p;
        }

        private void EnsureLoggedIn()
        {
            if (string.IsNullOrWhiteSpace(_sid))
                throw new InvalidOperationException("Synology DSM 로그인 후 사용해야 합니다.");
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().FullName);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try { _httpClient.Dispose(); }
            catch { }

            try { _handler.Dispose(); }
            catch { }
        }

        #region DTO

        internal class SynoLoginResponse
        {
            public bool success { get; set; }
            public SynoLoginData data { get; set; }
            public SynoError error { get; set; }
        }

        internal class SynoLoginData
        {
            public string sid { get; set; }
        }

        internal class SynoError
        {
            public int code { get; set; }
        }

        internal class SynoListResponse
        {
            public bool success { get; set; }
            public SynoListData data { get; set; }
            public SynoError error { get; set; }
        }

        internal class SynoListData
        {
            public int total { get; set; }
            public int offset { get; set; }
            public List<SynoListFile> files { get; set; }
        }

        internal class SynoListFile
        {
            public string name { get; set; }
            public string path { get; set; }
            public bool isdir { get; set; }
            public SynoListAdditional additional { get; set; }
        }

        internal class SynoListAdditional
        {
            public long size { get; set; }
            public SynoListTime time { get; set; }
        }

        internal class SynoListTime
        {
            public long mtime { get; set; }
        }

        #endregion
    }

    #region DTO classes

    internal class SynoLoginResponse
    {
        public bool success { get; set; }
        public SynoLoginData data { get; set; }
    }

    internal class SynoLoginData
    {
        public string sid { get; set; }
    }

    internal class SynoSearchStartResponse
    {
        public bool success { get; set; }
        public SynoSearchStartData data { get; set; }
    }

    internal class SynoSearchStartData
    {
        public string taskid { get; set; }
    }

    internal class SynoSearchListResponse
    {
        public bool success { get; set; }
        public SynoSearchListData data { get; set; }
    }

    internal class SynoSearchListData
    {
        public bool finished { get; set; }
        public int total { get; set; }
        public List<SynoSearchFile> files { get; set; }
    }

    internal class SynoSearchFile
    {
        public bool isdir { get; set; }
        public string name { get; set; }
        public string path { get; set; }
        public SynoSearchAdditional additional { get; set; }
    }

    internal class SynoSearchAdditional
    {
        public SynoSearchTime time { get; set; }
        public long size { get; set; }
    }

    internal class SynoSearchTime
    {
        public long? mtime { get; set; }
    }

    #endregion
}