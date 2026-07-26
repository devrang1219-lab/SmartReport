using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace SynologyIntegration
{
    public class SynologyFileUploader : IDisposable
    {
        private readonly SynologyFileDownloaderConfig _config;
        private readonly HttpClient _httpClient;
        private readonly HttpClientHandler _handler;
        private readonly JavaScriptSerializer _serializer;
        private readonly bool _ownsHttpClient;
        private string _sid;
        private bool _disposed;

        public SynologyFileUploader(SynologyFileDownloaderConfig config)
            : this(config, null, null)
        {
        }

        /// <summary>
        /// 생성자: 외부에서 준비한 HttpClient와 초기 SID를 전달할 수 있습니다.
        /// HttpClient를 전달하면 인스턴스는 해당 HttpClient를 소유하지 않으므로 Dispose 시 해제하지 않습니다.
        /// </summary>
        public SynologyFileUploader(SynologyFileDownloaderConfig config, HttpClient httpClient, string sid = null)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));

            _config = config;
            _serializer = new JavaScriptSerializer();

            if (httpClient != null)
            {
                _httpClient = httpClient;
                _handler = null;
                _ownsHttpClient = false;
                _sid = sid; // may be null; LoginAsync will set if needed
            }
            else
            {
                var cookieContainer = new CookieContainer();
                _handler = new HttpClientHandler
                {
                    CookieContainer = cookieContainer,
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
                };

                _httpClient = new HttpClient(_handler)
                {
                    Timeout = TimeSpan.FromMinutes(10)
                };
                _ownsHttpClient = true;
            }
        }

        public bool IsLoggedIn => !string.IsNullOrWhiteSpace(_sid);

        public async Task LoginAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();

            string url = _config.BaseUrl.TrimEnd('/') + "/webapi/auth.cgi" +
                "?api=SYNO.API.Auth&version=6&method=login" +
                "&account=" + Uri.EscapeDataString(_config.Username) +
                "&passwd=" + Uri.EscapeDataString(_config.Password) +
                "&session=FileStation&format=sid";

            string json = await GetStringAsync(url, cancellationToken).ConfigureAwait(false);
            var response = _serializer.Deserialize<SynoLoginResponse>(json);
            if (response == null || !response.success || response.data == null || string.IsNullOrWhiteSpace(response.data.sid))
                throw new InvalidOperationException("Synology 로그인 실패. 응답: " + json);

            _sid = response.data.sid;
        }

        public async Task UploadFilesAsync(IEnumerable<string> localFiles, string remoteFolder, CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();
            EnsureLoggedIn();

            if (localFiles == null) return;


            File.WriteAllText(@"C:\temp\test.txt", "abc");
            await UploadFileAsync(@"C:\temp\test.txt", remoteFolder);

            foreach (var f in localFiles)
            {
                if (string.IsNullOrEmpty(f) || !File.Exists(f)) continue;
                await UploadFileAsync(f, remoteFolder, cancellationToken).ConfigureAwait(false);
            }
        }

        public async Task UploadFileAsync(string localFilePath, string remoteFolder, CancellationToken cancellationToken = default(CancellationToken))
        {
            ThrowIfDisposed();
            EnsureLoggedIn();

            if (string.IsNullOrEmpty(localFilePath) || !File.Exists(localFilePath))
                throw new FileNotFoundException("로컬 파일을 찾을 수 없습니다.", localFilePath);

            //string url = _config.BaseUrl.TrimEnd('/') + "/webapi/entry.cgi";
            Debug.WriteLine($"upload sid = [{_sid}]");

            string url =
    _config.BaseUrl.TrimEnd('/') +
    "/webapi/entry.cgi" +
    "?api=SYNO.FileStation.Upload" +
    "&method=upload" +
    "&version=2";// +
   // "&SynoToken=" + synoToken;


            remoteFolder = remoteFolder
                .Replace("\r", "")
                .Replace("\n", "")
                .Trim();

            using (var content = new MultipartFormDataContent())
            using (var fileStream = File.OpenRead(localFilePath))
            {
                // API 파라미터
                //content.Add(new StringContent("SYNO.FileStation.Upload"), "api");
                //content.Add(new StringContent("3"), "version");
                //content.Add(new StringContent("upload"), "method");
                //content.Add(new StringContent(remoteFolder ?? "/"), "path");
                //content.Add(new StringContent("true"), "create_parents");
                //content.Add(new StringContent("true"), "overwrite");
                //content.Add(new StringContent(_sid), "_sid");

                content.Add(
        new StringContent(DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString()),
        "mtime");

                content.Add(
                    new StringContent("true"),
                    "overwrite");

                content.Add(
                    new StringContent(remoteFolder),
                    "path");

                var fi = new FileInfo(localFilePath);

                content.Add(
                    new StringContent(fi.Length.ToString()),
                    "size");

                Debug.WriteLine($"api=SYNO.FileStation.Upload");
                Debug.WriteLine($"version=3");
                Debug.WriteLine($"method=upload");
                Debug.WriteLine($"path=[{remoteFolder}]");
                Debug.WriteLine($"sid=[{_sid}]");
                Debug.WriteLine($"file=[{Path.GetFileName(localFilePath)}]");

                var streamContent = new StreamContent(fileStream);
                //streamContent.Headers.Add("Content-Type", "application/octet-stream");
                streamContent.Headers.ContentType =
                    new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                // name must be 'file' according to FileStation upload
                content.Add(streamContent, "file", Path.GetFileName(localFilePath));

                foreach (var item in content)
                {
                    Debug.WriteLine(item.Headers.ContentDisposition);
                }

                Debug.WriteLine(content.Headers.ContentType);


                _httpClient.DefaultRequestHeaders.Remove("X-SYNO-TOKEN");
                //_httpClient.DefaultRequestHeaders.Add(
                //    "X-SYNO-TOKEN",
                //    synoToken);

                using (var resp = await _httpClient.PostAsync(url, content, cancellationToken).ConfigureAwait(false))
                {
                    resp.EnsureSuccessStatusCode();
                    var bytes = await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                    string json = Encoding.UTF8.GetString(bytes);
                    Debug.WriteLine("STATUS = " + (int)resp.StatusCode);
                    Debug.WriteLine("RESPONSE = " + json);
                    var upResp = _serializer.Deserialize<GenericSynoResponse>(json);
                    if (upResp == null || !upResp.success)
                        throw new InvalidOperationException("파일 업로드 실패: " + json);
                }
            }
        }

        public async Task<bool> CreateFolderAsync(string parentPath, string folderName, CancellationToken cancellationToken = default)
        {
            ThrowIfDisposed();
            EnsureLoggedIn();

            string url = _config.BaseUrl.TrimEnd('/') + "/webapi/entry.cgi";

            var form = new Dictionary<string, string>
            {
                { "api", "SYNO.FileStation.CreateFolder" },
                { "version", "2" },
                { "method", "create" },
                { "path", parentPath ?? "/" },
                { "name", folderName },
                { "_sid", _sid }
            };

            using (var content = new FormUrlEncodedContent(form))
            using (var resp = await _httpClient.PostAsync(url, content, cancellationToken).ConfigureAwait(false))
            {
                resp.EnsureSuccessStatusCode();
                var bytes = await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                string json = Encoding.UTF8.GetString(bytes);
                var r = _serializer.Deserialize<GenericSynoResponse>(json);
                return r != null && r.success;
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

        private void EnsureLoggedIn()
        {
            if (string.IsNullOrWhiteSpace(_sid)) throw new InvalidOperationException("Synology DSM 로그인 후 사용해야 합니다.");
        }

        private void ThrowIfDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(GetType().FullName);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_ownsHttpClient)
            {
                try { _httpClient.Dispose(); } catch { }
                try { _handler?.Dispose(); } catch { }
            }
        }

        #region DTOs
        private class GenericSynoResponse
        {
            public bool success { get; set; }
            public object data { get; set; }
            public object error { get; set; }
        }

        #endregion
    }
}
