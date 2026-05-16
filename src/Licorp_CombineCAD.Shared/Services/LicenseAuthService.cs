using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Licorp_CombineCAD.Services
{
    internal enum LicenseState
    {
        Licensed,
        GraceMode,
        Expired,
        Revoked,
        Unlicensed
    }

    internal sealed class LicenseCheckResult
    {
        public LicenseState State { get; set; }
        public string Message { get; set; }
    }

    internal sealed class LicenseAuthService
    {
        private const int GraceDays = 7;
        private const string CacheVersion = "v1";
        private const string CacheIntegritySalt = "Licorp.CombineCAD.CacheSalt.2026";
        private static readonly TimeSpan TokenSkew = TimeSpan.FromMinutes(3);

        private readonly string _baseUrl;
        private readonly string _licenseKey;
        private readonly string _cacheFile;

        public LicenseAuthService()
        {
            _baseUrl = Environment.GetEnvironmentVariable("LICORP_LICENSE_API") ?? "https://your-license-server.com/api/v1";
            _licenseKey = Environment.GetEnvironmentVariable("LICORP_LICENSE_KEY") ?? string.Empty;
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            _cacheFile = Path.Combine(appData, "Licorp", "CombineCAD", "license.cache");
        }

        public LicenseCheckResult EnsureLicensed()
        {
            if (string.IsNullOrWhiteSpace(_licenseKey))
            {
                return new LicenseCheckResult
                {
                    State = LicenseState.Unlicensed,
                    Message = "Missing license key. Set LICORP_LICENSE_KEY."
                };
            }

            var fp = BuildFingerprint();
            var cache = LoadCache();

            try
            {
                var online = VerifyOnlineAsync(fp, cache).GetAwaiter().GetResult();
                if (online.State == LicenseState.Licensed || online.State == LicenseState.GraceMode)
                {
                    SaveCache(online);
                }

                return online;
            }
            catch
            {
                if (cache == null || cache.GraceUntilUtc < DateTime.UtcNow)
                {
                    return new LicenseCheckResult { State = LicenseState.Expired, Message = "License validation failed and grace window expired." };
                }

                return new LicenseCheckResult
                {
                    State = LicenseState.GraceMode,
                    Message = "Offline grace mode active."
                };
            }
        }

        private async Task<LicenseCacheModel> LoginAsync(HttpClient client, string fingerprint)
        {
            var payload = "{\"licenseKey\":\"" + Escape(_licenseKey) + "\",\"deviceFingerprint\":\"" + Escape(fingerprint) + "\"}";
            var response = await client.PostAsync(_baseUrl + "/login", new StringContent(payload, Encoding.UTF8, "application/json")).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
            var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            return ParseTokens(body);
        }

        private async Task<LicenseCheckResult> VerifyOrRefreshAsync(HttpClient client, string fingerprint, LicenseCacheModel cache)
        {
            var verifyPayload = "{\"accessToken\":\"" + Escape(cache.AccessToken) + "\",\"deviceFingerprint\":\"" + Escape(fingerprint) + "\"}";
            var verifyResponse = await client.PostAsync(_baseUrl + "/verify", new StringContent(verifyPayload, Encoding.UTF8, "application/json")).ConfigureAwait(false);

            if (verifyResponse.IsSuccessStatusCode)
            {
                cache.LastVerifiedUtc = DateTime.UtcNow;
                return BuildStateFromCache(cache);
            }

            var refreshPayload = "{\"refreshToken\":\"" + Escape(cache.RefreshToken) + "\",\"deviceFingerprint\":\"" + Escape(fingerprint) + "\"}";
            var refreshResponse = await client.PostAsync(_baseUrl + "/refresh", new StringContent(refreshPayload, Encoding.UTF8, "application/json")).ConfigureAwait(false);
            if (!refreshResponse.IsSuccessStatusCode)
            {
                if ((int)refreshResponse.StatusCode == 401 || (int)refreshResponse.StatusCode == 403)
                {
                    return new LicenseCheckResult { State = LicenseState.Revoked, Message = "License revoked or device deactivated." };
                }

                refreshResponse.EnsureSuccessStatusCode();
            }

            var body = await refreshResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
            var fresh = ParseTokens(body);
            fresh.Fingerprint = fingerprint;
            fresh.LastVerifiedUtc = DateTime.UtcNow;
            SaveCache(new LicenseCheckResult { State = LicenseState.Licensed, Message = "Licensed" }, fresh);
            return BuildStateFromCache(fresh);
        }

        private async Task<LicenseCheckResult> VerifyOnlineAsync(string fingerprint, LicenseCacheModel cache)
        {
            using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) })
            {
                if (cache == null || string.IsNullOrWhiteSpace(cache.AccessToken) || string.IsNullOrWhiteSpace(cache.RefreshToken) || cache.Fingerprint != fingerprint)
                {
                    var fresh = await LoginAsync(client, fingerprint).ConfigureAwait(false);
                    fresh.Fingerprint = fingerprint;
                    fresh.LastVerifiedUtc = DateTime.UtcNow;
                    return BuildStateFromCache(fresh);
                }

                return await VerifyOrRefreshAsync(client, fingerprint, cache).ConfigureAwait(false);
            }
        }

        private static LicenseCheckResult BuildStateFromCache(LicenseCacheModel cache)
        {
            if (cache.GraceUntilUtc < DateTime.UtcNow)
            {
                return new LicenseCheckResult { State = LicenseState.Expired, Message = "Grace window expired." };
            }

            var state = cache.AccessTokenExpiresUtc > DateTime.UtcNow.Add(TokenSkew) ? LicenseState.Licensed : LicenseState.GraceMode;
            return new LicenseCheckResult { State = state, Message = state == LicenseState.Licensed ? "Licensed" : "Grace mode" };
        }

        private void SaveCache(LicenseCheckResult result, LicenseCacheModel model = null)
        {
            var current = model ?? LoadCache();
            if (current == null)
            {
                return;
            }

            SaveCache(current);
        }

        private void SaveCache(LicenseCacheModel model)
        {
            var directory = Path.GetDirectoryName(_cacheFile);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var plaintext = string.Join("\n", new[]
            {
                model.AccessToken ?? string.Empty,
                model.RefreshToken ?? string.Empty,
                model.Fingerprint ?? string.Empty,
                model.AccessTokenExpiresUtc.Ticks.ToString(),
                model.GraceUntilUtc.Ticks.ToString(),
                model.LastVerifiedUtc.Ticks.ToString()
            });

            var protectedBytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(plaintext), null, DataProtectionScope.CurrentUser);
            var cipherTextBase64 = Convert.ToBase64String(protectedBytes);
            var checksum = ComputeSha256(cipherTextBase64 + "|" + CacheIntegritySalt + "|" + CacheVersion);
            var payload = CacheVersion + "\n" + checksum + "\n" + cipherTextBase64;
            File.WriteAllText(_cacheFile, payload, Encoding.UTF8);
        }

        private LicenseCacheModel LoadCache()
        {
            if (!File.Exists(_cacheFile))
            {
                return null;
            }

            var payload = File.ReadAllText(_cacheFile, Encoding.UTF8);
            var lines = payload.Split(new[] { '\n' }, StringSplitOptions.None);
            if (lines.Length < 3)
            {
                return null;
            }

            var version = lines[0];
            var expected = lines[1];
            var cipherTextBase64 = string.Join("\n", lines.Skip(2)).Trim();
            if (!string.Equals(version, CacheVersion, StringComparison.Ordinal))
            {
                return null;
            }

            var actual = ComputeSha256(cipherTextBase64 + "|" + CacheIntegritySalt + "|" + version);
            if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            byte[] protectedBytes;
            try
            {
                protectedBytes = Convert.FromBase64String(cipherTextBase64);
            }
            catch
            {
                return null;
            }

            var bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
            var plaintext = Encoding.UTF8.GetString(bytes);
            var tokenLines = plaintext.Split(new[] { '\n' }, StringSplitOptions.None);
            if (tokenLines.Length < 6)
            {
                return null;
            }

            long accessTicks;
            long graceTicks;
            long verifiedTicks;
            if (!long.TryParse(tokenLines[3], out accessTicks) || !long.TryParse(tokenLines[4], out graceTicks) || !long.TryParse(tokenLines[5], out verifiedTicks))
            {
                return null;
            }

            return new LicenseCacheModel
            {
                AccessToken = tokenLines[0],
                RefreshToken = tokenLines[1],
                Fingerprint = tokenLines[2],
                AccessTokenExpiresUtc = new DateTime(accessTicks, DateTimeKind.Utc),
                GraceUntilUtc = new DateTime(graceTicks, DateTimeKind.Utc),
                LastVerifiedUtc = new DateTime(verifiedTicks, DateTimeKind.Utc)
            };
        }

        private static string BuildFingerprint()
        {
            var raw = string.Join("|", new[] { Environment.MachineName, Environment.UserDomainName, Environment.OSVersion.VersionString });
            return "fpv1_" + ComputeSha256(raw).Substring(0, 32).ToLowerInvariant();
        }

        private static string ComputeSha256(string value)
        {
            using (var sha = SHA256.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
                var sb = new StringBuilder(hash.Length * 2);
                foreach (var b in hash)
                {
                    sb.Append(b.ToString("x2"));
                }

                return sb.ToString();
            }
        }

        private static string Escape(string value)
        {
            return (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        private LicenseCacheModel ParseTokens(string json)
        {
            var accessToken = MatchJsonString(json, "accessToken");
            var refreshToken = MatchJsonString(json, "refreshToken");
            var graceUntilText = MatchJsonString(json, "graceUntil");
            var expiresInText = MatchJsonNumber(json, "expiresIn");

            if (string.IsNullOrWhiteSpace(accessToken) || string.IsNullOrWhiteSpace(refreshToken) || string.IsNullOrWhiteSpace(graceUntilText))
            {
                throw new InvalidOperationException("Invalid token payload from licensing API.");
            }

            var graceUntil = DateTime.Parse(graceUntilText, null, System.Globalization.DateTimeStyles.AdjustToUniversal | System.Globalization.DateTimeStyles.AssumeUniversal);
            var expiresInSeconds = 900;
            int parsed;
            if (int.TryParse(expiresInText, out parsed) && parsed > 0)
            {
                expiresInSeconds = parsed;
            }

            return new LicenseCacheModel
            {
                AccessToken = accessToken,
                RefreshToken = refreshToken,
                GraceUntilUtc = graceUntil.ToUniversalTime(),
                AccessTokenExpiresUtc = DateTime.UtcNow.AddSeconds(expiresInSeconds)
            };
        }

        private static string MatchJsonString(string json, string key)
        {
            var m = Regex.Match(json ?? string.Empty, "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : string.Empty;
        }

        private static string MatchJsonNumber(string json, string key)
        {
            var m = Regex.Match(json ?? string.Empty, "\"" + Regex.Escape(key) + "\"\\s*:\\s*([0-9]+)");
            return m.Success ? m.Groups[1].Value : string.Empty;
        }

        private sealed class LicenseCacheModel
        {
            public string AccessToken { get; set; }
            public string RefreshToken { get; set; }
            public string Fingerprint { get; set; }
            public DateTime AccessTokenExpiresUtc { get; set; }
            public DateTime GraceUntilUtc { get; set; }
            public DateTime LastVerifiedUtc { get; set; }
        }
    }
}