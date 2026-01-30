using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace PluginsManager.Core
{
    /// <summary>
    /// Local storage for authentication data
    /// Stores refresh token and user email in %APPDATA%\Annotatix\
    /// </summary>
    public static class LocalAuthStorage
    {
        private static readonly string AppDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Annotatix");
        
        private static readonly string AuthTokenFile = Path.Combine(AppDataFolder, "auth_token.dat");
        
        // Simple encryption key (for basic obfuscation, not high security)
        private static readonly byte[] EncryptionKey = Encoding.UTF8.GetBytes("Annotatix2024Key!");
        
        /// <summary>
        /// Save authentication data locally
        /// </summary>
        public static bool SaveAuthData(string email, string refreshToken)
        {
            try
            {
                // Create directory if not exists
                if (!Directory.Exists(AppDataFolder))
                {
                    Directory.CreateDirectory(AppDataFolder);
                    DebugLogger.Log($"[AUTH-STORAGE] Created directory: {AppDataFolder}");
                }
                
                // Prepare data: email|refreshToken|timestamp
                var timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                var data = $"{email}|{refreshToken}|{timestamp}";
                
                // Encrypt data (basic XOR encryption for obfuscation)
                var encryptedData = EncryptString(data);
                
                // Write to file
                File.WriteAllText(AuthTokenFile, encryptedData);
                
                DebugLogger.Log($"[AUTH-STORAGE] Auth data saved for: {email}");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[AUTH-STORAGE] ERROR saving auth data: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Load authentication data from local storage
        /// </summary>
        /// <returns>Tuple of (email, refreshToken, timestamp) or null if not found/invalid</returns>
        public static (string email, string refreshToken, long timestamp)? LoadAuthData()
        {
            try
            {
                if (!File.Exists(AuthTokenFile))
                {
                    DebugLogger.Log("[AUTH-STORAGE] No saved auth data found");
                    return null;
                }
                
                // Read encrypted data
                var encryptedData = File.ReadAllText(AuthTokenFile);
                
                // Decrypt data
                var data = DecryptString(encryptedData);
                
                // Parse: email|refreshToken|timestamp
                var parts = data.Split('|');
                if (parts.Length != 3)
                {
                    DebugLogger.Log("[AUTH-STORAGE] Invalid auth data format");
                    return null;
                }
                
                var email = parts[0];
                var refreshToken = parts[1];
                var timestamp = long.Parse(parts[2]);
                
                // Check if token is not too old (e.g., 30 days)
                var currentTimestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                var daysSinceAuth = (currentTimestamp - timestamp) / 86400;
                
                if (daysSinceAuth > 30)
                {
                    DebugLogger.Log($"[AUTH-STORAGE] Auth data too old ({daysSinceAuth} days), clearing");
                    ClearAuthData();
                    return null;
                }
                
                DebugLogger.Log($"[AUTH-STORAGE] Auth data loaded for: {email} (saved {daysSinceAuth} days ago)");
                return (email, refreshToken, timestamp);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[AUTH-STORAGE] ERROR loading auth data: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Clear saved authentication data
        /// </summary>
        public static bool ClearAuthData()
        {
            try
            {
                if (File.Exists(AuthTokenFile))
                {
                    File.Delete(AuthTokenFile);
                    DebugLogger.Log("[AUTH-STORAGE] Auth data cleared");
                }
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[AUTH-STORAGE] ERROR clearing auth data: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Check if authentication data exists
        /// </summary>
        public static bool HasSavedAuth()
        {
            return File.Exists(AuthTokenFile);
        }
        
        // Simple XOR encryption for basic obfuscation
        private static string EncryptString(string plainText)
        {
            var plainBytes = Encoding.UTF8.GetBytes(plainText);
            var encryptedBytes = new byte[plainBytes.Length];
            
            for (int i = 0; i < plainBytes.Length; i++)
            {
                encryptedBytes[i] = (byte)(plainBytes[i] ^ EncryptionKey[i % EncryptionKey.Length]);
            }
            
            return Convert.ToBase64String(encryptedBytes);
        }
        
        private static string DecryptString(string encryptedText)
        {
            var encryptedBytes = Convert.FromBase64String(encryptedText);
            var decryptedBytes = new byte[encryptedBytes.Length];
            
            for (int i = 0; i < encryptedBytes.Length; i++)
            {
                decryptedBytes[i] = (byte)(encryptedBytes[i] ^ EncryptionKey[i % EncryptionKey.Length]);
            }
            
            return Encoding.UTF8.GetString(decryptedBytes);
        }
    }
}
