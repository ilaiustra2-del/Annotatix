using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using BCrypt.Net;

namespace PluginsManager.Core
{
    /// <summary>
    /// Service for authentication with Supabase
    /// </summary>
    public class AuthService
    {
        // ⚠️ ВАЖНО: Замените SUPABASE_ANON_KEY на ваш ключ из Settings → API!
        private const string SUPABASE_URL = "https://mwlrionsymujbjvhtcgp.supabase.co";
        private const string SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13bHJpb25zeW11amJqdmh0Y2dwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjgxNTU1MjEsImV4cCI6MjA4MzczMTUyMX0.-UqfA-bcnNY9BgziQLvUhDKEyhiMsvo93Osr0GIw5TM";
        
        private static readonly HttpClient _httpClient = new HttpClient();
        
        // Current authenticated user (static for access from anywhere)
        public static AuthResult CurrentUser { get; set; }
        
        static AuthService()
        {
            // CRITICAL: Enable TLS 1.2 for .NET Framework 4.8
            // Supabase requires TLS 1.2+, but .NET 4.8 defaults to TLS 1.0/1.1
            System.Net.ServicePointManager.SecurityProtocol = 
                System.Net.SecurityProtocolType.Tls12 | 
                System.Net.SecurityProtocolType.Tls11 | 
                System.Net.SecurityProtocolType.Tls;
            
            System.Diagnostics.Debug.WriteLine("[AUTH-SERVICE] TLS 1.2 enabled");
            
            // Configure HTTP client
            _httpClient.DefaultRequestHeaders.Add("apikey", SUPABASE_ANON_KEY);
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {SUPABASE_ANON_KEY}");
        }

        /// <summary>
        /// Authenticate user with login and password
        /// </summary>
        public async Task<AuthResult> AuthenticateAsync(string login, string password)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Starting authentication for: {login}");
                
                // Build Supabase REST API query
                // Query kv_store table for key = "user_data:{user_id}" where login matches
                // We need to search all user_data keys and parse their JSON values
                string query = $"{SUPABASE_URL}/rest/v1/kv_store_19422568?key=like.user_data:*";
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Request URL: {query}");
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Sending HTTP GET request...");
                
                var response = await _httpClient.GetAsync(query);
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Response status: {response.StatusCode}");
                
                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] SERVER ERROR - Status: {response.StatusCode}");
                    string errorBody = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Error body: {errorBody}");
                    
                    return new AuthResult 
                    { 
                        IsSuccess = false, 
                        ErrorMessage = $"Ошибка сервера: {response.StatusCode}" 
                    };
                }
                
                string jsonResponse = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Response body: {jsonResponse}");
                
                var records = JArray.Parse(jsonResponse);
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Records found: {records.Count}");
                
                // Find user by login in the JSON values
                JObject userData = null;
                string userId = null;
                
                foreach (var record in records)
                {
                    string key = record["key"]?.ToString();
                    var value = JObject.Parse(record["value"]?.ToString() ?? "{}");
                    string userLogin = value["login"]?.ToString();
                    
                    if (userLogin != null && userLogin.Equals(login, StringComparison.OrdinalIgnoreCase))
                    {
                        userData = value;
                        // Extract user_id from key: "user_data:30cf2909-73e9-4ae7-a3a0-f2e06d2a0a68"
                        userId = key?.Replace("user_data:", "");
                        System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] User found with key: {key}");
                        break;
                    }
                }
                
                // Check if user exists by login
                if (userData == null)
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH-SERVICE] User not found - login doesn't exist");
                    return new AuthResult 
                    { 
                        IsSuccess = false, 
                        ErrorMessage = "Пользователь не найден" 
                    };
                }
                
                // Get user data
                string storedPasswordHash = userData["password"]?.ToString();
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] User found, checking password...");
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Stored hash: {storedPasswordHash}");
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Input password length: {password?.Length}");
                
                // Check password - support both SHA-256 and BCrypt
                bool isPasswordValid = false;
                
                // Detect hash type by format
                if (storedPasswordHash != null && storedPasswordHash.Length == 64 && IsHexString(storedPasswordHash))
                {
                    // SHA-256 hash (64 hex characters)
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Detected SHA-256 hash");
                    string inputPasswordHash = ComputeSha256Hash(password);
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Computed hash: {inputPasswordHash}");
                    isPasswordValid = storedPasswordHash.Equals(inputPasswordHash, StringComparison.OrdinalIgnoreCase);
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] SHA-256 result: {isPasswordValid}");
                }
                else if (storedPasswordHash != null && storedPasswordHash.StartsWith("$2"))
                {
                    // BCrypt hash
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Detected BCrypt hash");
                    try
                    {
                        isPasswordValid = BCrypt.Net.BCrypt.Verify(password, storedPasswordHash);
                        System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] BCrypt result: {isPasswordValid}");
                    }
                    catch (Exception bcryptEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] BCrypt error: {bcryptEx.Message}");
                    }
                }
                else
                {
                    // Plain text or unknown format
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Unknown hash format, trying plain text");
                    isPasswordValid = (storedPasswordHash == password);
                }
                
                if (!isPasswordValid)
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH-SERVICE] Password mismatch");
                    
                    // Provide more detailed error for debugging
                    string errorDetails = "Неверный пароль";
                    if (storedPasswordHash != null && storedPasswordHash.StartsWith("$2"))
                    {
                        errorDetails += " (BCrypt hash detected)";
                    }
                    else if (storedPasswordHash != null)
                    {
                        errorDetails += $" (Hash type: unknown, length: {storedPasswordHash.Length})";
                    }
                    
                    return new AuthResult 
                    { 
                        IsSuccess = false, 
                        ErrorMessage = errorDetails
                    };
                }
                
                // Success - get user info
                // No subscription plan field in DB yet
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] SUCCESS - UserID: {userId}");
                
                // Now get user modules from kv_store table
                var modules = await GetUserModulesAsync(userId);
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Retrieved {modules.Count} modules");
                
                return new AuthResult
                {
                    IsSuccess = true,
                    UserId = userId,
                    Login = login,
                    SubscriptionPlan = null, // Not used yet
                    Modules = modules
                };
            }
            catch (HttpRequestException httpEx)
            {
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] HTTP EXCEPTION - Network error: {httpEx.Message}");
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] InnerException: {httpEx.InnerException?.Message}");
                
                string detailedError = $"Ошибка сети: {httpEx.Message}";
                if (httpEx.InnerException != null)
                {
                    detailedError += $"\n\nВнутренняя ошибка: {httpEx.InnerException.Message}";
                }
                
                return new AuthResult
                {
                    IsSuccess = false,
                    ErrorMessage = detailedError
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] EXCEPTION: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Stack trace: {ex.StackTrace}");
                
                return new AuthResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"Ошибка плагина: {ex.Message}"
                };
            }
        }
        
        /// <summary>
        /// Get user modules from kv_store table
        /// </summary>
        private async Task<List<UserModule>> GetUserModulesAsync(string userId)
        {
            var modules = new List<UserModule>();
            
            try
            {
                // Query kv_store table for keys starting with "user_modules:{userId}"
                // Example: user_modules:fc2c7a31-1677-40c5-a02b-c4cc5fbe0895:dwg2rvt
                string keyPrefix = $"user_modules:{userId}";
                string query = $"{SUPABASE_URL}/rest/v1/kv_store_19422568?key=like.{Uri.EscapeDataString(keyPrefix + "*")}";
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Fetching modules with query: {query}");
                
                var response = await _httpClient.GetAsync(query);
                
                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Failed to fetch modules: {response.StatusCode}");
                    return modules; // Return empty list
                }
                
                string jsonResponse = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Modules response: {jsonResponse}");
                
                var moduleRecords = JArray.Parse(jsonResponse);
                
                foreach (var record in moduleRecords)
                {
                    var value = JObject.Parse(record["value"]?.ToString() ?? "{}");
                    
                    string moduleTag = value["module_tag"]?.ToString();
                    string startDateStr = value["start_date"]?.ToString();
                    string endDateStr = value["end_date"]?.ToString();
                    
                    if (string.IsNullOrEmpty(moduleTag))
                        continue;
                    
                    DateTime startDate = DateTime.TryParse(startDateStr, out var sd) ? sd : DateTime.MinValue;
                    DateTime endDate = DateTime.TryParse(endDateStr, out var ed) ? ed : DateTime.MaxValue;
                    
                    modules.Add(new UserModule
                    {
                        ModuleTag = moduleTag,
                        StartDate = startDate,
                        EndDate = endDate
                    });
                    
                    System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Module: {moduleTag}, Active: {startDate <= DateTime.Now && DateTime.Now <= endDate}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Error fetching modules: {ex.Message}");
            }
            
            return modules;
        }
        
        /// <summary>
        /// Check if user has access to a module
        /// </summary>
        public static bool HasAccess(string moduleTag)
        {
            if (CurrentUser == null || !CurrentUser.IsSuccess)
                return false; // Not authenticated
            
            // Check if user has the module and it's active
            var module = CurrentUser.Modules?.FirstOrDefault(m => 
                m.ModuleTag.Equals(moduleTag, StringComparison.OrdinalIgnoreCase));
            
            if (module == null)
                return false; // Module not found
            
            // Check if module subscription is still valid
            return module.IsActive;
        }
        
        /// <summary>
        /// Logout current user
        /// </summary>
        public static void Logout()
        {
            CurrentUser = null;
        }
        
        /// <summary>
        /// Compute SHA-256 hash of a string
        /// </summary>
        private static string ComputeSha256Hash(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;
            
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
                StringBuilder builder = new StringBuilder();
                foreach (byte b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }
        
        /// <summary>
        /// Check if string contains only hexadecimal characters
        /// </summary>
        private static bool IsHexString(string str)
        {
            if (string.IsNullOrEmpty(str))
                return false;
            
            foreach (char c in str)
            {
                if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')))
                    return false;
            }
            return true;
        }
    }
    
    /// <summary>
    /// Authentication result
    /// </summary>
    public class AuthResult
    {
        public bool IsSuccess { get; set; }
        public string UserId { get; set; }
        public string Login { get; set; }
        public string SubscriptionPlan { get; set; }
        public string ErrorMessage { get; set; }
        public List<UserModule> Modules { get; set; } = new List<UserModule>();
    }
    
    /// <summary>
    /// User module information
    /// </summary>
    public class UserModule
    {
        public string ModuleTag { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public bool IsActive => DateTime.Now >= StartDate && DateTime.Now <= EndDate;
    }
}
