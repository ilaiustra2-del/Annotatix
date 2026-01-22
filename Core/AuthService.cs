using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace dwg2rvt.Core
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
                // Table: verification_test
                // First query: SELECT * FROM verification_test WHERE login = ?
                // We'll check password in code for better error messages
                string query = $"{SUPABASE_URL}/rest/v1/verification_test?login=eq.{Uri.EscapeDataString(login)}";
                
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
                
                var users = JArray.Parse(jsonResponse);
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] Users found: {users.Count}");
                
                // Check if user exists by login
                if (users.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH-SERVICE] User not found - login doesn't exist");
                    return new AuthResult 
                    { 
                        IsSuccess = false, 
                        ErrorMessage = "Пользователь не найден" 
                    };
                }
                
                // Get user data
                var user = users[0];
                string storedPassword = user["password"]?.ToString();
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] User found, checking password...");
                
                // Check password
                if (storedPassword != password)
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH-SERVICE] Password mismatch");
                    return new AuthResult 
                    { 
                        IsSuccess = false, 
                        ErrorMessage = "Неверный пароль" 
                    };
                }
                
                // Success - get user info
                string userId = user["user_id"]?.ToString();
                string sPlan = user["s_plan"]?.ToString() ?? "free";
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-SERVICE] SUCCESS - UserID: {userId}, Plan: {sPlan}");
                
                return new AuthResult
                {
                    IsSuccess = true,
                    UserId = userId,
                    Login = login,
                    SubscriptionPlan = sPlan
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
        /// Check if user has access to a feature based on subscription plan
        /// </summary>
        public static bool HasAccess(string feature)
        {
            if (CurrentUser == null || !CurrentUser.IsSuccess)
                return false; // Not authenticated
            
            string plan = CurrentUser.SubscriptionPlan?.ToLower() ?? "free";
            
            // Define feature access based on plans
            switch (feature.ToLower())
            {
                case "analyze":
                    // All plans can analyze
                    return true;
                
                case "annotate":
                    // Only pro and enterprise
                    return plan == "pro" || plan == "enterprise";
                
                case "place_elements":
                    // Only enterprise
                    return plan == "enterprise";
                
                default:
                    return false;
            }
        }
        
        /// <summary>
        /// Logout current user
        /// </summary>
        public static void Logout()
        {
            CurrentUser = null;
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
    }
}
