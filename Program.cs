using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Configuration;

namespace ToDoMaintenance;

class Program
{
    private static IConfiguration? _configuration;
    private static GraphServiceClient? _graphClient;
    private static string? _targetListId;
    private static List<TodoTask> _completedTasks = new();

    static async Task Main(string[] args)
    {
        try
        {
            // Step 1: Launch Application
            DisplayWelcomeMessage();

            // Load configuration
            LoadConfiguration();

            // Step 2: Authentication
            await AuthenticateAsync();

            // Step 3: List Discovery & Task Retrieval
            await DiscoverAndAnalyzeTasksAsync();

            // Step 4: Verification Display
            DisplayVerificationSummary();

            // Step 5: User Approval
            if (!PromptForUserApproval())
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\n⚠️  Operation cancelled by user. No tasks were deleted.");
                Console.ResetColor();
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
                return;
            }

            // Step 6: Delete Completed Tasks
            await DeleteCompletedTasksAsync();

            Console.WriteLine("\n✅ Task cleanup completed successfully!");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n❌ Error: {ex.Message}");
            Console.ResetColor();
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
    }

    /// <summary>
    /// Step 1: Display welcome message and application purpose
    /// </summary>
    private static void DisplayWelcomeMessage()
    {
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("╔════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║     Microsoft 365 To-Do Maintenance Tool                 ║");
        Console.WriteLine("║     Completed Task Cleanup Utility                        ║");
        Console.WriteLine("╚════════════════════════════════════════════════════════════╝");
        Console.ResetColor();
        Console.WriteLine();
        Console.WriteLine("Purpose: This tool will help you clean up completed tasks");
        Console.WriteLine("         from your Microsoft To-Do 'Tasks' list to improve");
        Console.WriteLine("         application performance.");
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("⚠️  Safety: Only tasks with 'Completed' status will be deleted.");
        Console.WriteLine("           You will be asked to approve before any deletion occurs.");
        Console.ResetColor();
        Console.WriteLine();
        Console.WriteLine("═══════════════════════════════════════════════════════════════");
        Console.WriteLine();
    }

    /// <summary>
    /// Load configuration from appsettings.json and environment variables
    /// </summary>
    private static void LoadConfiguration()
    {
        _configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();

        // Validate configuration
        var clientId = _configuration["AzureAd:ClientId"];
        if (string.IsNullOrEmpty(clientId) || clientId == "YOUR_CLIENT_ID_HERE")
        {
            throw new InvalidOperationException(
                "Azure AD Client ID is not configured. Please update appsettings.json with your Azure AD app registration Client ID.");
        }
    }

    /// <summary>
    /// Step 2: Authenticate with Microsoft 365 using interactive browser flow
    /// </summary>
    private static async Task AuthenticateAsync()
    {
        Console.WriteLine("🔐 Step 2: Authentication");
        Console.WriteLine("─────────────────────────────────────────────────────────────");
        Console.WriteLine();

        var clientId = _configuration!["AzureAd:ClientId"]!;
        var tenantId = _configuration["AzureAd:TenantId"] ?? "common";
        var scopes = _configuration.GetSection("AzureAd:Scopes").Get<string[]>() 
                     ?? new[] { "Tasks.ReadWrite", "User.Read" };

        Console.WriteLine($"Client ID: {clientId}");
        Console.WriteLine($"Tenant: {tenantId}");
        Console.WriteLine($"Scopes: {string.Join(", ", scopes)}");
        Console.WriteLine();
        Console.WriteLine("Opening browser for authentication...");
        Console.WriteLine();

        try
        {
            // Use InteractiveBrowserCredential for the most user-friendly experience
            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = tenantId,
                ClientId = clientId,
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                RedirectUri = new Uri("http://localhost")
            };

            var credential = new InteractiveBrowserCredential(options);

            // Create Graph client
            _graphClient = new GraphServiceClient(credential, scopes);

            // Test authentication by getting user info
            var user = await _graphClient.Me.GetAsync();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("✅ Authentication successful!");
            Console.ResetColor();
            Console.WriteLine($"   Signed in as: {user?.DisplayName} ({user?.UserPrincipalName})");
            Console.WriteLine();
        }
        catch (Azure.Identity.AuthenticationFailedException authEx)
        {
            throw new Exception($"Authentication failed: {authEx.Message}. Please ensure your Azure AD app is properly configured.", authEx);
        }
    }

    /// <summary>
    /// Step 3: Discover and locate the target task list, then retrieve and analyze all tasks
    /// </summary>
    private static async Task DiscoverAndAnalyzeTasksAsync()
    {
        Console.WriteLine("🔍 Step 3: Task List Discovery & Analysis");
        Console.WriteLine("─────────────────────────────────────────────────────────────");
        Console.WriteLine();

        var targetListName = _configuration!["ToDoSettings:TargetListName"] ?? "Tasks";
        Console.WriteLine($"Searching for task list: '{targetListName}'");
        Console.WriteLine();

        try
        {
            // Get all task lists with retry logic for throttling
            var listsResponse = await RetryWithThrottlingAsync(
                async () => await _graphClient!.Me.Todo.Lists.GetAsync(),
                "Getting task lists");
            var lists = listsResponse?.Value;

            if (lists == null || !lists.Any())
            {
                throw new Exception("No task lists found in your Microsoft To-Do account.");
            }

            Console.WriteLine($"Found {lists.Count} task list(s):");
            foreach (var list in lists)
            {
                var indicator = list.DisplayName?.Equals(targetListName, StringComparison.OrdinalIgnoreCase) == true ? "→" : " ";
                Console.WriteLine($"  {indicator} {list.DisplayName} (ID: {list.Id})");
            }
            Console.WriteLine();

            // Find the target list
            var targetList = lists.FirstOrDefault(l => 
                l.DisplayName?.Equals(targetListName, StringComparison.OrdinalIgnoreCase) == true);

            if (targetList == null)
            {
                throw new Exception($"Task list '{targetListName}' not found. Please check the list name in appsettings.json.");
            }

            _targetListId = targetList.Id;

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"✅ Successfully located target list: '{targetList.DisplayName}'");
            Console.ResetColor();
            Console.WriteLine($"   List ID: {targetList.Id}");
            Console.WriteLine($"   Well-known Name: {targetList.WellknownListName}");
            Console.WriteLine();

            // Get task count (handling pagination and throttling)
            Console.Write("   Counting tasks...");
            var taskCount = 0;
            _completedTasks.Clear();
            
            // Get initial page with retry logic for throttling
            var tasksResponse = await RetryWithThrottlingAsync(
                async () => await _graphClient!.Me.Todo.Lists[_targetListId].Tasks.GetAsync(),
                "Getting tasks");
            
            if (tasksResponse != null)
            {
                var pageIterator = Microsoft.Graph.PageIterator<TodoTask, TodoTaskCollectionResponse>
                    .CreatePageIterator(
                        _graphClient!,
                        tasksResponse,
                        (task) => {
                            taskCount++;
                            if (task.Status == Microsoft.Graph.Models.TaskStatus.Completed)
                            {
                                _completedTasks.Add(task);
                            }
                            return true; // Continue iterating
                        });
                
                // Iterate through all pages with retry logic for throttling
                await RetryWithThrottlingAsync(
                    async () => { await pageIterator.IterateAsync(); return true; },
                    "Iterating through task pages");
            }
            
            Console.WriteLine($" Done!");
            Console.WriteLine($"   Total tasks in list: {taskCount}");
            Console.WriteLine($"   Completed tasks: {_completedTasks.Count}");
            Console.WriteLine($"   Active tasks: {taskCount - _completedTasks.Count}");
            Console.WriteLine();
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
        {
            throw new Exception($"Microsoft Graph API error: {odataEx.Error?.Message ?? odataEx.Message}", odataEx);
        }
    }

    /// <summary>
    /// Step 4: Display pre-deletion verification and summary
    /// </summary>
    private static void DisplayVerificationSummary()
    {
        Console.WriteLine("📋 Step 4: Pre-Deletion Verification");
        Console.WriteLine("─────────────────────────────────────────────────────────────");
        Console.WriteLine();

        if (_completedTasks.Count == 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("✅ No completed tasks found. Your task list is already clean!");
            Console.ResetColor();
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        Console.WriteLine($"Tasks ready for deletion: {_completedTasks.Count}");
        Console.WriteLine();

        // Show sample of tasks to be deleted (first 10)
        var sampleSize = Math.Min(10, _completedTasks.Count);
        Console.WriteLine($"Sample of tasks to be deleted (showing {sampleSize} of {_completedTasks.Count}):");
        Console.WriteLine();

        for (int i = 0; i < sampleSize; i++)
        {
            var task = _completedTasks[i];
            var title = string.IsNullOrEmpty(task.Title) ? "(No title)" : task.Title;
            var completedDate = task.CompletedDateTime?.DateTime ?? "Unknown";
            Console.WriteLine($"  {i + 1}. {title}");
            Console.WriteLine($"     Completed: {completedDate}");
        }

        if (_completedTasks.Count > sampleSize)
        {
            Console.WriteLine($"     ... and {_completedTasks.Count - sampleSize} more");
        }

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("⚠️  WARNING: This operation cannot be undone!");
        Console.ResetColor();
        Console.WriteLine();
    }

    /// <summary>
    /// Step 5: Prompt user for approval before deletion
    /// </summary>
    private static bool PromptForUserApproval()
    {
        Console.WriteLine("🔐 Step 5: User Approval Required");
        Console.WriteLine("─────────────────────────────────────────────────────────────");
        Console.WriteLine();

        var dryRun = _configuration!.GetValue<bool>("ToDoSettings:DryRun");
        
        if (dryRun)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("🧪 DRY RUN MODE: No tasks will actually be deleted.");
            Console.ResetColor();
            Console.WriteLine();
        }

        Console.Write($"Do you want to delete {_completedTasks.Count} completed task(s)? (yes/no): ");
        
        while (true)
        {
            var input = Console.ReadLine()?.Trim().ToLowerInvariant();
            
            if (input == "yes" || input == "y")
            {
                return true;
            }
            else if (input == "no" || input == "n")
            {
                return false;
            }
            else
            {
                Console.Write("Please enter 'yes' or 'no': ");
            }
        }
    }

    /// <summary>
    /// Step 6: Delete completed tasks with batch processing, throttling handling, and progress feedback
    /// </summary>
    private static async Task DeleteCompletedTasksAsync()
    {
        Console.WriteLine();
        Console.WriteLine("🗑️  Step 6: Deleting Completed Tasks");
        Console.WriteLine("─────────────────────────────────────────────────────────────");
        Console.WriteLine();

        var dryRun = _configuration!.GetValue<bool>("ToDoSettings:DryRun");
        
        if (dryRun)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("🧪 DRY RUN MODE - Simulating deletion (no actual changes)");
            Console.ResetColor();
            Console.WriteLine();
        }

        Console.WriteLine("Using batch processing with throttling protection...");
        Console.WriteLine();

        var successCount = 0;
        var errorCount = 0;
        var skippedCount = 0;
        var throttleCount = 0;
        var errors = new List<string>();
        var startTime = DateTime.UtcNow;

        // Process tasks in batches of 20 (Graph API batch limit is 20)
        const int batchSize = 20;
        var totalBatches = (int)Math.Ceiling(_completedTasks.Count / (double)batchSize);

        for (int batchIndex = 0; batchIndex < totalBatches; batchIndex++)
        {
            var batchTasks = _completedTasks
                .Skip(batchIndex * batchSize)
                .Take(batchSize)
                .ToList();

            var batchNum = batchIndex + 1;
            var processedCount = batchIndex * batchSize;
            Console.Write($"\r  Processing batch {batchNum}/{totalBatches} ({processedCount}/{_completedTasks.Count} tasks)...");

            if (!dryRun)
            {
                // Track tasks that need to be deleted (initially all tasks in this batch)
                var tasksToDelete = new List<(string TaskId, string Title)>();
                
                foreach (var task in batchTasks)
                {
                    if (task.Id == null)
                    {
                        var title = string.IsNullOrEmpty(task.Title) ? "(No title)" : task.Title;
                        errors.Add($"Task '{title}': Skipped (null ID)");
                        skippedCount++;
                        continue;
                    }
                    
                    var taskTitle = string.IsNullOrEmpty(task.Title) ? "(No title)" : task.Title;
                    tasksToDelete.Add((task.Id, taskTitle));
                }

                // Skip empty batches
                if (tasksToDelete.Count == 0)
                {
                    continue;
                }

                // Retry logic - only retry tasks that failed or were throttled
                var batchRetryCount = 0;
                const int maxBatchRetries = 5;
                var batchRetryDelay = TimeSpan.FromSeconds(2);

                while (tasksToDelete.Count > 0 && batchRetryCount < maxBatchRetries)
                {
                    // Add delay before batch execution to avoid throttling
                    await Task.Delay(500);
                    
                    // Build batch request with DELETE operations for remaining tasks
                    var batchRequestContent = new Microsoft.Graph.BatchRequestContentCollection(_graphClient!);
                    var requestIds = new Dictionary<string, (string TaskId, string Title)>();
                    
                    foreach (var (taskId, title) in tasksToDelete)
                    {
                        var requestInfo = _graphClient!.Me.Todo.Lists[_targetListId].Tasks[taskId].ToDeleteRequestInformation();
                        var requestId = await batchRequestContent.AddBatchRequestStepAsync(requestInfo);
                        requestIds[requestId] = (taskId, title);
                    }
                    
                    Microsoft.Graph.BatchResponseContentCollection? batchResponse = null;
                    
                    try
                    {
                        // Execute the batch
                        batchResponse = await _graphClient!.Batch.PostAsync(batchRequestContent);
                        
                        // Check responses and collect tasks that need retry
                        var throttledTasks = new List<(string TaskId, string Title)>();
                        
                        foreach (var kvp in requestIds)
                        {
                            var requestId = kvp.Key;
                            var (taskId, title) = kvp.Value;

                            try
                            {
                                var response = await batchResponse.GetResponseByIdAsync(requestId);
                                
                                if (response.IsSuccessStatusCode)
                                {
                                    // Successfully deleted
                                    successCount++;
                                }
                                else if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                                {
                                    // Task already deleted (possibly from previous retry or external deletion)
                                    successCount++;
                                }
                                else if (response.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
                                {
                                    // Throttled - add to retry list
                                    throttledTasks.Add((taskId, title));
                                }
                                else
                                {
                                    // Other error - don't retry, record as error
                                    var errorContent = await response.Content.ReadAsStringAsync();
                                    errors.Add($"Task '{title}': HTTP {response.StatusCode} - {errorContent}");
                                    errorCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                // Exception getting response - don't retry, record as error
                                errors.Add($"Task '{title}': {ex.Message}");
                                errorCount++;
                            }
                        }
                        
                        // Update tasks to delete for next iteration (only throttled tasks)
                        tasksToDelete = throttledTasks;
                        
                        // If we have throttled tasks, prepare for retry
                        if (throttledTasks.Count > 0)
                        {
                            throttleCount++;
                            batchRetryCount++;
                            
                            if (batchRetryCount < maxBatchRetries)
                            {
                                // Try to get Retry-After from throttled responses, fallback to exponential backoff
                                var retryAfterSeconds = await GetRetryAfterFromBatchResponseAsync(batchResponse, requestIds.Where(r => throttledTasks.Any(t => t.TaskId == r.Value.TaskId)).Select(r => r.Key));
                                batchRetryDelay = retryAfterSeconds > 0 
                                    ? TimeSpan.FromSeconds(retryAfterSeconds)
                                    : TimeSpan.FromSeconds(Math.Pow(2, batchRetryCount)); // Exponential backoff: 2s, 4s, 8s, 16s, 32s
                                
                                var waitSeconds = (int)batchRetryDelay.TotalSeconds;
                                var delaySource = retryAfterSeconds > 0 ? "Retry-After header" : "exponential backoff";
                                
                                Console.Write($"\r  ⚠️  {throttledTasks.Count} task(s) throttled! Waiting {waitSeconds}s ({delaySource}) before retry {batchRetryCount}/{maxBatchRetries}...          ");
                                await Task.Delay(batchRetryDelay);
                                Console.Write($"\r  Processing batch {batchNum}/{totalBatches} ({processedCount}/{_completedTasks.Count} tasks)...");
                            }
                            else
                            {
                                // Max retries reached - mark throttled tasks as errors
                                foreach (var throttled in throttledTasks)
                                {
                                    errors.Add($"Task '{throttled.Title}': Throttled after {maxBatchRetries} retries");
                                    errorCount++;
                                }
                                tasksToDelete.Clear(); // Stop retrying
                            }
                        }
                    }
                    catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx) when (
                        odataEx.ResponseStatusCode == 429 || 
                        odataEx.Error?.Code == "activityLimitReached" ||
                        odataEx.Error?.Code == "TooManyRequests")
                    {
                        // Entire batch request was throttled - retry all tasks in this batch
                        throttleCount++;
                        batchRetryCount++;
                        
                        if (batchRetryCount < maxBatchRetries)
                        {
                            // Try to get Retry-After from the ODataError, fallback to exponential backoff
                            var retryAfterSeconds = GetRetryAfterFromODataError(odataEx);
                            batchRetryDelay = retryAfterSeconds > 0
                                ? TimeSpan.FromSeconds(retryAfterSeconds)
                                : TimeSpan.FromSeconds(Math.Pow(2, batchRetryCount)); // Exponential backoff
                            
                            var waitSeconds = (int)batchRetryDelay.TotalSeconds;
                            var delaySource = retryAfterSeconds > 0 ? "Retry-After header" : "exponential backoff";
                            
                            Console.Write($"\r  ⚠️  Batch throttled! Waiting {waitSeconds}s ({delaySource}) before retry {batchRetryCount}/{maxBatchRetries}...                    ");
                            await Task.Delay(batchRetryDelay);
                            Console.Write($"\r  Processing batch {batchNum}/{totalBatches} ({processedCount}/{_completedTasks.Count} tasks)...");
                            // tasksToDelete remains unchanged - will retry all tasks
                        }
                        else
                        {
                            // Max retries - mark all remaining tasks as errors
                            foreach (var (taskId, title) in tasksToDelete)
                            {
                                errors.Add($"Task '{title}': Batch throttled after {maxBatchRetries} retries");
                                errorCount++;
                            }
                            tasksToDelete.Clear(); // Stop retrying
                        }
                    }
                    catch (Exception ex)
                    {
                        // Other error - mark all remaining tasks as failed
                        foreach (var (taskId, title) in tasksToDelete)
                        {
                            errors.Add($"Task '{title}': Batch failed - {ex.Message}");
                            errorCount++;
                        }
                        tasksToDelete.Clear(); // Stop retrying
                    }
                }

                // Add delay between batches to avoid hitting rate limits
                if (batchIndex < totalBatches - 1)
                {
                    await Task.Delay(1000); // 1 second delay between batches
                }
            }
            else
            {
                // Dry run mode - simulate batch processing
                await Task.Delay(100);
                successCount += batchTasks.Count;
            }
        }

        var elapsed = DateTime.UtcNow - startTime;

        // Clear progress line and show final results
        Console.WriteLine($"\r  Processing complete: {totalBatches} batch(es) in {elapsed.TotalSeconds:F1}s                              ");
        Console.WriteLine();

        // Display summary
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"✅ Successfully {(dryRun ? "simulated deletion of" : "deleted")} {successCount} task(s)");
        Console.ResetColor();

        if (skippedCount > 0)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"⚠️  {skippedCount} task(s) skipped (null ID)");
            Console.ResetColor();
        }

        if (errorCount > 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"❌ {errorCount} task(s) encountered errors");
            Console.ResetColor();
            
            if (errors.Count > 0)
            {
                Console.WriteLine("\nErrors:");
                foreach (var error in errors.Take(10))
                {
                    Console.WriteLine($"  - {error}");
                }
                if (errors.Count > 10)
                {
                    Console.WriteLine($"  ... and {errors.Count - 10} more errors");
                }
            }
        }

        if (throttleCount > 0)
        {
            Console.WriteLine($"\n🔄 Handled {throttleCount} throttling event(s) with automatic retry");
        }

        Console.WriteLine($"\nPerformance: {_completedTasks.Count} tasks processed in {totalBatches} batch(es) over {elapsed.TotalSeconds:F1} seconds");
        Console.WriteLine($"             Average: {(_completedTasks.Count / elapsed.TotalSeconds):F1} tasks/second");
        Console.WriteLine();
    }

    /// <summary>
    /// Retry an async operation with throttling-aware exponential backoff
    /// </summary>
    private static async Task<T> RetryWithThrottlingAsync<T>(Func<Task<T>> operation, string operationName, int maxRetries = 5)
    {
        for (int attempt = 0; attempt < maxRetries; attempt++)
        {
            try
            {
                return await operation();
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx) when (
                (odataEx.ResponseStatusCode == 429 || 
                 odataEx.Error?.Code == "activityLimitReached" ||
                 odataEx.Error?.Code == "TooManyRequests") && 
                attempt < maxRetries - 1)
            {
                // Throttled - extract Retry-After and wait
                var retryAfterSeconds = GetRetryAfterFromODataError(odataEx);
                var delay = retryAfterSeconds > 0
                    ? TimeSpan.FromSeconds(retryAfterSeconds)
                    : TimeSpan.FromSeconds(Math.Pow(2, attempt + 1)); // Exponential backoff: 2s, 4s, 8s, 16s, 32s
                
                var waitSeconds = (int)delay.TotalSeconds;
                var delaySource = retryAfterSeconds > 0 ? "Retry-After" : "exponential backoff";
                
                Console.Write($"\r   ⚠️  Throttled during {operationName}. Waiting {waitSeconds}s ({delaySource})... Retry {attempt + 1}/{maxRetries - 1}");
                await Task.Delay(delay);
                Console.Write($"\r   Counting tasks...");
            }
            catch (Exception) when (attempt < maxRetries - 1)
            {
                // Other transient errors - use simple exponential backoff
                await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, attempt)));
            }
        }
        return await operation(); // Final attempt without catch
    }

    /// <summary>
    /// Extract Retry-After header value from HTTP response headers
    /// </summary>
    private static int GetRetryAfterSeconds(System.Net.Http.Headers.HttpResponseHeaders headers)
    {
        if (headers != null && headers.TryGetValues("Retry-After", out var values))
        {
            var retryAfter = values.FirstOrDefault();
            if (!string.IsNullOrEmpty(retryAfter))
            {
                // Retry-After can be either seconds (integer) or HTTP date
                if (int.TryParse(retryAfter, out int seconds))
                {
                    return seconds;
                }
                // If it's a date, parse it and calculate the difference
                if (DateTimeOffset.TryParse(retryAfter, out DateTimeOffset retryDate))
                {
                    var delay = retryDate - DateTimeOffset.UtcNow;
                    return delay.TotalSeconds > 0 ? (int)delay.TotalSeconds : 0;
                }
            }
        }
        return 0; // Return 0 if not found or invalid
    }

    /// <summary>
    /// Extract Retry-After header from batch response for throttled requests
    /// </summary>
    private static async Task<int> GetRetryAfterFromBatchResponseAsync(
        Microsoft.Graph.BatchResponseContentCollection? batchResponse, 
        IEnumerable<string> throttledRequestIds)
    {
        if (batchResponse == null)
            return 0;

        // Check each throttled response for Retry-After header
        foreach (var requestId in throttledRequestIds)
        {
            try
            {
                var response = await batchResponse.GetResponseByIdAsync(requestId);
                if (response?.StatusCode == System.Net.HttpStatusCode.TooManyRequests && response.Headers != null)
                {
                    var retryAfter = GetRetryAfterSeconds(response.Headers);
                    if (retryAfter > 0)
                    {
                        return retryAfter; // Return first valid Retry-After found
                    }
                }
            }
            catch
            {
                // Ignore errors when extracting, continue to next request
                continue;
            }
        }
        return 0; // No Retry-After header found
    }

    /// <summary>
    /// Extract Retry-After header from ODataError (batch-level throttling)
    /// </summary>
    private static int GetRetryAfterFromODataError(Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
    {
        // Try to extract Retry-After from the inner exception or error details
        // The ODataError may contain response headers in some cases
        
        // Check if there's an innerException with ResponseHeaders
        var innerException = odataEx.InnerException as System.Net.Http.HttpRequestException;
        if (innerException != null)
        {
            // Try to find HttpResponseMessage in the exception data or properties
            // This is SDK-version dependent and may not always be available
        }

        // Check the error message for retry-after information
        var errorMessage = odataEx.Error?.Message;
        if (!string.IsNullOrEmpty(errorMessage))
        {
            // Some error messages include "Retry after X seconds"
            var retryMatch = System.Text.RegularExpressions.Regex.Match(
                errorMessage, 
                @"retry after (\d+) second", 
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            
            if (retryMatch.Success && int.TryParse(retryMatch.Groups[1].Value, out int seconds))
            {
                return seconds;
            }
        }

        return 0; // Unable to extract Retry-After from ODataError
    }
}

