using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Helpers
{
    /*
        Extension methods for handling asynchronous operations with proper
        error handling and logging.
        
        These utilities wrap common patterns for executing potentially failing
        operations and converting exceptions into user-friendly error messages.
        The approach ensures consistent error handling across all services.
    */
    public static class ResultExtensions
    {
        /*
            Executes an async operation and captures any exceptions,
            returning a tuple with success status and optional error message.
        */
        public static async Task<(bool Success, string? Error)> TryExecuteAsync(
            this Task task, 
            ILogger logger, 
            string operationName)
        {
            try
            {
                await task;
                return (true, null);
            }
            catch (OperationCanceledException)
            {
                logger.LogWarning("{Operation} was cancelled", operationName);
                return (false, "Operation was cancelled");
            }
            catch (TimeoutException)
            {
                logger.LogWarning("{Operation} timed out", operationName);
                return (false, "Operation timed out");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error during {Operation}", operationName);
                return (false, ex.Message);
            }
        }

        /*
            Executes an async operation that returns a value, capturing
            any exceptions and returning a default value on failure.
        */
        public static async Task<(T? Result, bool Success, string? Error)> TryExecuteAsync<T>(
            this Task<T> task,
            ILogger logger,
            string operationName)
        {
            try
            {
                var result = await task;
                return (result, true, null);
            }
            catch (OperationCanceledException)
            {
                logger.LogWarning("{Operation} was cancelled", operationName);
                return (default, false, "Operation was cancelled");
            }
            catch (TimeoutException)
            {
                logger.LogWarning("{Operation} timed out", operationName);
                return (default, false, "Operation timed out");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error during {Operation}", operationName);
                return (default, false, ex.Message);
            }
        }

        /*
            Retries an async operation up to a specified number of times
            with exponential backoff between attempts.
        */
        public static async Task<T?> RetryAsync<T>(
            this Func<Task<T>> operation,
            int maxAttempts,
            ILogger logger,
            string operationName)
        {
            var delay = TimeSpan.FromMilliseconds(100);

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    return await operation();
                }
                catch (Exception ex) when (attempt < maxAttempts)
                {
                    logger.LogWarning(ex, 
                        "{Operation} failed on attempt {Attempt}/{Max}, retrying in {Delay}ms",
                        operationName, attempt, maxAttempts, delay.TotalMilliseconds);
                    
                    await Task.Delay(delay);
                    delay = TimeSpan.FromMilliseconds(delay.TotalMilliseconds * 2);
                }
            }

            return default;
        }
    }
}
