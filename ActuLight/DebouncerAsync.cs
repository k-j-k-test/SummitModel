using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace ActuLight
{
    public static class ThrottlerAsync
    {
        private static readonly ConcurrentDictionary<object, DateTime> _lastExecutionTimes = new ConcurrentDictionary<object, DateTime>();

        public static bool ShouldExecute(object key, int throttleMs)
        {
            var now = DateTime.UtcNow;
            return _lastExecutionTimes.AddOrUpdate(key, now, (_, last) =>
            {
                if ((now - last).TotalMilliseconds >= throttleMs)
                {
                    return now;
                }
                return last;
            }) == now;
        }
    }

    public static class DebouncerAsync
    {
        private static readonly ConcurrentDictionary<object, CancellationTokenSource> _cancellationTokenSources = new ConcurrentDictionary<object, CancellationTokenSource>();

        public static async Task Debounce(object key, int delay, Func<Task> action)
        {
            var cts = _cancellationTokenSources.AddOrUpdate(key,
                _ => new CancellationTokenSource(),
                (_, old) =>
                {
                    old.Cancel();
                    return new CancellationTokenSource();
                });

            try
            {               
                await Task.Delay(delay, cts.Token);
                await action();
            }
            catch (TaskCanceledException)
            {
                // Task was canceled, do nothing
            }
            finally
            {
                _cancellationTokenSources.TryRemove(key, out _);
            }
        }
    }
}

