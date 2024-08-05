﻿using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

public static class Debouncer
{
    private static readonly ConcurrentDictionary<object, CancellationTokenSource> CancellationTokenSources = new ConcurrentDictionary<object, CancellationTokenSource>();

    public static async Task Debounce(object key, int delay, Func<Task> action)
    {
        if (CancellationTokenSources.TryGetValue(key, out var oldCts))
        {
            oldCts.Cancel();
        }

        var cts = new CancellationTokenSource();
        CancellationTokenSources[key] = cts;

        try
        {
            await Task.Delay(delay, cts.Token);
            if (!cts.Token.IsCancellationRequested)
            {
                await action();
            }
        }
        catch (TaskCanceledException)
        {
            // Task was canceled, do nothing
        }
        finally
        {
            //CancellationTokenSources.TryRemove(key, out _);
        }
    }
}