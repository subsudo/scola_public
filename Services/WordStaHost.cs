using System.Collections.Concurrent;
using System.Diagnostics;
using System.Threading;

namespace VerlaufsakteApp.Services;

public sealed class WordStaHost : IDisposable
{
    private interface IWordStaWorkItem
    {
        string OperationName { get; }
        void Execute(WordService service);
        void Cancel(Exception exception);
    }

    private sealed class WordStaWorkItem<T> : IWordStaWorkItem
    {
        private readonly Func<WordService, T> _action;
        private readonly TaskCompletionSource<T> _completion;

        public WordStaWorkItem(string operationName, Func<WordService, T> action, TaskCompletionSource<T> completion)
        {
            OperationName = operationName;
            _action = action;
            _completion = completion;
        }

        public string OperationName { get; }

        public void Execute(WordService service)
        {
            try
            {
                _completion.TrySetResult(_action(service));
            }
            catch (Exception ex)
            {
                _completion.TrySetException(ex);
            }
        }

        public void Cancel(Exception exception) => _completion.TrySetException(exception);
    }

    private readonly BlockingCollection<IWordStaWorkItem> _queue = new();
    private readonly Thread _workerThread;
    private readonly TaskCompletionSource<bool> _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private volatile bool _isDisposed;
    private volatile Exception? _startupException;

    public WordStaHost()
    {
        _workerThread = new Thread(WorkerLoop)
        {
            IsBackground = true,
            Name = "Scola.WordStaHost"
        };
        _workerThread.SetApartmentState(ApartmentState.STA);
        _workerThread.Start();
    }

    public async Task RunAsync(string operationName, Action<WordService> action)
    {
        await RunAsync<object?>(operationName, service =>
        {
            action(service);
            return null;
        });
    }

    public async Task<T> RunAsync<T>(string operationName, Func<WordService, T> action)
    {
        await _ready.Task.ConfigureAwait(false);

        if (_startupException is not null)
        {
            throw new InvalidOperationException("Word-STA-Host konnte nicht gestartet werden.", _startupException);
        }

        if (_isDisposed)
        {
            throw new ObjectDisposedException(nameof(WordStaHost));
        }

        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var workItem = new WordStaWorkItem<T>(operationName, action, completion);

        try
        {
            AppLogger.Debug($"WordStaHost: queued '{operationName}'.");
            _queue.Add(workItem);
        }
        catch (InvalidOperationException)
        {
            throw new ObjectDisposedException(nameof(WordStaHost));
        }

        return await completion.Task.ConfigureAwait(false);
    }

    private void WorkerLoop()
    {
        WordService? service = null;

        try
        {
            service = new WordService();
            _ready.TrySetResult(true);
            AppLogger.Info("WordStaHost: STA-Worker gestartet.");

            foreach (var workItem in _queue.GetConsumingEnumerable())
            {
                var stopwatch = Stopwatch.StartNew();
                AppLogger.Debug($"WordStaHost: started '{workItem.OperationName}'.");
                workItem.Execute(service);
                stopwatch.Stop();
                AppLogger.Debug($"WordStaHost: completed '{workItem.OperationName}' in {stopwatch.ElapsedMilliseconds} ms.");
            }
        }
        catch (Exception ex)
        {
            _startupException = ex;
            _ready.TrySetException(ex);
            AppLogger.Error("WordStaHost: STA-Worker fehlgeschlagen.", ex);

            while (_queue.TryTake(out var pendingItem))
            {
                pendingItem.Cancel(new InvalidOperationException("Word-STA-Host ist nicht verfügbar.", ex));
            }
        }
        finally
        {
            _queue.Dispose();
            if (service is IDisposable disposable)
            {
                disposable.Dispose();
            }

            AppLogger.Info("WordStaHost: STA-Worker beendet.");
        }
    }

    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }

        _isDisposed = true;
        try
        {
            _queue.CompleteAdding();
        }
        catch (ObjectDisposedException)
        {
            // Der Worker kann die Queue bei einem fruehen Startup-Fehler bereits entsorgt haben.
        }

        if (Thread.CurrentThread != _workerThread && !_workerThread.Join(TimeSpan.FromSeconds(2)))
        {
            AppLogger.Warn("WordStaHost: Worker konnte beim Shutdown nicht rechtzeitig beendet werden.");
        }
    }
}
