using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Autodesk.Revit.UI;

namespace Licorp_CombineCAD.Services
{
    internal sealed class PendingOperation
    {
        public Action<UIApplication> Action { get; }
        public TaskCompletionSource<bool> Tcs { get; }

        public PendingOperation(Action<UIApplication> action)
        {
            Action = action;
            Tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        }
    }

    public class ExportEventHandler : IExternalEventHandler
    {
        private readonly ConcurrentQueue<PendingOperation> _queue;

        internal ExportEventHandler(ConcurrentQueue<PendingOperation> queue)
        {
            _queue = queue;
        }

        public string Name => "Licorp_CombineCAD_Export";

        // Drain ALL queued operations in one Execute call.
        // Revit may coalesce multiple Raise() calls into a single Execute(),
        // so we must process everything available, not just one item.
        public void Execute(UIApplication application)
        {
            while (_queue.TryDequeue(out var op))
            {
                try
                {
                    op.Action(application);
                    op.Tcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[ExportEventHandler] Error: {ex.Message}");
                    op.Tcs.TrySetException(ex);
                }
            }
        }

        public string GetName() => Name;
    }

    public class RevitThreadService
    {
        private readonly ConcurrentQueue<PendingOperation> _queue = new ConcurrentQueue<PendingOperation>();
        private readonly ExternalEvent _externalEvent;
        private readonly ExportEventHandler _handler;

        public RevitThreadService()
        {
            _handler = new ExportEventHandler(_queue);
            _externalEvent = ExternalEvent.Create(_handler);
        }

        /// <summary>Fire-and-forget: enqueues action with no completion signal.</summary>
        public void RunOnRevitThread(Action<UIApplication> action)
        {
            _queue.Enqueue(new PendingOperation(action));
            _externalEvent.Raise();
        }

        public async Task<T> RunOnRevitThreadAsync<T>(Func<UIApplication, T> action, TimeSpan? timeout = null)
        {
            T result = default;
            var op = new PendingOperation(app => { result = action(app); });

            _queue.Enqueue(op);
            _externalEvent.Raise();
            Trace.WriteLine("[RevitThread] External event raised");

            if (timeout.HasValue)
            {
                var completed = await Task.WhenAny(op.Tcs.Task, Task.Delay(timeout.Value)).ConfigureAwait(false);
                if (completed != op.Tcs.Task)
                    throw new TimeoutException($"Revit thread operation timed out after {timeout.Value.TotalSeconds}s.");
            }
            else
            {
                await op.Tcs.Task.ConfigureAwait(false);
            }

            return result;
        }

        public Task RunOnRevitThreadAsync(Action<UIApplication> action, TimeSpan? timeout = null)
        {
            return RunOnRevitThreadAsync<object>(app =>
            {
                action(app);
                return null;
            }, timeout);
        }

        public T ExecuteOnMainThread<T>(Func<T> action)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
                return action();

            T result = default;
            Exception exception = null;
            dispatcher.Invoke(new Action(() =>
            {
                try { result = action(); }
                catch (Exception ex) { exception = ex; }
            }), DispatcherPriority.Normal);
            if (exception != null)
                throw exception;
            return result;
        }

        public void ExecuteOnMainThread(Action action)
        {
            ExecuteOnMainThread<object>(() => { action(); return null; });
        }

        public static void DoEvents()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || !dispatcher.CheckAccess()) return;
            var frame = new DispatcherFrame();
            dispatcher.BeginInvoke(DispatcherPriority.Background,
                new DispatcherOperationCallback(f =>
                {
                    ((DispatcherFrame)f).Continue = false;
                    return null;
                }), frame);
            Dispatcher.PushFrame(frame);
        }

        public void Dispose()
        {
            // Cancel any pending operations before disposing
            while (_queue.TryDequeue(out var op))
                op.Tcs.TrySetException(new ObjectDisposedException(nameof(RevitThreadService)));

            try { _externalEvent?.Dispose(); } catch { }
        }
    }
}
