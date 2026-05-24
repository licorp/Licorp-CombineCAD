using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Autodesk.Revit.UI;

namespace Licorp_CombineCAD.Services
{
    public class ExportEventHandler : IExternalEventHandler
    {
        public Action<UIApplication> ExecuteAction { get; set; }
        public string Name => "Licorp_CombineCAD_Export";

        public void Execute(UIApplication application)
        {
            try
            {
                ExecuteAction?.Invoke(application);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[ExportEventHandler] Error: {ex.Message}");
                Trace.WriteLine($"[ExportEventHandler] Stack: {ex.StackTrace}");
            }
        }

        public string GetName() => Name;
    }

    public class RevitThreadService
    {
        private readonly ExternalEvent _externalEvent;
        private readonly ExportEventHandler _handler;
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

        public RevitThreadService()
        {
            _handler = new ExportEventHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public void RunOnRevitThread(Action<UIApplication> action)
        {
            _handler.ExecuteAction = action;
            _externalEvent.Raise();
        }

        public async Task<T> RunOnRevitThreadAsync<T>(Func<UIApplication, T> action, TimeSpan? timeout = null)
        {
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

            await _semaphore.WaitAsync().ConfigureAwait(false);
            try
            {
                _handler.ExecuteAction = app =>
                {
                    try
                    {
                        tcs.SetResult(action(app));
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                };
                _externalEvent.Raise();
                Trace.WriteLine("[RevitThread] External event raised");
            }
            finally
            {
                _semaphore.Release();
            }

            // Await without capturing the UI SynchronizationContext so the UI thread
            // stays free to process Revit's ExternalEvent dispatch on all versions.
            if (timeout.HasValue)
            {
                var completed = await Task.WhenAny(tcs.Task, Task.Delay(timeout.Value)).ConfigureAwait(false);
                if (completed != tcs.Task)
                    throw new TimeoutException($"Revit thread operation timed out after {timeout.Value.TotalSeconds}s.");
            }

            return await tcs.Task.ConfigureAwait(false);
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
            try { _externalEvent?.Dispose(); } catch { }
            _semaphore?.Dispose();
        }
    }
}
