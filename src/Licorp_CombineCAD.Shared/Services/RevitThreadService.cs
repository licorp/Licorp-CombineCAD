using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Autodesk.Revit.DB;
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
                Debug.WriteLine($"[ExportEventHandler] Error: {ex.Message}");
                Trace.WriteLine($"[ExportEventHandler] Error: {ex.Message}");
            }
        }

        public string GetName()
        {
            return Name;
        }
    }

    public class RevitThreadService
    {
        private ExternalEvent _externalEvent;
        private ExportEventHandler _handler;

        public RevitThreadService()
        {
            _handler = new ExportEventHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public void RunOnRevitThread(Action<UIApplication> action)
        {
            _handler.ExecuteAction = action;
            _externalEvent.Raise();
            Debug.WriteLine("[RevitThread] Raised external event");
            Trace.WriteLine("[RevitThread] Raised external event");
        }

        public Task<T> RunOnRevitThreadAsync<T>(Func<UIApplication, T> action)
        {
            var tcs = new TaskCompletionSource<T>();
            _handler.ExecuteAction = (app) =>
            {
                try
                {
                    var result = action(app);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            };
            _externalEvent.Raise();
            Debug.WriteLine("[RevitThread] Raised external event (async with result)");
            Trace.WriteLine("[RevitThread] Raised external event (async with result)");
            return tcs.Task;
        }

        public Task RunOnRevitThreadAsync(Action<UIApplication> action)
        {
            return RunOnRevitThreadAsync<object>(app =>
            {
                action(app);
                return null;
            });
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
        }
    }
}
