using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Licorp_CombineCAD.Services;

namespace Licorp_CombineCAD.ViewModels
{
    public class ProgressViewModel : INotifyPropertyChanged
    {
        private string _phase = "Initializing";
        private string _currentItem = "";
        private int _current = 0;
        private int _total = 0;
        private double _percentage = 0;
        private string _progressText = "";
        private bool _completed = false;
        private bool _isCancelled = false;
        private string _elapsedTime = "00:00:00";
        private readonly Action _onCancel;
        private readonly Stopwatch _stopwatch;
        private CancellationTokenSource _timerCts;
        private Task _timerTask;
        private string _timeRemaining = "Calculating...";

        public ProgressViewModel(Action onCancel = null)
        {
            _onCancel = onCancel;
            _stopwatch = new Stopwatch();
            CancelCommand = new RelayCommand(() => ExecuteCancel(), () => !Completed);
        }

        public void StartTimer()
        {
            _stopwatch.Restart();
            _timerCts = new CancellationTokenSource();
            _timerTask = Task.Run(async () =>
            {
                while (!_timerCts.Token.IsCancellationRequested)
                {
                    await Task.Delay(1000, _timerCts.Token);
                    if (_timerCts.Token.IsCancellationRequested) break;
                    
                    var ts = _stopwatch.Elapsed;
                    var timeStr = $"{ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2}";
                    
                    string remainingStr = "Calculating...";
                    if (_percentage > 0)
                    {
                        var elapsedMs = _stopwatch.ElapsedMilliseconds;
                        var totalEstimatedMs = elapsedMs / (_percentage / 100.0);
                        var remainingMs = totalEstimatedMs - elapsedMs;
                        if (remainingMs > 0)
                        {
                            var tsRemaining = TimeSpan.FromMilliseconds(remainingMs);
                            remainingStr = $"{tsRemaining.Hours:D2}:{tsRemaining.Minutes:D2}:{tsRemaining.Seconds:D2}";
                        }
                        else
                        {
                            remainingStr = "00:00:00";
                        }
                    }

                    Application.Current?.Dispatcher?.InvokeAsync(
                        () => 
                        {
                            ElapsedTime = timeStr;
                            TimeRemaining = remainingStr;
                        },
                        DispatcherPriority.Normal);
                }
            }, _timerCts.Token);
            
            UpdateElapsedTime();
        }

        public void StopTimer()
        {
            _stopwatch.Stop();
            _timerCts?.Cancel();
            UpdateElapsedTime();
            TimeRemaining = "00:00:00";
        }

        public string TimeRemaining
        {
            get => _timeRemaining;
            set { _timeRemaining = value; OnPropertyChanged(); }
        }

        public string Phase
        {
            get => _phase;
            set { _phase = value; OnPropertyChanged(); }
        }
        public string CurrentItem
        {
            get => _currentItem;
            set { _currentItem = value; OnPropertyChanged(); }
        }

        public int Current
        {
            get => _current;
            set { _current = value; OnPropertyChanged(); UpdatePercentage(); }
        }

        public int Total
        {
            get => _total;
            set { _total = value; OnPropertyChanged(); UpdatePercentage(); }
        }

        public double Percentage
        {
            get => _percentage;
            set { _percentage = value; OnPropertyChanged(); }
        }

        public string ProgressText
        {
            get => _progressText;
            set { _progressText = value; OnPropertyChanged(); }
        }

        public bool Completed
        {
            get => _completed;
            set { _completed = value; OnPropertyChanged(); }
        }

        public bool IsCancelled
        {
            get => _isCancelled;
            set { _isCancelled = value; OnPropertyChanged(); }
        }

        public string ElapsedTime
        {
            get => _elapsedTime;
            set { _elapsedTime = value; OnPropertyChanged(); }
        }

        public ICommand CancelCommand { get; }

        private void ExecuteCancel()
        {
            IsCancelled = true;
            _onCancel?.Invoke();
        }

        private void UpdateElapsedTime()
        {
            var ts = _stopwatch.Elapsed;
            ElapsedTime = $"{ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2}";
        }

        public void Update(string phase, string currentItem, int current, int total)
        {
            Phase = phase;
            CurrentItem = currentItem;
            Current = current;
            Total = total;
            var percentage = total > 0 ? (double)current / total * 100 : 0;
            ProgressText = $"{current}/{total} ({percentage:F0}%)";
        }

        public void UpdatePhase(string phase)
        {
            Phase = phase;
            Current = 0;
            Total = 0;
        }

        private void UpdatePercentage()
        {
            Percentage = Total > 0 ? (double)Current / Total * 100 : 0;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
