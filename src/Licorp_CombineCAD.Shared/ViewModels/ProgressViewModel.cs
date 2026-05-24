using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
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
        private DispatcherTimer _timer;

        public ProgressViewModel(Action onCancel = null)
        {
            _onCancel = onCancel;
            _stopwatch = new Stopwatch();
            CancelCommand = new RelayCommand(() => ExecuteCancel(), () => !Completed);
        }

        public void StartTimer()
        {
            StopTimer();
            _stopwatch.Restart();

            _timer = new DispatcherTimer(DispatcherPriority.Normal, Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher)
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _timer.Tick += (s, e) =>
            {
                var ts = _stopwatch.Elapsed;
                ElapsedTime = $"{ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2}";
            };
            _timer.Start();
        }

        public void StopTimer()
        {
            _timer?.Stop();
            _timer = null;
            _stopwatch.Stop();
            var ts = _stopwatch.Elapsed;
            ElapsedTime = $"{ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2}";
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
