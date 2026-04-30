using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Licorp_CombineCAD.Services;
using Licorp_CombineCAD.ViewModels;

namespace Licorp_CombineCAD.Views
{
    public partial class ReorderSheetsDialog : Window
    {
        private readonly ObservableCollection<ReorderItem> _items = new ObservableCollection<ReorderItem>();
        private Point _dragStart;
        private bool _isDragging;

        public ReorderSheetsDialog(IEnumerable<SheetItemViewModel> sheets)
        {
            InitializeComponent();

            var index = 1;
            foreach (var sheet in sheets ?? Enumerable.Empty<SheetItemViewModel>())
            {
                _items.Add(new ReorderItem
                {
                    Index = index++,
                    IdValue = ViewSheetSetService.GetElementIdValue(sheet.ElementId),
                    DisplayName = sheet.DisplayText
                });
            }

            ItemsListBox.ItemsSource = _items;
        }

        public List<string> OrderedIds { get; private set; } = new List<string>();

        private void ItemsListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStart = e.GetPosition(ItemsListBox);
            _isDragging = false;
        }

        private void ItemsListBox_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _isDragging)
                return;

            var diff = _dragStart - e.GetPosition(ItemsListBox);
            if (System.Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
                System.Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            var item = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);
            if (item == null)
                return;

            var data = item.DataContext as ReorderItem;
            if (data == null)
                return;

            _isDragging = true;
            DragDrop.DoDragDrop(item, data, DragDropEffects.Move);
            _isDragging = false;
        }

        private void ItemsListBox_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(typeof(ReorderItem))
                ? DragDropEffects.Move
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void ItemsListBox_Drop(object sender, DragEventArgs e)
        {
            var dropped = e.Data.GetData(typeof(ReorderItem)) as ReorderItem;
            var targetContainer = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);
            var target = targetContainer?.DataContext as ReorderItem;

            if (dropped == null || target == null || dropped == target)
                return;

            var oldIndex = _items.IndexOf(dropped);
            var newIndex = _items.IndexOf(target);
            if (oldIndex < 0 || newIndex < 0)
                return;

            _items.Move(oldIndex, newIndex);
            UpdateIndices();
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            MoveSelected(-1);
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            MoveSelected(1);
        }

        private void MoveSelected(int delta)
        {
            var item = ItemsListBox.SelectedItem as ReorderItem;
            if (item == null)
                return;

            var oldIndex = _items.IndexOf(item);
            var newIndex = oldIndex + delta;
            if (oldIndex < 0 || newIndex < 0 || newIndex >= _items.Count)
                return;

            _items.Move(oldIndex, newIndex);
            ItemsListBox.SelectedItem = item;
            UpdateIndices();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            OrderedIds = _items.Select(i => i.IdValue).ToList();
            DialogResult = true;
        }

        private void UpdateIndices()
        {
            for (int i = 0; i < _items.Count; i++)
                _items[i].Index = i + 1;
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T match)
                    return match;

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        private class ReorderItem : INotifyPropertyChanged
        {
            private int _index;

            public int Index
            {
                get => _index;
                set { _index = value; OnPropertyChanged(); }
            }

            public string IdValue { get; set; }
            public string DisplayName { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged([CallerMemberName] string propertyName = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
