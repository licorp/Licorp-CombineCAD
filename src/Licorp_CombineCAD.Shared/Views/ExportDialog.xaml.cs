using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;
using Autodesk.Revit.UI;
using Licorp_CombineCAD.Models;
using Licorp_CombineCAD.ViewModels;

namespace Licorp_CombineCAD.Views
{
    public partial class ExportDialog : Window
    {
        private bool _isDragging;
        private bool _isBulkUpdatingCheckboxes;
        private Point _dragStartPoint;
        private SheetItemViewModel _dragSourceItem;

        public ExportDialog(UIDocument uiDocument, ExportMode preselectedMode = ExportMode.MultiLayout)
        {
            InitializeComponent();
            var viewModel = new ExportDialogViewModel(uiDocument);
            viewModel.ExportMode = preselectedMode;
            DataContext = viewModel;
            Loaded += ExportDialog_Loaded;
            Closing += ExportDialog_Closing;
        }

        private async void ExportDialog_Loaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is ExportDialogViewModel vm)
                await vm.InitializeAsync();
        }

        private void ExportDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (DataContext is ExportDialogViewModel vm)
                vm.Dispose();
        }

        private void SheetDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var grid = sender as DataGrid;
            if (grid == null) return;

            var row = GetDataGridRow(grid, e);
            if (row == null) return;

            _dragStartPoint = e.GetPosition(grid);
            _dragSourceItem = row.Item as SheetItemViewModel;
            _isDragging = false;
        }

        private void SheetDataGrid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_dragSourceItem == null || e.LeftButton != MouseButtonState.Pressed) return;

            var grid = sender as DataGrid;
            if (grid == null) return;

            var currentPos = e.GetPosition(grid);
            var diff = _dragStartPoint - currentPos;

            if (!_isDragging && (System.Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                                 System.Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                _isDragging = true;
                DragDrop.DoDragDrop(grid, _dragSourceItem, DragDropEffects.Move);
                _isDragging = false;
                _dragSourceItem = null;
            }
        }

        private void SheetDataGrid_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(SheetItemViewModel)))
            {
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
        }

        private void SheetDataGrid_Drop(object sender, DragEventArgs e)
        {
            if (!(e.Data.GetDataPresent(typeof(SheetItemViewModel)))) return;

            var grid = sender as DataGrid;
            if (grid == null) return;

            var droppedItem = e.Data.GetData(typeof(SheetItemViewModel)) as SheetItemViewModel;
            if (droppedItem == null) return;

            var targetRow = GetDataGridRowFromPoint(grid, e.GetPosition(grid));
            if (targetRow == null) return;

            var targetItem = targetRow.Item as SheetItemViewModel;
            if (targetItem == null || targetItem == droppedItem) return;

            if (DataContext is ExportDialogViewModel vm)
            {
                vm.ReorderSheet(droppedItem, targetItem);
            }

            e.Handled = true;
        }

        private void SheetSelectionCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (_isBulkUpdatingCheckboxes)
                return;

            var checkbox = sender as CheckBox;
            var clickedSheet = checkbox?.DataContext as SheetItemViewModel;
            if (checkbox == null || clickedSheet == null || SheetDataGrid?.SelectedItems == null)
                return;

            if (SheetDataGrid.SelectedItems.Count <= 1 || !SheetDataGrid.SelectedItems.Contains(clickedSheet))
                return;

            _isBulkUpdatingCheckboxes = true;
            try
            {
                var newState = checkbox.IsChecked == true;
                foreach (var item in SheetDataGrid.SelectedItems.OfType<SheetItemViewModel>())
                    item.IsSelected = newState;
            }
            finally
            {
                _isBulkUpdatingCheckboxes = false;
            }
        }

        private DataGridRow GetDataGridRow(DataGrid grid, MouseButtonEventArgs e)
        {
            var element = grid.InputHitTest(e.GetPosition(grid)) as System.Windows.Media.Visual;
            while (element != null && !(element is DataGridRow))
            {
                element = System.Windows.Media.VisualTreeHelper.GetParent(element) as System.Windows.Media.Visual;
            }
            return element as DataGridRow;
        }

        private DataGridRow GetDataGridRowFromPoint(DataGrid grid, Point point)
        {
            var element = grid.InputHitTest(point) as System.Windows.Media.Visual;
            while (element != null && !(element is DataGridRow))
            {
                element = System.Windows.Media.VisualTreeHelper.GetParent(element) as System.Windows.Media.Visual;
            }
            return element as DataGridRow;
        }
    }
}
