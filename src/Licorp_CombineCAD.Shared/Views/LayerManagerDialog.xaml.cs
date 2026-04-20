using System.Windows;
using Autodesk.Revit.UI;
using Licorp_CombineCAD.ViewModels;

namespace Licorp_CombineCAD.Views
{
    public partial class LayerManagerDialog : Window
    {
        public LayerManagerDialog(UIDocument uiDocument)
        {
            InitializeComponent();
            DataContext = new LayerManagerViewModel(uiDocument);
        }
    }
}