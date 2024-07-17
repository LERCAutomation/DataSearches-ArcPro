using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace DataSearches.UI
{
    /// <summary>
    /// Interaction logic for PaneHeader2View.xaml
    /// </summary>
    public partial class PaneHeader2View : UserControl
    {
        public PaneHeader2View()
        {
            InitializeComponent();
        }

        private void ListViewLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<MapLayer> added = e.AddedItems.OfType<MapLayer>().ToList();
            List<MapLayer> removed = e.RemovedItems.OfType<MapLayer>().ToList();

            var listView = sender as ListView;
            var itemsSelected = listView.Items.OfType<MapLayer>().ToList().Where(s => s.IsSelected == true).ToList();
            var itemsUnselected = listView.Items.OfType<MapLayer>().Where(p => p.IsSelected == false).ToList();
            var selectedItems = listView.SelectedItems.OfType<MapLayer>().ToList();

            // Ensure any removed items are actually unselected.
            if (removed.Count > 1)
            {
                e.RemovedItems.OfType<MapLayer>().ToList().ForEach(p => p.IsSelected = false);

                if (selectedItems.Count == 1)
                    listView.Items.OfType<MapLayer>().ToList().Where(s => selectedItems.All(s2 => s2.NodeName != s.NodeName)).ToList().ForEach(p => p.IsSelected = false);
            }
        }
    }
}