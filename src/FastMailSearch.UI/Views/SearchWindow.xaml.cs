using System;
using System.Windows;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.UI.ViewModels;

namespace FastMailSearch.UI.Views
{
    public partial class SearchWindow : Window
    {
        public SearchWindow()
        {
            Logger.LogStage(15, "SearchWindow constructor entered");
            Logger.LogStage(16, "InitializeComponent entered");
            InitializeComponent();
            Logger.LogStage(17, "InitializeComponent completed");

            Loaded += SearchWindow_Loaded;
            Closed += SearchWindow_Closed;
        }

        public SearchWindow(SearchViewModel viewModel) : this()
        {
            DataContext = viewModel;
        }

        private void SearchWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Logger.LogStage(21, "Window Loaded event");
        }

        private void SearchWindow_Closed(object? sender, EventArgs e)
        {
            Logger.LogStage(26, "Window Closed event");
            (DataContext as SearchViewModel)?.Cleanup();
        }

        private void OpenIndexHealth_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is SearchViewModel vm)
            {
                var healthWindow = new IndexHealthWindow(vm.HealthService, vm.Repository, vm.IndexerService)
                {
                    Owner = this
                };
                healthWindow.ShowDialog();
            }
        }

        private void OpenDiagnostics_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is SearchViewModel vm)
            {
                var diagWindow = new DiagnosticsWindow(vm.HealthService)
                {
                    Owner = this
                };
                diagWindow.ShowDialog();
            }
            else
            {
                var diagWindow = new DiagnosticsWindow()
                {
                    Owner = this
                };
                diagWindow.ShowDialog();
            }
        }
    }
}
