//Copyright (c) Microsoft Corporation.  All rights reserved.

using Microsoft.WindowsAPICodePack.Controls;
using Microsoft.WindowsAPICodePack.Shell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Microsoft.WindowsAPICodePack.Samples
{
    public partial class ExplorerBrowserTestForm : Form
    {
        private readonly AutoResetEvent itemsChanged = new AutoResetEvent(false);
        private readonly AutoResetEvent selectionChanged = new AutoResetEvent(false);
        private readonly System.Windows.Forms.Timer uiDecoupleTimer = new System.Windows.Forms.Timer();

        public ExplorerBrowserTestForm()
        {
            InitializeComponent();

            // initialize known folder combo box
            knownFolderCombo.Items.AddRange(KnownFolders.All.Select(f => f.CanonicalName).OrderBy(s => s).ToArray());

            // initial property grids
            navigationPropertyGrid.SelectedObject = explorerBrowser.NavigationOptions;
            visibilityPropertyGrid.SelectedObject = explorerBrowser.NavigationOptions.PaneVisibility;
            contentPropertyGrid.SelectedObject = explorerBrowser.ContentOptions;

            // setup ExplorerBrowser navigation events
            explorerBrowser.NavigationPending += new EventHandler<NavigationPendingEventArgs>(explorerBrowser_NavigationPending);
            explorerBrowser.NavigationFailed += new EventHandler<NavigationFailedEventArgs>(explorerBrowser_NavigationFailed);
            explorerBrowser.NavigationComplete += new EventHandler<NavigationCompleteEventArgs>(explorerBrowser_NavigationComplete);
            explorerBrowser.ItemsChanged += new EventHandler(explorerBrowser_ItemsChanged);
            explorerBrowser.SelectionChanged += new EventHandler(explorerBrowser_SelectionChanged);
            explorerBrowser.ViewEnumerationComplete += new EventHandler(explorerBrowser_ViewEnumerationComplete);

            // set up Navigation log event and button state
            explorerBrowser.NavigationLog.NavigationLogChanged += new EventHandler<NavigationLogEventArgs>(NavigationLog_NavigationLogChanged);
            backButton.Enabled = false;
            forwardButton.Enabled = false;

            uiDecoupleTimer.Tick += new EventHandler(uiDecoupleTimer_Tick);
            uiDecoupleTimer.Interval = 100;
            uiDecoupleTimer.Start();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            explorerBrowser.Navigate((ShellObject)KnownFolders.Desktop);
        }

        private void backButton_Click(object sender, EventArgs e) =>
            // Move backwards through navigation log
            explorerBrowser.NavigateLogLocation(NavigationLogDirection.Backward);

        private void clearHistoryButton_Click(object sender, EventArgs e) =>
            // clear navigation log
            explorerBrowser.NavigationLog.ClearLog();

        private void explorerBrowser_ItemsChanged(object sender, EventArgs e) => itemsChanged.Set();

        private void explorerBrowser_NavigationComplete(object sender, NavigationCompleteEventArgs args) =>
            // This event is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            BeginInvoke(new MethodInvoker(delegate ()
            {
                // update event history text box
                var location = (args.NewLocation == null) ? "(unknown)" : args.NewLocation.Name;
                eventHistoryTextBox.Text =
                    eventHistoryTextBox.Text +
                    "Navigation completed. New Location = " + location + "\n";
            }));

        private void explorerBrowser_NavigationFailed(object sender, NavigationFailedEventArgs args) =>
            // This event is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            BeginInvoke(new MethodInvoker(delegate ()
            {
                // update event history text box
                var location = (args.FailedLocation == null) ? "(unknown)" : args.FailedLocation.Name;
                eventHistoryTextBox.Text =
                    eventHistoryTextBox.Text +
                    "Navigation failed. Failed Location = " + location + "\n";

                if (explorerBrowser.NavigationLog.CurrentLocationIndex == -1)
                    navigationHistoryCombo.Text = "";
                else
                    navigationHistoryCombo.SelectedIndex = explorerBrowser.NavigationLog.CurrentLocationIndex;
            }));

        private void explorerBrowser_NavigationPending(object sender, NavigationPendingEventArgs args)
        {
            // fail navigation if check selected (this must be synchronous)
            args.Cancel = failNavigationCheckBox.Checked;

            // This portion is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            BeginInvoke(new MethodInvoker(delegate ()
            {
                // update event history text box
                var message = "";
                var location = (args.PendingLocation == null) ? "(unknown)" : args.PendingLocation.Name;
                if (args.Cancel)
                {
                    message = "Navigation Failing. Pending Location = " + location;
                }
                else
                {
                    message = "Navigation Pending. Pending Location = " + location;
                }
                eventHistoryTextBox.Text =
                    eventHistoryTextBox.Text + message + "\n";
            }));
        }

        private void explorerBrowser_SelectionChanged(object sender, EventArgs e) => selectionChanged.Set();

        private void explorerBrowser_ViewEnumerationComplete(object sender, EventArgs e)
        {
            // This event is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            BeginInvoke(new MethodInvoker(delegate ()
            {
                eventHistoryTextBox.Text =
                    eventHistoryTextBox.Text +
                    "View enumeration complete.\n";
            }));

            selectionChanged.Set();
            itemsChanged.Set();
        }

        private void filePathEdit_TextChanged(object sender, EventArgs e) => filePathNavigate.Enabled = (filePathEdit.Text.Length > 0);

        private void filePathNavigate_Click(object sender, EventArgs e)
        {
            try
            {
                // Navigates to a specified file (must be a container file to work, i.e., ZIP, CAB)
                explorerBrowser.Navigate(ShellFile.FromFilePath(filePathEdit.Text));
            }
            catch (COMException)
            {
                MessageBox.Show("Navigation not possible.");
            }
        }

        private void forwardButton_Click(object sender, EventArgs e) =>
            // Move forwards through navigation log
            explorerBrowser.NavigateLogLocation(NavigationLogDirection.Forward);

        private void knownFolderCombo_SelectedIndexChanged(object sender, EventArgs e) => knownFolderNavigate.Enabled = (knownFolderCombo.Text.Length > 0);

        private void knownFolderNavigate_Click(object sender, EventArgs e)
        {
            try
            {
                // Navigate to a known folder
                var kf =
                    KnownFolderHelper.FromCanonicalName(
                        knownFolderCombo.Items[knownFolderCombo.SelectedIndex].ToString());

                explorerBrowser.Navigate((ShellObject)kf);
            }
            catch (COMException)
            {
                MessageBox.Show("Navigation not possible.");
            }
        }

        private void navigateButton_Click(object sender, EventArgs e)
        {
            try
            {
                // navigate to specific folder
                explorerBrowser.Navigate(ShellFileSystemFolder.FromFolderPath(pathEdit.Text));
            }
            catch (COMException)
            {
                MessageBox.Show("Navigation not possible.");
            }
        }

        private void navigationHistoryCombo_SelectedIndexChanged(object sender, EventArgs e) =>
            // navigating to specific index in navigation log
            explorerBrowser.NavigateLogLocation(navigationHistoryCombo.SelectedIndex);

        private void NavigationLog_NavigationLogChanged(object sender, NavigationLogEventArgs args) =>
            // This event is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            BeginInvoke(new MethodInvoker(delegate ()
            {
                // calculate button states
                if (args.CanNavigateBackwardChanged)
                {
                    backButton.Enabled = explorerBrowser.NavigationLog.CanNavigateBackward;
                }
                if (args.CanNavigateForwardChanged)
                {
                    forwardButton.Enabled = explorerBrowser.NavigationLog.CanNavigateForward;
                }

                // update history combo box
                if (args.LocationsChanged)
                {
                    navigationHistoryCombo.Items.Clear();
                    foreach (var shobj in explorerBrowser.NavigationLog.Locations)
                    {
                        navigationHistoryCombo.Items.Add(shobj.Name);
                    }
                }
                if (explorerBrowser.NavigationLog.CurrentLocationIndex == -1)
                    navigationHistoryCombo.Text = "";
                else
                    navigationHistoryCombo.SelectedIndex = explorerBrowser.NavigationLog.CurrentLocationIndex;
            }));

        private void pathEdit_TextChanged(object sender, EventArgs e) => navigateButton.Enabled = (pathEdit.Text.Length > 0);

        private void uiDecoupleTimer_Tick(object sender, EventArgs e)
        {
            if (selectionChanged.WaitOne(1))
            {
                var itemsText = new StringBuilder();

                foreach (ShellObject item in explorerBrowser.SelectedItems)
                {
                    if (item != null)
                        itemsText.AppendLine("\tItem = " + item.GetDisplayName(DisplayNameType.Default));
                }

                selectedItemsTextBox.Text = itemsText.ToString();
                itemsTabControl.TabPages[1].Text = "Selected Items (Count=" + explorerBrowser.SelectedItems.Count.ToString() + ")";
            }

            if (itemsChanged.WaitOne(1))
            {
                // update items text box
                var itemsText = new StringBuilder();

                foreach (ShellObject item in explorerBrowser.Items)
                {
                    if (item != null)
                        itemsText.AppendLine("\tItem = " + item.GetDisplayName(DisplayNameType.Default));
                }

                itemsTextBox.Text = itemsText.ToString();
                itemsTabControl.TabPages[0].Text = "Items (Count=" + explorerBrowser.Items.Count.ToString() + ")";
            }
        }
    }
}