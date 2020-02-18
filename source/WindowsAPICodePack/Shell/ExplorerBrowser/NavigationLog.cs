﻿//Copyright (c) Microsoft Corporation.  All rights reserved.

using Microsoft.WindowsAPICodePack.Controls.WindowsForms;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.Resources;
using System;
using System.Collections.Generic;

namespace Microsoft.WindowsAPICodePack.Controls
{
    /// <summary>The navigation log is a history of the locations visited by the explorer browser.</summary>
    public class ExplorerBrowserNavigationLog
    {
        private readonly List<ShellObject> _locations = new List<ShellObject>();

        private readonly ExplorerBrowser parent = null;

        /// <summary>The pending navigation log action. null if the user is not navigating via the navigation log.</summary>
        private PendingNavigation pendingNavigation;

        internal ExplorerBrowserNavigationLog(ExplorerBrowser parent)
        {
            if (parent == null)
            {
                throw new ArgumentException(LocalizedMessages.NavigationLogNullParent, "parent");
            }

            // Hook navigation events from the parent to distinguish between navigation log induced navigation, and other navigations.
            this.parent = parent;
            this.parent.NavigationComplete += new EventHandler<NavigationCompleteEventArgs>(OnNavigationComplete);
            this.parent.NavigationFailed += new EventHandler<NavigationFailedEventArgs>(OnNavigationFailed);
        }

        /// <summary>Fires when the navigation log changes or the current navigation position changes</summary>
        public event EventHandler<NavigationLogEventArgs> NavigationLogChanged;

        /// <summary>Indicates the presence of locations in the log that can be reached by calling Navigate(Backward)</summary>
        public bool CanNavigateBackward => CurrentLocationIndex > 0;

        /// <summary>Indicates the presence of locations in the log that can be reached by calling Navigate(Forward)</summary>
        public bool CanNavigateForward => CurrentLocationIndex < _locations.Count - 1;

        /// <summary>Gets the shell object in the Locations collection pointed to by CurrentLocationIndex.</summary>
        public ShellObject CurrentLocation => CurrentLocationIndex < 0 ? null : _locations[CurrentLocationIndex];

        /// <summary>
        /// An index into the Locations collection. The ShellObject pointed to by this index is the current location of the ExplorerBrowser.
        /// </summary>
        public int CurrentLocationIndex { get; private set; } = -1;

        /// <summary>The navigation log</summary>
        public IEnumerable<ShellObject> Locations
        {
            get { foreach (var obj in _locations) { yield return obj; } }
        }

        /// <summary>Clears the contents of the navigation log.</summary>
        public void ClearLog()
        {
            // nothing to do
            if (_locations.Count == 0) { return; }

            var oldCanNavigateBackward = CanNavigateBackward;
            var oldCanNavigateForward = CanNavigateForward;

            _locations.Clear();
            CurrentLocationIndex = -1;

            var args = new NavigationLogEventArgs
            {
                LocationsChanged = true,
                CanNavigateBackwardChanged = oldCanNavigateBackward != CanNavigateBackward,
                CanNavigateForwardChanged = oldCanNavigateForward != CanNavigateForward
            };
            NavigationLogChanged?.Invoke(this, args);
        }

        internal bool NavigateLog(NavigationLogDirection direction)
        {
            // determine proper index to navigate to
            int locationIndex;
            if (direction == NavigationLogDirection.Backward && CanNavigateBackward)
            {
                locationIndex = CurrentLocationIndex - 1;
            }
            else if (direction == NavigationLogDirection.Forward && CanNavigateForward)
            {
                locationIndex = CurrentLocationIndex + 1;
            }
            else
            {
                return false;
            }

            // initiate traversal request
            var location = _locations[locationIndex];
            pendingNavigation = new PendingNavigation(location, locationIndex);
            parent.Navigate(location);
            return true;
        }

        internal bool NavigateLog(int index)
        {
            // can't go anywhere
            if (index >= _locations.Count || index < 0) { return false; }

            // no need to re navigate to the same location
            if (index == CurrentLocationIndex) { return false; }

            // initiate traversal request
            var location = _locations[index];
            pendingNavigation = new PendingNavigation(location, index);
            parent.Navigate(location);
            return true;
        }

        private void OnNavigationComplete(object sender, NavigationCompleteEventArgs args)
        {
            var eventArgs = new NavigationLogEventArgs();
            var oldCanNavigateBackward = CanNavigateBackward;
            var oldCanNavigateForward = CanNavigateForward;

            if (pendingNavigation != null)
            {
                // navigation log traversal in progress

                // determine if new location is the same as the traversal request
                pendingNavigation.Location.NativeShellItem.Compare(
                    args.NewLocation.NativeShellItem, SICHINTF.SICHINT_ALLFIELDS, out var result);
                var shellItemsEqual = result == 0;
                if (shellItemsEqual == false)
                {
                    // new location is different than traversal request, behave is if it never happened! remove history following
                    // currentLocationIndex, append new item
                    if (CurrentLocationIndex < _locations.Count - 1)
                    {
                        _locations.RemoveRange(CurrentLocationIndex + 1, _locations.Count - (CurrentLocationIndex + 1));
                    }
                    _locations.Add(args.NewLocation);
                    CurrentLocationIndex = _locations.Count - 1;
                    eventArgs.LocationsChanged = true;
                }
                else
                {
                    // log traversal successful, update index
                    CurrentLocationIndex = pendingNavigation.Index;
                    eventArgs.LocationsChanged = false;
                }
                pendingNavigation = null;
            }
            else
            {
                // remove history following currentLocationIndex, append new item
                if (CurrentLocationIndex < _locations.Count - 1)
                {
                    _locations.RemoveRange(CurrentLocationIndex + 1, _locations.Count - (CurrentLocationIndex + 1));
                }
                _locations.Add(args.NewLocation);
                CurrentLocationIndex = _locations.Count - 1;
                eventArgs.LocationsChanged = true;
            }

            // update event args
            eventArgs.CanNavigateBackwardChanged = oldCanNavigateBackward != CanNavigateBackward;
            eventArgs.CanNavigateForwardChanged = oldCanNavigateForward != CanNavigateForward;

            NavigationLogChanged?.Invoke(this, eventArgs);
        }

        private void OnNavigationFailed(object sender, NavigationFailedEventArgs args) => pendingNavigation = null;
    }

    /// <summary>A navigation traversal request</summary>
    internal class PendingNavigation
    {
        internal PendingNavigation(ShellObject location, int index)
        {
            Location = location;
            Index = index;
        }

        internal int Index { get; set; }
        internal ShellObject Location { get; set; }
    }
}