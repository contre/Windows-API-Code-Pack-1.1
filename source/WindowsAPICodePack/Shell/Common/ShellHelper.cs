//Copyright (c) Microsoft Corporation.  All rights reserved.

using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using Microsoft.WindowsAPICodePack.Shell.Resources;
using MS.WindowsAPICodePack.Internal;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Microsoft.WindowsAPICodePack.Shell
{
    /// <summary>A helper class for Shell Objects</summary>
    internal static class ShellHelper
    {
        internal static PropertyKey ItemTypePropertyKey = new PropertyKey(new Guid("28636AA6-953D-11D2-B5D6-00C04FD918D0"), 11);

        internal static string GetAbsolutePath(string path) => Uri.IsWellFormedUriString(path, UriKind.Absolute) ? path : Path.GetFullPath(path);

        internal static string GetItemType(IShellItem2 shellItem) => shellItem != null && shellItem.GetString(ref ItemTypePropertyKey, out var itemType) == HResult.Ok ? itemType : null;

        internal static string GetParsingName(IShellItem shellItem)
        {
            if (shellItem == null) { return null; }

            string path = null;

            var hr = shellItem.GetDisplayName(ShellNativeMethods.ShellItemDesignNameOptions.DesktopAbsoluteParsing, out var pszPath);

            if (hr != HResult.Ok && hr != HResult.InvalidArguments)
            {
                throw new ShellException(LocalizedMessages.ShellHelperGetParsingNameFailed, hr);
            }

            if (pszPath != IntPtr.Zero)
            {
                path = Marshal.PtrToStringAuto(pszPath);
                Marshal.FreeCoTaskMem(pszPath);
            }

            return path;
        }

        internal static IntPtr PidlFromParsingName(string name) => 
            CoreErrorHelper.Succeeded(ShellNativeMethods.SHParseDisplayName(name, IntPtr.Zero, out var pidl, 0, out _)) ? pidl : IntPtr.Zero;

        internal static IntPtr PidlFromShellItem(IShellItem nativeShellItem) => PidlFromUnknown(Marshal.GetIUnknownForObject(nativeShellItem));

        internal static IntPtr PidlFromUnknown(IntPtr unknown) => CoreErrorHelper.Succeeded(ShellNativeMethods.SHGetIDListFromObject(unknown, out var pidl)) ? pidl : IntPtr.Zero;
    }
}