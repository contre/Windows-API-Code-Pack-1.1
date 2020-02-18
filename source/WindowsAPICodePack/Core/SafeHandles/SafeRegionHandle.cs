﻿//Copyright (c) Microsoft Corporation.  All rights reserved.

namespace MS.WindowsAPICodePack.Internal
{
    /// <summary>Safe Region Handle</summary>
    public class SafeRegionHandle : ZeroInvalidHandle
    {
        /// <summary>Release the handle</summary>
        /// <returns>true if handled is release successfully, false otherwise</returns>
        protected override bool ReleaseHandle() => CoreNativeMethods.DeleteObject(handle);
    }
}