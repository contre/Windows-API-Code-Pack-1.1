//Copyright (c) Microsoft Corporation.  All rights reserved.

using Microsoft.WindowsAPICodePack.Shell;
using System;
using System.Collections.ObjectModel;
using System.Text;

namespace Microsoft.WindowsAPICodePack.Dialogs
{
    /// <summary>Stores the file extensions used when filtering files in File Open and File Save dialogs.</summary>
    public class CommonFileDialogFilter
    {
        private string rawDisplayName;

        /// <summary>Creates a new instance of this class.</summary>
        public CommonFileDialogFilter() => Extensions = new Collection<string>();

        /// <summary>Creates a new instance of this class with the specified display name and file extension list.</summary>
        /// <param name="rawDisplayName">The name of this filter.</param>
        /// <param name="extensionList">The list of extensions in this filter. See remarks.</param>
        /// <remarks>
        /// The <paramref name="extensionList"/> can use a semicolon(";") or comma (",") to separate extensions. Extensions can be prefaced
        /// with a period (".") or with the file wild card specifier "*.".
        /// </remarks>
        /// <permission cref="ArgumentNullException">The <paramref name="extensionList"/> cannot be null or a zero-length string.</permission>
        public CommonFileDialogFilter(string rawDisplayName, string extensionList)
            : this()
        {
            if (string.IsNullOrEmpty(extensionList))
            {
                throw new ArgumentNullException(nameof(extensionList));
            }

            this.rawDisplayName = rawDisplayName;

            // Parse string and create extension strings.
            // Format: "bat,cmd", or "bat;cmd", or "*.bat;*.cmd" Can support leading "." or "*." - these will be stripped.
            var rawExtensions = extensionList.Split(',', ';');
            foreach (var extension in rawExtensions)
            {
                Extensions.Add(CommonFileDialogFilter.NormalizeExtension(extension));
            }
        }

        /// <summary>Gets or sets the display name for this filter.</summary>
        /// <permission cref="ArgumentNullException">The value for this property cannot be set to null or a zero-length string.</permission>
        public string DisplayName
        {
            get => ShowExtensions
                    ? string.Format(System.Globalization.CultureInfo.InvariantCulture,
                        "{0} ({1})",
                        rawDisplayName,
                        CommonFileDialogFilter.GetDisplayExtensionList(Extensions))
                    : rawDisplayName;

            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new ArgumentNullException(nameof(value));
                }
                rawDisplayName = value;
            }
        }

        /// <summary>Gets a collection of the individual extensions described by this filter.</summary>
        public Collection<string> Extensions { get; private set; }

        /// <summary>Gets or sets a value that controls whether the extensions are displayed.</summary>
        public bool ShowExtensions { get; set; } = true;

        /// <summary>Returns a string representation for this filter that includes the display name and the list of extensions.</summary>
        /// <returns>A <see cref="string"/>.</returns>
        public override string ToString() => string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "{0} ({1})",
                rawDisplayName,
                CommonFileDialogFilter.GetDisplayExtensionList(Extensions));

        /// <summary>Internal helper that generates a single filter specification for this filter, used by the COM API.</summary>
        /// <returns>Filter specification for this filter</returns>
        internal ShellNativeMethods.FilterSpec GetFilterSpec()
        {
            var filterList = new StringBuilder();
            foreach (var extension in Extensions)
            {
                if (filterList.Length > 0) { filterList.Append(";"); }

                filterList.Append("*.");
                filterList.Append(extension);
            }
            return new ShellNativeMethods.FilterSpec(DisplayName, filterList.ToString());
        }

        private static string GetDisplayExtensionList(Collection<string> extensions)
        {
            var extensionList = new StringBuilder();
            foreach (var extension in extensions)
            {
                if (extensionList.Length > 0) { extensionList.Append(", "); }
                extensionList.Append("*.");
                extensionList.Append(extension);
            }

            return extensionList.ToString();
        }

        private static string NormalizeExtension(string rawExtension)
        {
            rawExtension = rawExtension.Trim();
            rawExtension = rawExtension.Replace("*.", null);
            rawExtension = rawExtension.Replace(".", null);
            return rawExtension;
        }
    }
}