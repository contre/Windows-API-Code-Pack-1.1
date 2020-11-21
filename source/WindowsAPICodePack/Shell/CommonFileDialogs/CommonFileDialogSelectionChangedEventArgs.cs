using System;

namespace Microsoft.WindowsAPICodePack.Shell
{
    /// <summary>
    /// Custom event arguments.
    /// </summary>
    public class CommonFileDialogSelectionChangedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CommonFileDialogSelectionChangedEventArgs"/> class.
        /// </summary>
        public CommonFileDialogSelectionChangedEventArgs(string folder, string fileName)
        {
            Folder = folder;
            FileName = fileName;
        }



        /// <summary>
        /// Gets the name of the selected folder.
        /// </summary>
        /// <value>
        /// The name of the selected folder.
        /// </value>
        public string Folder { get; }



        /// <summary>
        /// Gets the name of the selected file.
        /// </summary>
        /// <value>
        /// The name of the selected file.
        /// </value>
        public string FileName { get; }
    }
}
