//Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Globalization;

namespace Microsoft.WindowsAPICodePack.Dialogs
{
    /// <summary>Represents a command-link.</summary>
    public class TaskDialogCommandLink : TaskDialogButton
    {

        /// <summary>Creates a new instance of this class.</summary>
        public TaskDialogCommandLink() { }

        /// <summary>Creates a new instance of this class with the specified name and label.</summary>
        /// <param name="name">The name for this button.</param>
        /// <param name="text">The label for this button.</param>
        public TaskDialogCommandLink(string name, string text) : base(name, text) { }

        /// <summary>Creates a new instance of this class with the specified name,label, and instruction.</summary>
        /// <param name="name">The name for this button.</param>
        /// <param name="text">The label for this button.</param>
        /// <param name="instruction">The instruction for this command link.</param>
        public TaskDialogCommandLink(string name, string text, string instruction)
            : base(name, text) => Instruction = instruction;

        /// <summary>Gets or sets the instruction associated with this command link button.</summary>
        public string Instruction { get; set; }

        /// <summary>Returns a string representation of this object.</summary>
        /// <returns>A <see cref="string"/></returns>
        public override string ToString() => string.Format(CultureInfo.CurrentCulture, "{0}{1}{2}",
                Text ?? string.Empty,
                (!string.IsNullOrEmpty(Text) && !string.IsNullOrEmpty(Instruction)) ? Environment.NewLine : string.Empty,
                Instruction ?? string.Empty);
    }
}