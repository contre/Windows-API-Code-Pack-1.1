// Copyright (c) Microsoft Corporation. All rights reserved.

using System;
using System.Threading;

namespace Microsoft.WindowsAPICodePack.ExtendedLinguisticServices
{
    /// <summary><see cref="IAsyncResult">IAsyncResult</see> implementation for use with asynchronous calls to ELS.</summary>
    public class MappingAsyncResult : IAsyncResult, IDisposable
    {
        private readonly ManualResetEvent _waitHandle;
        private MappingResultState _resultState;

        internal MappingAsyncResult(
            object callerData,
            AsyncCallback asyncCallback)
        {
            CallerData = callerData;
            AsyncCallback = asyncCallback;
            _waitHandle = new ManualResetEvent(false);
        }

        // returns MappingResultState
        /// <summary>Returns the result state.</summary>
        public object AsyncState => ResultState;

        /// <summary>Gets the WaitHandle which will be notified when the opration completes (successfully or not).</summary>
        public WaitHandle AsyncWaitHandle => _waitHandle;

        /// <summary>Returns the caller data associated with this operation.</summary>
        public object CallerData { get; private set; }

        /// <summary>From MSDN: Most implementers of the IAsyncResult interface will not use this property and should return false.</summary>
        public bool CompletedSynchronously => false;

        /// <summary>Queries whether the operation has completed.</summary>
        public bool IsCompleted
        {
            get
            {
                Thread.MemoryBarrier();
                return AsyncWaitHandle.WaitOne(0, false);
            }
        }

        /// <summary>Gets the resulting <see cref="MappingPropertyBag">MappingPropertyBag</see> (if it exists).</summary>
        public MappingPropertyBag PropertyBag { get; private set; }

        /// <summary>Returns the current result state associated with this operation.</summary>
        public MappingResultState ResultState => _resultState;

        /// <summary>Queries whether the operation completed successfully.</summary>
        public bool Succeeded => PropertyBag != null && _resultState.HResult == 0;

        internal AsyncCallback AsyncCallback { get; private set; }

        /// <summary>Dispose the MappingAsyncresult</summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        internal void SetResult(MappingPropertyBag bag, MappingResultState resultState)
        {
            _resultState = resultState;
            PropertyBag = bag;
        }

        /// <summary>Dispose the MappingAsyncresult</summary>
        protected virtual void Dispose(bool disposed)
        {
            if (disposed)
            {
                _waitHandle.Close();
            }
        }
    }
}