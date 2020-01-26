// Copyright (c) Microsoft Corporation. All rights reserved.

using System;

namespace Microsoft.WindowsAPICodePack.Sensors
{
	/// <summary>Represents all the data from a single sensor data report.</summary>
	public class SensorReport
	{

		/// <summary>Gets the sensor that is the source of this data report.</summary>
		public Sensor Source { get; private set; }

		/// <summary>Gets the time when the data report was generated.</summary>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "TimeStamp")]
		public DateTime TimeStamp { get; private set; } = new DateTime();

		/// <summary>Gets the data values in the report.</summary>
		public SensorData Values { get; private set; }

		internal static SensorReport FromNativeReport(Sensor originator, ISensorDataReport iReport)
		{
			iReport.GetTimestamp(out var systemTimeStamp);
			SensorNativeMethods.SystemTimeToFileTime(ref systemTimeStamp, out var ftTimeStamp);
			var lTimeStamp = (((long)ftTimeStamp.dwHighDateTime) << 32) + ftTimeStamp.dwLowDateTime;
			var timeStamp = DateTime.FromFileTime(lTimeStamp);

			return new SensorReport
			{
				Source = originator,
				TimeStamp = timeStamp,
				Values = SensorData.FromNativeReport(originator.internalObject, iReport)
			};
		}
	}
}