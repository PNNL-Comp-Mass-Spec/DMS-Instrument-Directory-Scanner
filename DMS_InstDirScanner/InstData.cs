//*********************************************************************************************************
// Written by Dave Clark for the US Department of Energy
// Pacific Northwest National Laboratory, Richland, WA
// Copyright 2009, Battelle Memorial Institute
// Created 07/27/2009
//
//*********************************************************************************************************

using System.IO;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Class to hold data for each instrument
    /// </summary>
    public class InstrumentData
    {
        // Ignore Spelling: fso, secfso, bionet

        /// <summary>
        /// Storage volume, for example, \\QExactP04.bionet\
        /// </summary>
        public string StorageVolume { get; set; }

        /// <summary>
        /// Storage path, typically ProteomicsData\
        /// </summary>
        public string StoragePath { get; set; }

        /// <summary>
        /// Capture method
        /// </summary>
        /// <remarks>
        /// fso if on a domain computer
        /// secfso if on bionet
        /// </remarks>
        public string CaptureMethod { get; set; }

        /// <summary>
        /// Instrument name
        /// </summary>
        public string InstName { get; set; }

        /// <summary>
        /// Instrument name: StorageVolume\StoragePath
        /// </summary>
        public override string ToString()
        {
            return InstName + ": " + Path.Combine(StorageVolume, StoragePath);
        }
    }
}