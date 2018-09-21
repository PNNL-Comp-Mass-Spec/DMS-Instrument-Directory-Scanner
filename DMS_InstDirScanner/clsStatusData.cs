﻿
//*********************************************************************************************************
// Written by Dave Clark for the US Department of Energy
// Pacific Northwest National Laboratory, Richland, WA
// Copyright 2009, Battelle Memorial Institute
// Created 08/14/2009
//*********************************************************************************************************

using System.Collections.Generic;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Class to hold long-term data for status reporting.
    /// This is a hack to avoid adding an instance of the status file class to the log tools class
    /// </summary>
    class clsStatusData
    {

        private static string m_MostRecentLogMessage;
        private static readonly Queue<string> m_ErrorQueue = new Queue<string>();


        public static string MostRecentLogMessage
        {
            get => m_MostRecentLogMessage;
            set
            {
                // Filter out routine startup and shutdown messages
                if (value.Contains("=== Started") || (value.Contains("===== Closing")))
                {
                    // Do nothing
                }
                else
                {
                    m_MostRecentLogMessage = value;
                }
            }
        }

        public static Queue<string> ErrorQueue => m_ErrorQueue;


        public static void AddErrorMessage(string ErrMsg)
        {
            // Add the most recent error message
            m_ErrorQueue.Enqueue(ErrMsg);

            // If there are > 4 entries in the queue, delete the oldest ones
            while (m_ErrorQueue.Count > 4)
            {
                m_ErrorQueue.Dequeue();
            }
        }

    }
}