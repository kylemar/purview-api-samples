// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Diagnostics;
using System.IO;
using System.Text;

namespace PurviewAPIExp
{
    internal enum LogType
    {
        All,
        Screen,
        Console,
        File
    }
    /// <summary>
    /// Possibility the worlds most simple logger
    /// </summary>
    internal class Logger(StringBuilder sbLog)
    {
        string? fileName;
        readonly StringBuilder sblog = sbLog;

        public void Log(string message, LogType logType = LogType.All)
        {
            string messageToShow;

            messageToShow = $"{DateTime.Now}-{message}";

            if (logType == LogType.All || logType == LogType.Console)
            {
                Debug.WriteLine(messageToShow);
            }

            if (logType == LogType.All || logType == LogType.Screen)
            {
                if (sblog.Length > 65536)
                {
                    sblog.Clear();
                }
                sblog.AppendLine(messageToShow);
            }

            if (logType == LogType.All || logType == LogType.File)
            {
                if (fileName == null)
                {
                    fileName = $"{DateTime.Now:yyyy-MM-dd-HH-mm-ss}.log";
                }
                try
                {
                    File.AppendAllText(fileName, $"{messageToShow}\n");
                }
                catch { }
            }
        }
    }
}
