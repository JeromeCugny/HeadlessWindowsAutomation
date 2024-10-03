using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace HeadlessWindowsAutomation
{
    /// <summary>
    /// Provides utility methods for managing and interacting with processes.
    /// </summary>
    public class ProcessHelper
    {
        /// <summary>
        /// Optional. Function called when checking the validity of a process in the context of "GetApplicationProcessIdFromParent".
        /// </summary>
        public Func<Process, bool> AdditionalProcessValidCheck { get; set; } = null;

        /// <summary>
        /// Method to get the PID of the actual application process started by the command. 
        /// Useful when your application is a child of another process.
        /// </summary>
        /// <param name="parent">Parent process which started your application.</param>
        /// <param name="maxWaitTimeMS">Optional. Max wait time is ms for your application to start.</param>
        /// <param name="checkIntervalMS">Optional. Time interval in ms between each check for your application.</param>
        /// <returns>The application process id or the parent id when not found.</returns>
        public int GetApplicationProcessIdFromParent(Process parent, int maxWaitTimeMS = 15000, int checkIntervalMS = 500)
        {
            int parentPid = parent.Id;
            DateTime startTime = parent.StartTime;
            int elapsedTime = 0;

            while (elapsedTime < maxWaitTimeMS)
            {
                foreach (Process process in Process.GetProcesses())
                {
                    if (this.IsProcessValid(process, startTime) && this.IsAncestorProcess(process.Id, parentPid))
                    {
                        Console.WriteLine($"Found child process {process.ProcessName}, pid: {process.Id} from parent {parentPid}");
                        return process.Id;
                    }
                }
                System.Threading.Thread.Sleep(checkIntervalMS); // Wait for a short period before checking again
                elapsedTime += checkIntervalMS;
            }
            // Not found
            return parentPid;
        }

        /// <summary>
        /// Check if a process is valid in the context of 'GetApplicationProcessIdFromParent'.  
        /// A process is valid if it started after "startTime", has a window handle and respect "AdditionalProcessValidCheck".
        /// </summary>
        /// <param name="proc">Process to check.</param>
        /// <param name="startTime">Time when the expected parent process started.</param>
        /// <returns>True if the process is valid.</returns>
        private bool IsProcessValid(Process proc, DateTime startTime)
        {
            try
            {
                // true if started afterward and with a window
                bool isValid = proc != null && proc.StartTime >= startTime && proc.MainWindowHandle != IntPtr.Zero;
                if (this.AdditionalProcessValidCheck != null) isValid = isValid && this.AdditionalProcessValidCheck(proc);
                return isValid;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// Check if the given process "ancestorPid" is the ancestor of the other process "pid".
        /// </summary>
        /// <param name="pid">Process id for the child process.</param>
        /// <param name="ancestorPid">Process id for the ancestor process.</param>
        /// <param name="maxDepth">Optional. Max number of ancestors to check.</param>
        /// <returns>True if "ancestorPid" is an ancestor of "pid".</returns>
        public bool IsAncestorProcess(int pid, int ancestorPid, int maxDepth = 10)
        {
            int count = 0;
            try
            {
                while (pid != 0 && count < maxDepth)
                {
                    if (pid == ancestorPid)
                    {
                        return true;
                    }
                    pid = GetParentProcessId(pid);
                    count++;
                }
            }
            catch (Exception e)
            {
                // Skip processes that we cannot access
            }
            return false;
        }

        /// <summary>
        /// P/Invoke to get the parent process ID
        /// </summary>
        [DllImport("ntdll.dll")]
        private static extern int NtQueryInformationProcess(IntPtr processHandle, int processInformationClass, ref PROCESS_BASIC_INFORMATION processInformation, uint processInformationLength, out uint returnLength);

        private struct PROCESS_BASIC_INFORMATION
        {
            public IntPtr Reserved1;
            public IntPtr PebBaseAddress;
            public IntPtr Reserved2_0;
            public IntPtr Reserved2_1;
            public IntPtr UniqueProcessId;
            public IntPtr InheritedFromUniqueProcessId;
        }

        /// <summary>
        /// Get the parent process id for a given process id.
        /// </summary>
        /// <param name="pid">Process id to get the parent from</param>
        /// <returns>Process id of the parent. Throw on failure.</returns>
        public int GetParentProcessId(int pid)
        {
            // This method can throw, especially on access denied
            PROCESS_BASIC_INFORMATION pbi = new PROCESS_BASIC_INFORMATION();
            uint returnLength;
            IntPtr hProcess = Process.GetProcessById(pid).Handle;
            int status = NtQueryInformationProcess(hProcess, 0, ref pbi, (uint)Marshal.SizeOf(pbi), out returnLength);
            if (status != 0)
            {
                throw new Exception("Failed to query process information");
            }
            return pbi.InheritedFromUniqueProcessId.ToInt32();
        }

        /// <summary>
        /// Kill the whole process tree for the given process.
        /// </summary>
        /// <param name="pid">The process ID of the root process to kill.</param>
        public void KillProcessTree(int pid)
        {
            try
            {
                Process process = Process.GetProcessById(pid);
                KillProcessTree(process);
            }
            catch (Exception e) { }
        }

        /// <summary>
        /// Kill the whole process tree for the given process.
        /// </summary>
        /// <param name="rootProcess">The root process to kill.</param>
        public void KillProcessTree(Process rootProcess)
        {
            void AuxKill(Process proc)
            {
                try
                {
                    string name = proc.ProcessName;
                    int id = proc.Id;
                    proc.Kill();
                    Console.WriteLine($"> Kill {name}, pid {id}");
                }
                catch (Exception e) { }
            }

            var processesToKill = new List<Process> { rootProcess };
            int rootId = rootProcess.Id;

            foreach (Process process in Process.GetProcesses())
            {
                if (this.IsAncestorProcess(process.Id, rootId))
                {
                    processesToKill.Add(process);
                }
            }

            foreach (Process proc in processesToKill)
            {
                AuxKill(proc);
            }
        }
    }
}
