/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * 
 * This class implements IPC using shared memory and EventWaitHandle for signalling.
 * 
 * */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelIdea
{
    public class SharedMemoryIPC
    {
        private MemoryMappedFile phile;
        private MemoryMappedViewStream phile_stream;
        private MemoryMappedViewAccessor phile_prefix;
        private EventWaitHandle message_wait, message_handled;
        private Mutex sharedmem_mutex;

        public SharedMemoryIPC()
        {
            bool success = true;
            success = EventWaitHandle.TryOpenExisting("ijiBearExcelIdeaMESSAGEWAIT", out message_wait);
            success = EventWaitHandle.TryOpenExisting("ijiBearExcelIdeaMESSAGEHANDLED", out message_handled);

            if (!success)
            {
                //throw error here for computation core not running
            }
            else
            {
                phile = MemoryMappedFile.OpenExisting("ijiBearExcelIdeaSHAREDMEM");
                phile_prefix = phile.CreateViewAccessor(0, sizeof(long)); //the length prefix
                phile_stream = phile.CreateViewStream(sizeof(long), 128 * 1024 * 1024);
                sharedmem_mutex = Mutex.OpenExisting("ijiBearExcelIdeaSHAREDMEMMUTEX", System.Security.AccessControl.MutexRights.FullControl);
            }
        }

        public SMemIPCData SendMessage(SMemIPCData data_in)
        {
            try
            {
                sharedmem_mutex.WaitOne(); //take the mutex

                //save message in shared memory
                IPCCodec.Serialize(phile_stream, data_in);
                //prepare the message prefix
                phile_prefix.Write(0, phile_stream.Position);
                message_wait.Set(); //signal message ready

                message_handled.WaitOne(); //wait for reply ready signal
                //copy results from shared memory
                long result_size = phile_prefix.ReadInt64(0);
                if (result_size > phile_stream.Capacity)
                {
                    //enlarge the view stream if the result is too big
                    try
                    {
                        phile_stream.Close();
                        phile_stream = phile.CreateViewStream(sizeof(long), result_size);
                    }
                    catch
                    {
                        throw new SystemException("ijiCore: Result set too large.");
                    }
                }
                return IPCCodec.Deserialize(phile_stream);
            }
            catch(Exception e)
            {
                return null;
            }
            finally
            {
                sharedmem_mutex.ReleaseMutex();
            }
        }
    }

    public class SMemIPCData
    {
        /* Message types:
         * 1 - Reserved
         * 2 - Compute request
         * 3 - Results response
         * 4 - Store data into GPU
         * 5 - Retrieve data from GPU
         * 6 - Compute Excel file
         * */

        public byte messageType;

        public string messageValue; //stores compute_string for compute requests as well as results

        public List<List<ResultStruct>> parameters; //parameters for computation

        public int N; //number of parallels

        public int num_parameters; //number of parameters

        public List<ResultStruct> Results; //results of computation
    }
}
