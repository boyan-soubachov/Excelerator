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
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelIdea
{
    class SharedMemoryIPC
    {
        private MemoryMappedFile phile;
        private MemoryMappedViewStream phile_stream;
        private MemoryMappedViewAccessor phile_prefix_stream; //stores the length prefix of the data
        private EventWaitHandle message_wait, message_handled;
        private Mutex sharedmem_mutex;
        private ComputeCoreVBA vba_core;

        public SharedMemoryIPC()
        {
            vba_core = new ComputeCoreVBA();
            message_wait = new EventWaitHandle(false, EventResetMode.AutoReset, "ijiBearExcelIdeaMESSAGEWAIT");
            message_handled = new EventWaitHandle(false, EventResetMode.AutoReset, "ijiBearExcelIdeaMESSAGEHANDLED");
            
            phile = MemoryMappedFile.CreateNew("ijiBearExcelIdeaSHAREDMEM", 1024 * 1024 * 1024);
            phile_prefix_stream = phile.CreateViewAccessor(0, sizeof(long));
            phile_stream = phile.CreateViewStream(sizeof(long), 0);

            sharedmem_mutex = new Mutex(true, "ijiBearExcelIdeaSHAREDMEMMUTEX");
            sharedmem_mutex.ReleaseMutex();

            NewEventWaitHandleThread();
        }

        private void NewEventWaitHandleThread()
        {
            bool exit = false;
            while (true)
            {
                exit = EventWaitHandleThread();
                if (exit) break;
            }
        }

        private bool EventWaitHandleThread()
        {
            message_wait.WaitOne();
            //read what's in the shared memory
            SMemIPCData data = IPCCodec.Deserialize(phile_stream);

            //pass to program for processing
            switch (data.messageType)
            {
                case 2: //compute function
                    SMemIPCData results = new SMemIPCData();
                    results.messageType = 3;

                    //call VBA handler
                    List<ResultStruct> res_temp = vba_core.ComputeOperator(data.messageValue, data.parameters, data.N, data.num_parameters);

                    //place results in shared memory
                    results.Results = res_temp;
                    IPCCodec.Serialize(phile_stream, results);

                    //store the result prefix length
                    phile_prefix_stream.Write(0, phile_stream.Position);
                    break;

                case 4: //store in GPU
                    break;

                case 5: //retrieve from GPU
                    break;

                case 6:
                    string[] parameters = data.messageValue.Split('|');
                    Program.HandleFile(parameters[0], parameters[1]);
                    break;
                default:
                    break;
            }

            message_handled.Set();
            return false;
        }

        public void Close()
        {
            sharedmem_mutex.Close();
            message_handled.Close();
            message_wait.Close();
            phile_stream.Close();
            phile.Dispose();
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
