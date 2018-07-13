/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelIdea
{
    class Program
    {
        static ComputeCore compute_core;
        static EventWaitHandle CC_up;
        static SharedMemoryIPC ipc_server;
        public static CUDAProcessor CUDA_Core;

        static void Main(string[] args)
        {
            CC_up = new EventWaitHandle(true, EventResetMode.ManualReset, "ijiBearExcelIdeaCCUp");
            CC_up.Set();

            if (args.Length > 0) //if we've got input parameters
            {
                /* 0 = source type
                 * 1->... = type-specific parameters
                 * last = auth token covering the previous parameters
                 */

                //check authentication token
                if (!CheckAuthToken(args))
                {
                    Console.WriteLine("ERROR: This machine has not been licensed for use. Press any key to EXIT.");
                    Console.ReadLine();
                    return;
                }
                switch (args[0])
                {
                    case "file":
                        HandleFile(args[1], args[2]);
                        break;
                    case "ipc":
                        HandleIPC(args);
                        break;
                    case "server":
                        StartServer();
                        break;
                    default:
                        break;
                }

            }
            else
            {
                //NOTE: Testing code, remove later on.
                //compute_core = new ComputeCore(@"C:\Users\Bob\Desktop\tests\test5.xlsx");
                //compute_core.StartCompute();
                //compute_core.SaveToFile(@"C:\Users\Bob\Desktop\results_test.xlsx");
                StartServer();
            }
            if (ipc_server != null)
                ipc_server.Close();
            CC_up.Reset();
        }

        private static void StartServer()
        {
            //fire up the CUDA_Core;
            CUDA_Core = new CUDAProcessor();
            CUDA_Core.Init();
            CUDA_Core.SetContext();

            //fire up the IPC server
            ipc_server = new SharedMemoryIPC();   
        }

        private static void HandleIPC(string[] args)
        {
            //TODO: This needs completion
        }

        private static bool CheckAuthToken(string[] args)
        {
            switch (args[0])
            {
                case "file":
                    return Tools.SHA512String("!hello file " + args[1] + "|" + args[2] + "h4x0r!") == args[3] ? true : false;
                case "ipc":
                    return Tools.SHA512String("!hello ipc h4x0r!") == args[1] ? true : false;
                case "server":
                    return Tools.SHA512String("!hello server h4x0r!") == args[1] ? true : false;
                default:
                    return false;
            }
        }

        public static void HandleFile(string source, string dest)
        {
            /* 1 = source file
             * 2 = destination file
             * */

            //create compute core and do calculations
            if (compute_core == null)
                compute_core = new ComputeCore(source);
            else
                compute_core.LoadFromFile(source);

            compute_core.StartCompute();
            compute_core.SaveToFile(dest);
        }
    }
}
