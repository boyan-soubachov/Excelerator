/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * 
 * This class implements named pipe IPC for communication to clients providing data.
 * 
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Pipes;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Diagnostics;

namespace ExcelIdea
{
    public delegate void DelegateMessage(string Reply);

    //server code
    class PipeIPCServer
    {
        string _pipeName;

        public void Listen(string PipeName)
        {
            try
            {
                _pipeName = PipeName;
                NamedPipeServerStream pipeServer = new NamedPipeServerStream(PipeName, PipeDirection.In, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);

                pipeServer.BeginWaitForConnection(new AsyncCallback(WaitForConnectionCallBack), pipeServer);
            }
            catch (Exception oEX)
            {
                throw new SystemException(oEX.Message);
            }
        }

        private void WaitForConnectionCallBack(IAsyncResult iar)
        {
            try
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();
                Console.WriteLine("=> Receiving pipe message");

                NamedPipeServerStream pipeServer = (NamedPipeServerStream)iar.AsyncState;
                pipeServer.EndWaitForConnection(iar);

                //PipeMessage.Invoke(stringData);

                IFormatter fmtr = new BinaryFormatter();
                //fmtr.Binder = new MessageAssemblyBinder();
                PipeMessage msg_in = (PipeMessage)fmtr.Deserialize(pipeServer); // <- possibly the slow bit. Maybe use named pipes for signalling and shared memory for chunks.

                pipeServer.Close();
                pipeServer = null;
                pipeServer = new NamedPipeServerStream(_pipeName, PipeDirection.In, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
                sw.Stop();
                Console.WriteLine("\t --Pipe message received, time: " + sw.Elapsed.TotalMilliseconds + " ms");
                pipeServer.BeginWaitForConnection(new AsyncCallback(WaitForConnectionCallBack), pipeServer);
            }
            catch(Exception e)
            {
                throw new SystemException(e.Message);
            }
        }
    }

    //client code
    class PipeIPCClient
    {

    }
}
