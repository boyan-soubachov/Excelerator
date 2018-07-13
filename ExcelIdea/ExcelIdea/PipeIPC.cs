using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Pipes;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace ExcelIdea
{
    class PipeClient
    {
        private string pipeName;
        private int timeOut;
        public PipeClient(string pipe_name, int time_out = 500)
        {
            this.pipeName = pipe_name;
            this.timeOut = time_out;
        }

        public void Send(PipeMessage msg_in)
        {
            try
            {
                NamedPipeClientStream pipeStream = new NamedPipeClientStream(".", pipeName, PipeDirection.Out, PipeOptions.Asynchronous);

                pipeStream.Connect(timeOut);

                IFormatter fmtr = new BinaryFormatter();
                fmtr.Serialize(pipeStream, msg_in);
            }
            catch (Exception except)
            {
                throw new SystemException(except.Message);
            }
        }

        private void AsyncSend(IAsyncResult iar)
        {
            try
            {
                NamedPipeClientStream pipeStream = (NamedPipeClientStream)iar.AsyncState;

                pipeStream.EndWrite(iar);
                pipeStream.Flush();
                pipeStream.Close();
                pipeStream.Dispose();
            }
            catch (Exception except)
            {
                throw new SystemException(except.Message);
            }
        }
    }
}
