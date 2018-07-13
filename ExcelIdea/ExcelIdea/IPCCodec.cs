using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelIdea
{
    class IPCCodec
    {
        public static void Serialize(Stream stream_in, SMemIPCData data_in, long offset = 0)
        {
            /* Packing format:
             * 1 => messageType
             * 2 => messageValue
             * 3 => N
             * 4 => num_parameters
             * 5 => parameters
             * 6 => results
             * */

            /* ResultStruct packing format:
             * 1 => Cols
             * 2 => Rows
             * 3 => Type
             * 4 => Value
             * */

            BinaryWriter bw = new BinaryWriter(stream_in);
            stream_in.Position = offset;
            //static data
            bw.Write(data_in.messageType);
            bw.Write(data_in.messageValue ?? String.Empty);
            bw.Write(data_in.N);
            bw.Write(data_in.num_parameters);

            //parameters
            if (data_in.parameters == null)
                bw.Write((int)0);
            else
            {
                bw.Write(data_in.parameters.Count);   //write 1-d length L
                //for each element i in (L)
                for (int i = 0; i < data_in.parameters.Count; i++)
                {
                    //write 2-d length (M) at i
                    bw.Write(data_in.parameters[i].Count);

                    //for each element j in M
                    for (int j = 0; j < data_in.parameters[i].Count; j++)
                    {
                        //write result struct static values
                        bw.Write(((ResultStruct)data_in.parameters[i][j]).Cols);
                        bw.Write(((ResultStruct)data_in.parameters[i][j]).Rows);
                        bw.Write((byte)((ResultStruct)data_in.parameters[i][j]).Type);

                        if ((((ResultStruct)data_in.parameters[i][j]).Type & ResultEnum.Unit) == ResultEnum.Unit)
                        {
                            if ((((ResultStruct)data_in.parameters[i][j]).Type & ResultEnum.Float) == ResultEnum.Float)
                                bw.Write((float)((ResultStruct)data_in.parameters[i][j]).Value);
                            else
                                bw.Write((string)((ResultStruct)data_in.parameters[i][j]).Value);
                        }
                        else if ((((ResultStruct)data_in.parameters[i][j]).Type & ResultEnum.Array) == ResultEnum.Array)
                        {
                            //write the value array length
                            bw.Write((long)(((ResultStruct)data_in.parameters[i][j]).Cols * ((ResultStruct)data_in.parameters[i][j]).Rows));

                            if ((((ResultStruct)data_in.parameters[i][j]).Type & ResultEnum.Float) == ResultEnum.Float)
                                for (long k = 0; k < ((ResultStruct)data_in.parameters[i][j]).Cols * ((ResultStruct)data_in.parameters[i][j]).Rows; k++)
                                    bw.Write((((ResultStruct)data_in.parameters[i][j]).Value as float[])[k]);
                            else
                                for (long k = 0; k < ((ResultStruct)data_in.parameters[i][j]).Cols * ((ResultStruct)data_in.parameters[i][j]).Rows; k++)
                                    bw.Write((((ResultStruct)data_in.parameters[i][j]).Value as string[])[k]);
                        }
                    }
                }
            }

            //results
            if (data_in.Results == null)
                bw.Write((int)0);
            else
            {
                bw.Write(data_in.Results.Count);
                for (int i = 0; i < data_in.Results.Count; i++)
                {
                    bw.Write(((ResultStruct)data_in.Results[i]).Cols);
                    bw.Write(((ResultStruct)data_in.Results[i]).Rows);
                    bw.Write((byte)((ResultStruct)data_in.Results[i]).Type);
                    if ((((ResultStruct)data_in.Results[i]).Type & ResultEnum.Unit) == ResultEnum.Unit)
                        if ((((ResultStruct)data_in.Results[i]).Type & ResultEnum.Float) == ResultEnum.Float)
                            bw.Write((float)((ResultStruct)data_in.Results[i]).Value);
                        else
                            bw.Write((string)((ResultStruct)data_in.Results[i]).Value);
                    else if ((((ResultStruct)data_in.Results[i]).Type & ResultEnum.Array) == ResultEnum.Array)
                    {
                        bw.Write((long)((ResultStruct)data_in.Results[i]).Cols * ((ResultStruct)data_in.Results[i]).Rows);

                        if ((((ResultStruct)data_in.Results[i]).Type & ResultEnum.Float) == ResultEnum.Float)
                            for (long j = 0; j < ((ResultStruct)data_in.Results[i]).Cols * ((ResultStruct)data_in.Results[i]).Rows; j++)
                                bw.Write((((ResultStruct)data_in.Results[i]).Value as float[])[j]);
                        else
                            for (long j = 0; j < ((ResultStruct)data_in.Results[i]).Cols * ((ResultStruct)data_in.Results[i]).Rows; j++)
                                bw.Write((((ResultStruct)data_in.Results[i]).Value as string[])[j]);
                    }
                }
            }
        }

        public static SMemIPCData Deserialize(Stream stream_in, long offset = 0)
        {
            /* Packing format:
             * 1 => messageType
             * 2 => messageValue
             * 3 => N
             * 4 => num_parameters
             * 5 => parameters
             * 6 => results
             * */

            /* ResultStruct packing format:
             * 1 => Cols
             * 2 => Rows
             * 3 => Type
             * 4 => Value
             * */

            SMemIPCData result = new SMemIPCData();
            int subarray_count;
            long value_count;
            BinaryReader br = new BinaryReader(stream_in);
            ResultStruct rs_temp = new ResultStruct();
            stream_in.Position = offset;

            result.messageType = br.ReadByte();
            string temp = br.ReadString();
            result.messageValue = temp == string.Empty ? null : temp;
            result.N = br.ReadInt32();
            result.num_parameters = br.ReadInt32();

            //parameters
            int par_count = br.ReadInt32();
            if (par_count == 0)
                result.parameters = null;
            else
            {
                result.parameters = new List<List<ResultStruct>>(par_count);
                for (int i = 0; i < par_count; i++)
                {
                    subarray_count = br.ReadInt32();
                    result.parameters.Add(new List<ResultStruct>(subarray_count));
                    for (int j = 0; j < subarray_count; j++)
                    {
                        rs_temp.Cols = br.ReadInt32();
                        rs_temp.Rows = br.ReadInt32();
                        rs_temp.Type = (ResultEnum)br.ReadByte();
                        if ((rs_temp.Type & ResultEnum.Unit) == ResultEnum.Unit)
                            if ((rs_temp.Type & ResultEnum.Float) == ResultEnum.Float)
                                rs_temp.Value = br.ReadSingle();
                            else
                                rs_temp.Value = br.ReadString();
                        else if ((rs_temp.Type & ResultEnum.Array) == ResultEnum.Array)
                        {
                            value_count = br.ReadInt64();
                            if ((rs_temp.Type & ResultEnum.Float) == ResultEnum.Float)
                            {
                                rs_temp.Value = new float[value_count];
                                for (int k = 0; k < value_count; k++)
                                    (rs_temp.Value as float[])[k] = br.ReadSingle();
                            }
                            else
                            {
                                rs_temp.Value = new string[value_count];
                                for (int k = 0; k < value_count; k++)
                                    (rs_temp.Value as string[])[k] = br.ReadString();
                            }
                        }
                        result.parameters[i].Add(rs_temp);
                    }
                }
            }

            //results
            par_count = br.ReadInt32();
            if (par_count == 0)
                result.Results = null;
            else
            {
                result.Results = new List<ResultStruct>(par_count);
                for (int i = 0; i < par_count; i++)
                {
                    rs_temp.Cols = br.ReadInt32();
                    rs_temp.Rows = br.ReadInt32();
                    rs_temp.Type = (ResultEnum)br.ReadByte();
                    if ((rs_temp.Type & ResultEnum.Unit) == ResultEnum.Unit)
                        if ((rs_temp.Type & ResultEnum.Float) == ResultEnum.Float)
                            rs_temp.Value = br.ReadSingle();
                        else
                            rs_temp.Value = br.ReadString();
                    else if ((rs_temp.Type & ResultEnum.Array) == ResultEnum.Array)
                    {
                        value_count = br.ReadInt64();
                        if ((rs_temp.Type & ResultEnum.Float) == ResultEnum.Float)
                        {
                            rs_temp.Value = new float[value_count];
                            for (int k = 0; k < value_count; k++)
                                (rs_temp.Value as float[])[k] = br.ReadSingle();
                        }
                        else
                        {
                            rs_temp.Value = new string[value_count];
                            for (int k = 0; k < value_count; k++)
                                (rs_temp.Value as string[])[k] = br.ReadString();
                        }
                    }
                    result.Results.Add(rs_temp);
                }
            }
            return result;
        }
    }
}
