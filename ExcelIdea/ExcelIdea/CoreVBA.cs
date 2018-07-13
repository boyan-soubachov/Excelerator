using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ExcelIdea
{
    [ComVisible(true)]
    public interface ICoreVBA
    {
        float[] ArrayRandom(int n);
        float[] PowerTransform(float[] data, float lambda);
        float[] BlackScholesCallPrice(float[] spot_price, float[] maturity_time, float[] strike_price, float[] risk_free_rate, float[] volatility);
        float[] RandN(float mean, float std_dev, int count);
        float[] MCWienerCallPrice(float[] spot_price, float[] maturity_time, float[] risk_free_rate, float[] volatility, int time_steps, float[] strike_price, byte style);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class CoreVBA
    {
        public CoreVBA()
        {

        }

        public float[] MCWienerCallPrice(float[] spot_price, float[] maturity_time, float[] risk_free_rate, float[] volatility, int time_steps, float[] strike_price, byte style)
        {
            /* style types:
             * 0 => European
             * 1 => American
             * 2 => Asian
             * 3 => Lookback
             * 4 => Barrier
             * */

            //check that all inputs are of the same length
            if (maturity_time.Length != spot_price.Length ||
                risk_free_rate.Length != spot_price.Length ||
                strike_price.Length != spot_price.Length ||
                volatility.Length != spot_price.Length)
            {
                //error here for mismatching lengths
                return null;
            }

            //prepare request
            SMemIPCData msg_data = new SMemIPCData();
            msg_data.messageType = 2;
            msg_data.messageValue = "MC-WIENER-CALL-PRICE";
            msg_data.N = 1;
            msg_data.parameters = new List<List<ResultStruct>>(7);
            for (int i = 0; i < 7; i++)
                msg_data.parameters.Add(new List<ResultStruct>());
            ResultStruct rs_temp = new ResultStruct();

            //spot price
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = spot_price;
            msg_data.parameters[0].Add(rs_temp);

            //time to maturity
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = maturity_time;
            msg_data.parameters[1].Add(rs_temp);

            //risk free rate
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = risk_free_rate;
            msg_data.parameters[2].Add(rs_temp);

            //volatility
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = volatility;
            msg_data.parameters[3].Add(rs_temp);

            //number of steps
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
            rs_temp.Value = (float)time_steps;
            msg_data.parameters[4].Add(rs_temp);

            //strike price
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = strike_price;
            msg_data.parameters[5].Add(rs_temp);

            //style
            rs_temp.Cols = 1;
            rs_temp.Rows = 1;
            rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
            rs_temp.Value = (float)style;
            msg_data.parameters[6].Add(rs_temp);

            SMemIPCData response = Globals.ThisAddIn.smem_ipc.SendMessage(msg_data);
            if (response.messageType == 3)
                return response.Results[0].Value as float[];
            else
                return null;
        }

        public float[] RandN(float mean, float std_dev, int count)
        {
            if ((std_dev < 0) || (count <= 0)) //check std_dev and count
                return null;

            //prepare request
            SMemIPCData msg_data = new SMemIPCData();
            msg_data.messageType = 2;
            msg_data.messageValue = "RANDN";
            msg_data.N = 1;
            msg_data.parameters = new List<List<ResultStruct>>(3);
            for (int i = 0; i < 3; i++)
                msg_data.parameters.Add(new List<ResultStruct>());
            ResultStruct rs_temp = new ResultStruct();

            //mean
            rs_temp.Cols = 1;
            rs_temp.Rows = 1;
            rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
            rs_temp.Value = mean;
            msg_data.parameters[0].Add(rs_temp);

            //std dev
            rs_temp.Cols = 1;
            rs_temp.Rows = 1;
            rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
            rs_temp.Value = std_dev;
            msg_data.parameters[1].Add(rs_temp);

            //count
            rs_temp.Cols = 1;
            rs_temp.Rows = 1;
            rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
            rs_temp.Value = (float)count;
            msg_data.parameters[2].Add(rs_temp);

            SMemIPCData response = Globals.ThisAddIn.smem_ipc.SendMessage(msg_data);
            if (response.messageType == 3)
                return response.Results[0].Value as float[];
            else
                return null;
        }

        public float[] BlackScholesCallPrice(float[] spot_price, float[] maturity_time, float[] strike_price, float[] risk_free_rate, float[] volatility)
        {
            //check that all inputs are of the same length
            if (maturity_time.Length != spot_price.Length ||
                strike_price.Length != spot_price.Length ||
                risk_free_rate.Length != spot_price.Length ||
                volatility.Length != spot_price.Length)
            {
                //error here for mismatching lengths
                return null;
            }

            //prepare request
            SMemIPCData msg_data = new SMemIPCData();
            msg_data.messageType = 2;
            msg_data.messageValue = "BLACK-SCHOLES-CALL";
            msg_data.N = 1;
            msg_data.parameters = new List<List<ResultStruct>>(5);
            for (int i = 0; i < 5; i++)
                msg_data.parameters.Add(new List<ResultStruct>());
            ResultStruct rs_temp = new ResultStruct();

            //spot price
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = spot_price;
            msg_data.parameters[0].Add(rs_temp);

            //time to maturity
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = maturity_time;
            msg_data.parameters[1].Add(rs_temp);

            //strike price
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = strike_price;
            msg_data.parameters[2].Add(rs_temp);

            //risk free rate
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = risk_free_rate;
            msg_data.parameters[3].Add(rs_temp);

            //volatility
            rs_temp.Cols = 1;
            rs_temp.Rows = spot_price.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = volatility;
            msg_data.parameters[4].Add(rs_temp);

            SMemIPCData response = Globals.ThisAddIn.smem_ipc.SendMessage(msg_data);
            if (response.messageType == 3)
                return response.Results[0].Value as float[];
            else
                return null;
        }

        public float[] PowerTransform(float[] data, float lambda)
        {
            //check that all the data is > 0
            for (int i = 0; i < data.Length; i++)
                if (data[i] <= 0)
                    //send error
                    return null;

            //send request
            SMemIPCData msg_data = new SMemIPCData();
            msg_data.messageType = 2;
            msg_data.messageValue = "POWER-TRANSFORM";
            msg_data.N = 1;
            msg_data.parameters = new List<List<ResultStruct>>(2);
            msg_data.parameters[0] = new List<ResultStruct>();
            msg_data.parameters[1] = new List<ResultStruct>();
            ResultStruct rs_temp = new ResultStruct();
            //data
            rs_temp.Cols = 1;
            rs_temp.Rows = data.Length;
            rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
            rs_temp.Value = data;
            msg_data.parameters[0].Add(rs_temp);

            //lambda
            rs_temp.Rows = 1;
            rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
            rs_temp.Value = lambda;
            msg_data.parameters[1].Add(rs_temp);

            SMemIPCData response = Globals.ThisAddIn.smem_ipc.SendMessage(msg_data);
            if (response.messageType == 3)
                return response.Results[0].Value as float[];
            else
                return null;
        }

        public float[] ArrayRandom(int n)
        {
            SMemIPCData msg_data = new SMemIPCData();
            msg_data.messageType = 2;
            msg_data.messageValue = "RANDOM";
            msg_data.N = n;
            SMemIPCData response = Globals.ThisAddIn.smem_ipc.SendMessage(msg_data);
            if (response.messageType == 3)
            {
                float[] res = new float[response.Results.Count];
                Parallel.For(0, response.Results.Count, i =>
                {
                    res[i] = (float)response.Results[i].Value;
                });
                return res;
                //TODO: mutex needs to handle case where if current client crashes, it must free itself after a certain timeout.
            }
            else
                return null;
        }
    }
}
