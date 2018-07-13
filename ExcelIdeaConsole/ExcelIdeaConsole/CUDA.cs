/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * 
 * This class contains the CUDA fields used to perform computations on the GPU.
 * 
 * */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Cudafy;
using Cudafy.Atomics;
using Cudafy.Host;
using Cudafy.Translator;
using Cudafy.Maths.RAND;
using Cudafy.Maths.BLAS;
using System.Diagnostics;

namespace ExcelIdea
{
    class CUDAProcessor
    {
        static CudafyModule km;
        static GPGPU gpu;
        static GPGPUBLAS gpu_blas;
        static GPGPURAND gpu_rand;
        static GPGPUProperties gpu_prop;
        private static object gpu_lock = new object();
        public bool core_up = false;

        public CUDAProcessor()
        {
               
        }
        
        public void Init()
        {
            km = CudafyModule.TryDeserialize();
            if (km == null || !km.TryVerifyChecksums())
            {
                km = CudafyTranslator.Cudafy(ePlatform.x64, eArchitecture.sm_35);
                km.Serialize();
            }
            CudafyTranslator.GenerateDebug = true; //NOTE: temporary for development
            gpu = CudafyHost.GetDevice(eGPUType.Cuda);
            gpu.LoadModule(km);

            gpu_blas = GPGPUBLAS.Create(gpu);
            gpu_rand = GPGPURAND.Create(gpu, curandRngType.CURAND_RNG_PSEUDO_DEFAULT);
            gpu_prop = CudafyHost.GetDeviceProperties(CudafyModes.Target, false).First();
            core_up = true;
            //Console.WriteLine("\t Using GPU: " + gpu_prop.Name + ", Memory: " + (gpu_prop.TotalMemory / (1024 * 1024)).ToString() + " MB");
            //Console.WriteLine("\t Grid size: " + gpu_prop.MaxGridSize.x.ToString() + " x " + gpu_prop.MaxGridSize.y.ToString() + " x " + gpu_prop.MaxGridSize.z.ToString());
            //Console.WriteLine("\t Constant memory: " + gpu_prop.TotalConstantMemory.ToString() + ", Shared memory / block: " + gpu_prop.SharedMemoryPerBlock.ToString());
            Console.WriteLine("=> CUDA cores online");
        }

        public void SetContext()
        {
            gpu.SetCurrentContext();
            gpu.EnableMultithreading();
        }

        public void Dispose()
        {
            gpu.DisableMultithreading();
            gpu.Synchronize();
            km.TrySerialize();
            gpu.FreeAll();
            gpu_blas.Dispose();
            gpu_rand.Dispose();
            gpu.Dispose();
            core_up = false;
        }

        #region VBA Compute launcher

        public List<ResultStruct> VBACompute(string operation, List<List<ResultStruct>> parameters, int N, int num_arguments)
        {
            List<ResultStruct> res;

            switch (operation)
            {
                case "STUTZER-INDEX-OPTIMAL-ALLOCATION":
                    return VBAFn_StutzerIndexOptimalWeights(N, parameters, num_arguments);
                case "STUTZER-INDEX":
                    return VBAFn_StutzerIndex(N, parameters, num_arguments);
                case "MC-WIENER-CALL-PRICE":
                    return VBAFn_MCWiener(N, parameters, num_arguments);
                case "RANDN": //gaussian random
                    return VBAFn_RandN(N, parameters, num_arguments);
                case "BLACK-SCHOLES-CALL":
                    return VBAFn_BlackScholes(N, parameters, num_arguments);
                case "BLACK-SCHOLES-PUT":
                    return VBAFn_BlackScholes(N, parameters, num_arguments, false);
                case "POWER-TRANSFORM":
                    return VBAFn_PowerTransform(N, parameters, num_arguments);
                default:
                    res = new List<ResultStruct>();
                    res.AddRange(new ResultStruct[N]);
                    Parallel.For(0, N, i => 
                    {
                        ResultStruct rs_temp = new ResultStruct();
                        rs_temp.Type = ResultEnum.Error | ResultEnum.String;
                        rs_temp.Value = "#@@CU_2";
                        res[i] = rs_temp;
                    });
                    return res;
            }
        }

        private List<ResultStruct> VBAFn_BlackScholes(int N, List<List<ResultStruct>> parameters, int num_arguments, bool call = true)
        {
            /* parameters:
             * 0 => spot price (S)
             * 1 => time to maturity (T-t)
             * 2 => strike price (K)
             * 3 => risk-free rate (r)
             * 4 => volatility (sigma)
             * */

            //NOTE: solution to problem of mixed input parameters (e.g. some arrays, others units) = require that all are of same type and same length.
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            Parallel.For(0, N, i => 
            {
                int num_parallels = parameters[0][i].Cols * parameters[0][i].Rows;
                float[] A = new float[num_parallels];
                gpu.Lock();
                float[] dev_A = gpu.CopyToDevice<float>(parameters[0][i].Value as float[]);
                float[] dev_B = gpu.CopyToDevice<float>(parameters[1][i].Value as float[]);
                float[] dev_C = gpu.CopyToDevice<float>(parameters[2][i].Value as float[]);
                float[] dev_D = gpu.CopyToDevice<float>(parameters[3][i].Value as float[]);
                float[] dev_E = gpu.CopyToDevice<float>(parameters[4][i].Value as float[]);

                if (call)
                    gpu.Launch(num_parallels, 1, "vba_op_black_scholes_call", dev_A, dev_B, dev_C, dev_D, dev_E);
                else
                    gpu.Launch(num_parallels, 1, "vba_op_black_scholes_put", dev_A, dev_B, dev_C, dev_D, dev_E);

                gpu.Synchronize();

                gpu.CopyFromDevice<float>(dev_A, A);
                gpu.Free(dev_A);
                gpu.Free(dev_B);
                gpu.Free(dev_C);
                gpu.Free(dev_D);
                gpu.Free(dev_E);
                gpu.Unlock();
                ResultStruct rs_temp = new ResultStruct();
                rs_temp.Cols = parameters[0][i].Cols;
                rs_temp.Rows = parameters[0][i].Rows;
                rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
                rs_temp.Value = A;
                res[i] = rs_temp;
            });

            return res;
        }

        private List<ResultStruct> VBAFn_PowerTransform(int N, List<List<ResultStruct>> parameters, int num_arguments)
        {
            /* Parameters:
             * 0 => The array, a
             * 1 => Exponent, lambda
             * */
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            Parallel.For(0, N, i =>
            {
                ResultStruct rs_temp = new ResultStruct();

                if ((parameters[0][i].Type & ResultEnum.Array) == ResultEnum.Array)
                {
                    //calculate geometric mean
                    float GM = 1;
                    for (int j = 0; j < parameters[0][i].Cols * parameters[0][i].Rows; j++)
                        GM *= (parameters[0][i].Value as float[])[j];
                    GM = (float)Math.Pow(GM, 1 / (parameters[0][i].Cols * parameters[0][i].Rows));

                    //call CUDA function
                    float[] A = new float[parameters[0][i].Cols * parameters[0][i].Rows];
                    gpu.Lock();
                    float[] dev_A = gpu.CopyToDevice<float>(parameters[0][i].Value as float[]);

                    gpu.Launch(A.Length, 1, "vba_op_power_transform", dev_A, (float)parameters[1][i].Value, GM);

                    gpu.CopyFromDevice<float>(dev_A, A);
                    gpu.Free(dev_A);
                    gpu.Unlock();

                    rs_temp.Cols = parameters[0][i].Cols;
                    rs_temp.Rows = parameters[0][i].Rows;
                    rs_temp.Type = ResultEnum.Array | ResultEnum.Float;
                    rs_temp.Value = A;
                    res[i] = rs_temp;
                }
            });

            return res;
        }

        private List<ResultStruct> VBAFn_RandN(int N, List<List<ResultStruct>> parameters, int num_arguments)
        {
            /* arguments:
             * 0 - mean
             * 1 - standard deviation
             * 2 - count
             * */
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            Parallel.For(0, N, i => 
            {
                int rands_count = (int)((float)parameters[2][i].Value);
                float[] rands = new float[rands_count];
                gpu.Lock();
                gpu_rand.GenerateNormal(rands, (float)parameters[0][i].Value, (float)parameters[1][i].Value, rands_count);
                gpu.Synchronize();
                gpu.Unlock();
                ResultStruct rs_temp = new ResultStruct();
                rs_temp.Cols = 1;
                rs_temp.Rows = rands_count;
                rs_temp.Type = ResultEnum.Float | ResultEnum.Array;
                rs_temp.Value = rands;
                res[i] = rs_temp;
            });

            return res;
        }

        private List<ResultStruct> VBAFn_MCWiener(int N, List<List<ResultStruct>> parameters, int num_arguments, bool call = true)
        {
           /* parameters:
            * 0 => spot price (S)
            * 1 => time to maturity (T-t)
            * 2 => risk-free rate (r)
            * 3 => volatility (sigma)
            * 4 => time_steps (N_t) [unit]
            * 5 => strike_price (K) [applies only for some style of options]
            * 6 => style [unit]
            * */

            /* style types:
             * 0 => European
             * 1 => American
             * 2 => Asian
             * 3 => Lookback
             * 4 => Barrier
             * */

            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);
            byte style = (byte)((float)parameters[6][0].Value);
            Stopwatch sw = new Stopwatch();
            sw.Start();

            Parallel.For(0, N, i => 
            {
                int num_parallels = parameters[0][i].Cols * parameters[0][i].Rows;
                float[] final_prices = new float[num_parallels];
                gpu.Lock();

                float[] dev_GaussianRandoms = gpu.Allocate<float>(num_parallels * (int)((float)parameters[4][i].Value));
                gpu_rand.GenerateNormal(dev_GaussianRandoms, 0, 1, num_parallels * (int)((float)parameters[4][i].Value));
                
                float[] dev_SpotPrice = gpu.CopyToDevice<float>(parameters[0][i].Value as float[]);
                float[] dev_TimeToMaturity = gpu.CopyToDevice<float>(parameters[1][i].Value as float[]);
                float[] dev_RiskFreeRate = gpu.CopyToDevice<float>(parameters[2][i].Value as float[]);
                float[] dev_Volatility = gpu.CopyToDevice<float>(parameters[3][i].Value as float[]);
                float[] dev_StrikePrice = gpu.CopyToDevice<float>(parameters[5][i].Value as float[]);

                //compute the paths
                switch (style)
                {
                    case 0:
                        gpu.Launch(num_parallels, 1, "vba_op_monte_carlo_wiener_european", dev_SpotPrice, dev_TimeToMaturity, dev_RiskFreeRate, dev_Volatility, (int)((float)parameters[4][i].Value), dev_GaussianRandoms, dev_StrikePrice, call ? 0 : 1);
                        break;
                }
                
                gpu.Synchronize();

                //get final prices
                gpu.CopyFromDevice<float>(dev_SpotPrice, final_prices);

                gpu.Free(dev_RiskFreeRate);
                gpu.Free(dev_SpotPrice);
                gpu.Free(dev_TimeToMaturity);
                gpu.Free(dev_Volatility);
                gpu.Free(dev_GaussianRandoms);
                gpu.Free(dev_StrikePrice);
                gpu.Unlock();
            });
            sw.Stop();

            return res;
        }

        private List<ResultStruct> VBAFn_StutzerIndex(int N, List<List<ResultStruct>> parameters, int num_arguments)
        {
            /* parameters:
             * 0 => Instrument returns
             * 1 => Benchmark returns
             * 2 => Search accuracy (how many theta samples to search)
             * */

            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            //TODO: figure out how to reduce all the threads and store only the maximum value in one constant. Requires some shared memory between all GPU threads

            Parallel.For(0, N, i => 
            { 
                //get information statistic (excess return divided by variance) and use its negative as a starting point for theta
                float mean_excess_return = (parameters[0][i].Value as float[])[0] - (parameters[1][i].Value as float[])[0];
                float variance = 0;
                float delta = 0;

                for (int j = 1; j < parameters[0][i].Cols * parameters[0][i].Rows; j++)
                {
                    delta = (parameters[0][i].Value as float[])[j] - (parameters[1][i].Value as float[])[j] - mean_excess_return;
                    mean_excess_return += (((parameters[0][i].Value as float[])[j] - (parameters[1][i].Value as float[])[j]) - mean_excess_return) / (j + 1);
                    variance += delta * ((parameters[0][i].Value as float[])[j] - (parameters[1][i].Value as float[])[j] - mean_excess_return);
                }
                variance /= (parameters[0][i].Cols * parameters[0][i].Rows);

                //generate L theta samples (theta should be negative)
                int num_theta = (int)(float)(parameters[2][i].Value);
                float[] theta_arr = new float[num_theta];
                for (int j = 0; j < num_theta; j++)
                {
                    theta_arr[j] = -1 * ((mean_excess_return / variance) + ((100 * (j / num_theta)) - 50));
                }

                //send the L samples along with data to GPU
                gpu.Lock();
                float[] dev_A = gpu.CopyToDevice<float>(parameters[0][i].Value as float[]);
                float[] dev_B = gpu.CopyToDevice<float>(parameters[1][i].Value as float[]);
                float[] dev_C = gpu.CopyToDevice<float>(theta_arr);

                //calculate all L theta possibilities and find the maximum theta (on GPU)
                gpu.Launch(theta_arr.Length, 1, "vba_op_stutzer_index", dev_A, dev_B, dev_C);
                gpu.Synchronize();

                float[] indices = new float[theta_arr.Length];
                gpu.CopyFromDevice<float>(dev_C, indices);
                gpu.Unlock();

                //find maximum, NOTE: this code can be massively improved by making the GPU do a reduce operation.
                float max = -float.MaxValue;
                for (int j = 0; j < indices.Length; j++)
                    if (indices[j] > max)
                    {
                        max = indices[j];
                    }

                //set the result to be the index
                ResultStruct rs_temp = new ResultStruct();
                rs_temp.Cols = 1;
                rs_temp.Rows = 1;
                rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
                rs_temp.Value = (Math.Abs(mean_excess_return) / mean_excess_return) * Math.Sqrt(2 * max);
                res[i] = rs_temp;
            });

            return res;
        }

        private List<ResultStruct> VBAFn_StutzerIndexOptimalWeights(int N, List<List<ResultStruct>> parameters, int num_arguments)
        {
            /* parameters:
             * 0 => Portfolio returns (column = instrument, row = time index)
             * 1 => Benchmark returns
             * 2 => Search accuracy (how many theta samples to search)
             * */
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            Parallel.For(0, N, i => 
            {
                //generate list of weight permutations
                //generate list of theta permutations
                //copy data to GPU
                //for each weight permutations
                    //copy weights to GPU
                    //run simulation on GPU
                    //find the maximum index value
                    //if maximum index value is greater than global maximum
                        //store permutation and index value
            });

            return res;
        }

        #endregion

        #region Compute launcher
        public List<ResultStruct> Compute(string operation, List<List<ResultStruct>> parameters, int N, int num_arguments)
        {
            List<ResultStruct> res;
            switch (operation)
            {
                case ">":
                case "<":
                case "^":
                case "*":
                case "/":
                case "-":
                case "+":
                    return FnA_Operations(N, parameters, num_arguments, operation);
                case "RAND":
                    return Fn_Random(N); //uniform random number
                case "SUM":
                    return Fn_Sum(N, parameters, num_arguments);
                //case "PRODUCT":
                //    return Fn_Prod(N, parameters, num_arguments);
                case "POWER":
                    return Fn_Power(N, parameters, num_arguments);
                //case "SQRT":
                //    return Fn_Power(N, parameters);
                case "TRANSPOSE":
                    //because it's a transpose, the resultant area is actually swapped (m -> n, n -> m)
                    return FnA_Transpose(N, parameters, num_arguments);
                case "MMULT":
                    //NOTE: 2. Cuda can do transpose, some way to let it do it in the matrix mult instead of doing it manually and then matrix multing.
                    return Fn_Mmult(N, parameters, num_arguments);
                case "VLOOKUP":
                    return Fn_Vlookup(N, parameters, num_arguments);
                case "SUMPRODUCT":
                    return Fn_Sumproduct(N, parameters, num_arguments);
                case "IFERROR":
                    return Fn_IfError(N, parameters, num_arguments);
                case "CONCATENATE":
                    return Fn_Concatenate(N, parameters, num_arguments);
                case "AVERAGE":
                    return Fn_Average(N, parameters, num_arguments);
                case "VAR.P": case "VARP":
                    //convert to 1-D list
                    return FnA_Variance(N, parameters, num_arguments);
                case "VAR.S": case "VAR":
                    return FnA_Variance(N, parameters, num_arguments, false);
                //case "STDEV.P": case "STDEVP":
                //    //square root of VARP
                //    res = FnA_Variance(N, parameters, num_arguments);
                //    (res[0] as ArrayList)[0] = (float)Math.Sqrt((float)(res[0] as ArrayList)[0]);
                //    return res;
                //case "STDEV.S": case "STDEV":
                //    res = FnA_Variance(N, parameters, num_arguments, false);
                //    res[0].Value = (float)Math.Sqrt((float)res[0].Value);
                //    return res;
                default:
                    //mark as error
                    res = new List<ResultStruct>();
                    res.AddRange(new ResultStruct[N]);
                    Parallel.For(0, N, i => 
                    {
                        ResultStruct rs_temp = new ResultStruct();
                        rs_temp.Type = ResultEnum.Error | ResultEnum.String;
                        rs_temp.Value = "#@@CU_1";
                        res[i] = rs_temp;
                    });
                    return res;
            }
        }
        #endregion

        #region Computation functions

        private List<ResultStruct> Fn_IfError(int N, List<List<ResultStruct>> parameters, int num_parameters)
        {
            List<ResultStruct> res = new List<ResultStruct>();
            res.AddRange(new ResultStruct[N]);

            Parallel.For(0, N, i =>
            {
                ResultStruct rs_temp = (ResultStruct)(parameters[0])[i];
                if ((rs_temp.Type & ResultEnum.Error) == ResultEnum.Error)
                    res[i] = (parameters[1])[i];
                else
                    res[i] = (parameters[0])[i];
            });

            return res;
        }

        private List<ResultStruct> FnA_Operations(int N, List<List<ResultStruct>> parameters, int num_parameters, string operation)
        {
            //NB: num_parameters should always be 2 due to the operations (+, -, *, /, etc)
            List<ResultStruct> res = new List<ResultStruct>();
            res.AddRange(new ResultStruct[N]);

            //if both left and right are units, do all of them in one go            
            if ((((ResultStruct)parameters[0][0]).Type & ((ResultStruct)parameters[1][0]).Type & ResultEnum.Unit) == ResultEnum.Unit)
            {
                //use the GPU
                float[] A = new float[N];
                float[] B = new float[N];
                Parallel.For(0, N, i =>
                {
                    A[i] = (float)((ResultStruct)parameters[0][i]).Value;
                    B[i] = (float)((ResultStruct)parameters[1][i]).Value;
                });
                gpu.Lock();
                float[] dev_A = gpu.CopyToDevice<float>(A);
                float[] dev_B = gpu.CopyToDevice<float>(B);
                switch (operation)
                {
                    case "+":
                        gpu_blas.AXPY(1, dev_A, dev_B);
                        break;
                    case "-":
                        gpu_blas.SWAP(dev_A, dev_B);
                        gpu_blas.AXPY(-1, dev_A, dev_B);
                        break;
                    case "*":
                        gpu.Launch(N, 1, "op_mul", dev_A, dev_B);
                        break;
                    case "/":
                        gpu.Launch(N, 1, "op_div", dev_A, dev_B);
                        break;
                    case "^":
                        gpu.Launch(N, 1, "op_exp", dev_A, dev_B);
                        break;
                    default:
                        break;
                }
                gpu.Synchronize();
                gpu.CopyFromDevice(dev_B, B);
                gpu.Free(dev_A);
                gpu.Free(dev_B);
                gpu.Unlock();
                Parallel.For(0, N, i =>
                {
                    ResultStruct temp_res = new ResultStruct();
                    temp_res.Value = B[i];
                    temp_res.Cols = 1;
                    temp_res.Rows = 1;
                    temp_res.Type = ResultEnum.Float | ResultEnum.Unit;
                    res[i] = temp_res;
                });
                
            }
            else if ((((ResultStruct)parameters[0][0]).Type & ~((ResultStruct)parameters[1][0]).Type & ResultEnum.Unit) == ResultEnum.Unit) //iff left is unit
            {
                Parallel.For(0, N, i => 
                {
                    ResultStruct rs_left = (ResultStruct)(parameters[0])[i];
                    float[] B;
                    if ((((ResultStruct)parameters[1][i]).Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                        B = RangeRepository.RangeRepositoryStore[((ResultStruct)parameters[1][i]).Value.ToString()].Values as float[];
                    else
                        B = ((ResultStruct)parameters[1][i]).Value as float[];
                    
                    gpu.Lock();
                    float[] dev_B = gpu.CopyToDevice<float>(B);
                    switch (operation)
                    {
                        case "+":
                            gpu.Launch(B.Length, 1, "op_onesided_add", (float)((ResultStruct)(parameters[0])[i]).Value, dev_B);
                            break;
                        case "-":
                            gpu.Launch(B.Length, 1, "op_onesided_sub", (float)((ResultStruct)(parameters[0])[i]).Value, dev_B, false);
                            break;
                        case "^":
                            gpu.Launch(B.Length, 1, "op_onesided_exp", (float)((ResultStruct)(parameters[0])[i]).Value, dev_B, false);
                            break;
                        case "*":
                            gpu_blas.SCAL((float)((ResultStruct)(parameters[0])[i]).Value, dev_B);
                            break;
                        case "/":
                            gpu.Launch(B.Length, 1, "op_onesided_div", (float)((ResultStruct)(parameters[0])[i]).Value, dev_B, false);
                            break;
                        default:
                            break;
                    }
                    gpu.Synchronize();
                    gpu.CopyFromDevice(dev_B, B);
                    gpu.Free(dev_B);
                    gpu.Unlock();
                    ResultStruct temp_res = new ResultStruct();
                    temp_res.Value = B;
                    temp_res.Cols = rs_left.Cols;
                    temp_res.Rows = rs_left.Rows;
                    temp_res.Type = ResultEnum.Float | ResultEnum.Array;
                    res[i] = temp_res;
                });
            }
            else if ((~((ResultStruct)(parameters[0])[0]).Type & ((ResultStruct)(parameters[1])[0]).Type & ResultEnum.Unit) == ResultEnum.Unit) //iff right is unit
            {
                Parallel.For(0, N, i =>
                {
                    ResultStruct rs_right = (ResultStruct)parameters[1][i];
                    float[] B;
                    if ((((ResultStruct)(parameters[0])[i]).Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                        B = RangeRepository.RangeRepositoryStore[((ResultStruct)(parameters[0])[i]).Value.ToString()].Values as float[];
                    else
                        B = ((ResultStruct)(parameters[0])[i]).Value as float[];
                    gpu.Lock();
                    float[] dev_B = gpu.CopyToDevice<float>(B);
                    switch (operation)
                    {
                        case "+":
                            gpu.Launch(B.Length, 1, "op_onesided_add", (float)((ResultStruct)(parameters[1])[i]).Value, dev_B);
                            break;
                        case "-":
                            gpu.Launch(B.Length, 1, "op_onesided_sub", (float)((ResultStruct)(parameters[1])[i]).Value, dev_B, true);
                            break;
                        case "^":
                            gpu.Launch(B.Length, 1, "op_onesided_exp", (float)((ResultStruct)(parameters[1])[i]).Value, dev_B, true);
                            break;
                        case "*":
                            gpu_blas.SCAL((float)((ResultStruct)(parameters[1])[i]).Value, dev_B);
                            break;
                        case "/":
                            gpu.Launch(B.Length, 1, "op_onesided_div", (float)((ResultStruct)(parameters[1])[i]).Value, dev_B, true);
                            break;
                        default:
                            break;
                    }
                    gpu.Synchronize();
                    gpu.CopyFromDevice(dev_B, B);
                    gpu.Free(dev_B);
                    gpu.Unlock();
                    ResultStruct temp_res = new ResultStruct();
                    temp_res.Value = B;
                    temp_res.Cols = rs_right.Cols;
                    temp_res.Rows = rs_right.Rows;
                    temp_res.Type = ResultEnum.Float | ResultEnum.Unit;
                    res[i] = temp_res;
                });
            }
            else //either both arrays or both mapped ranges
            {
                Parallel.For(0, N, i => 
                {
                    ResultStruct rs_left = (ResultStruct)(parameters[0])[0];
                    ResultStruct rs_right = (ResultStruct)(parameters[1])[0];
                    rs_left = (ResultStruct)(parameters[0])[i];
                    rs_right = (ResultStruct)(parameters[1])[i];
                    int left_size = rs_left.Rows * rs_left.Cols;
                    int right_size = rs_right.Rows * rs_right.Cols;
                    int new_dim_row = Math.Min(rs_left.Rows, rs_right.Rows);
                    int new_dim_col = Math.Min(rs_left.Cols, rs_right.Cols);
                    float[] A = new float[new_dim_col * new_dim_row];
                    float[] B = new float[new_dim_col * new_dim_row];

                    if ((((ResultStruct)(parameters[0])[i]).Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                        for (int c = 0; c < new_dim_col; c++)
                            for (int r = 0; r < new_dim_row; r++)
                                A[c * new_dim_row + r] = (RangeRepository.RangeRepositoryStore[((ResultStruct)(parameters[0])[i]).Value.ToString()].Values as float[])[c * rs_left.Rows + r];
                    else
                        for (int c = 0; c < new_dim_col; c++)
                            for (int r = 0; r < new_dim_row; r++)
                                A[c * new_dim_row + r] = (((ResultStruct)(parameters[0])[i]).Value as float[])[c * rs_left.Rows + r];
                    if ((((ResultStruct)(parameters[1])[i]).Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                        for (int c = 0; c < new_dim_col; c++)
                            for (int r = 0; r < new_dim_row; r++)
                                B[c * new_dim_row + r] = (RangeRepository.RangeRepositoryStore[((ResultStruct)(parameters[1])[i]).Value.ToString()].Values as float[])[c * rs_right.Rows + r];
                    else
                        for (int c = 0; c < new_dim_col; c++)
                            for (int r = 0; r < new_dim_row; r++)
                                B[c * new_dim_row + r] = (((ResultStruct)(parameters[1])[i]).Value as float[])[c * rs_right.Rows + r];

                    gpu.Lock();
                    float[] dev_A = gpu.CopyToDevice<float>(A);
                    float[] dev_B = gpu.CopyToDevice<float>(B);
                    switch (operation)
                    {
                        case "+":
                            gpu_blas.AXPY(1, dev_A, dev_B);
                            break;
                        case "-":
                            gpu_blas.SWAP(dev_A, dev_B);
                            gpu_blas.AXPY(-1, dev_A, dev_B);
                            break;
                        case "*":
                            gpu.Launch(N, 1, "op_mul", dev_A, dev_B);
                            break;
                        case "/":
                            gpu.Launch(N, 1, "op_div", dev_A, dev_B);
                            break;
                        case "^":
                            gpu.Launch(N, 1, "op_exp", dev_A, dev_B);
                            break;
                        default:
                            break;
                    }
                    gpu.Synchronize();
                    gpu.CopyFromDevice(dev_B, B);
                    gpu.Free(dev_A);
                    gpu.Free(dev_B);
                    gpu.Unlock();
                    ResultStruct temp_res = new ResultStruct();
                    if (left_size > right_size)
                    {
                        A = new float[rs_left.Rows * rs_left.Cols];
                        temp_res.Cols = rs_left.Cols;
                        temp_res.Rows = rs_left.Rows;
                        for (int c = 0; c < rs_left.Cols; c++)
                            for (int r = 0; r < rs_left.Rows; r++)
                                if (c + 1 > new_dim_col || r + 1 > new_dim_row)
                                    A[c * rs_left.Rows + r] = float.NaN;
                                else
                                    A[c * rs_left.Rows + r] = B[c * new_dim_row + r];
                        temp_res.Value = A;
                    }
                    else if (right_size > left_size)
                    {
                        temp_res.Cols = rs_right.Cols;
                        temp_res.Rows = rs_right.Rows;
                        A = new float[rs_right.Rows * rs_right.Cols];
                        for (int c = 0; c < rs_right.Cols; c++)
                            for (int r = 0; r < rs_right.Rows; r++)
                                if (c + 1 > new_dim_col || r + 1 > new_dim_row)
                                    A[c * rs_right.Rows + r] = float.NaN;
                                else
                                    A[c * rs_right.Rows + r] = B[c * rs_right.Rows + r];
                        temp_res.Value = A;
                    }
                    else
                    {
                        temp_res.Cols = rs_left.Cols;
                        temp_res.Rows = rs_left.Rows;
                        temp_res.Value = B;
                    }
                    temp_res.Type = ResultEnum.Float | ResultEnum.Array;
                    res[i] = temp_res;
                });
            }
            return res;
        }

        private List<ResultStruct> Fn_Concatenate(int N, List<List<ResultStruct>> parameters, int num_parameters)
        {
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);
            Parallel.For(0, N, i => 
            {
                StringBuilder sb = new StringBuilder();
                for (int j = 0; j < num_parameters; j++)
                {
                    ResultStruct rs_temp = (ResultStruct)(parameters[j])[i];
                    if ((rs_temp.Type & ResultEnum.Empty) == ResultEnum.Empty) continue;
                    if ((rs_temp.Type & ResultEnum.Unit) != ResultEnum.Unit)
                    {
                        //throw error, concat can only take units, not ranges
                    }
                    sb.Append(rs_temp.Value.ToString());
                }
                ResultStruct output_res = new ResultStruct();
                output_res.Type = ResultEnum.Unit | ResultEnum.String;
                output_res.Value = sb.ToString();
                output_res.Rows = 1;
                output_res.Cols = 1;
                res[i] = output_res;
            });

            return res;
        }

        private List<ResultStruct> Fn_Mmult(int N, List<List<ResultStruct>> parameters, int num_parameters, byte first_op = 0, byte second_op = 0)
        {
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);
            Parallel.For(0, N, i => 
            {
                ResultStruct rs_A = (ResultStruct)(parameters[0])[i];
                ResultStruct rs_B = (ResultStruct)(parameters[1])[i];
                ResultStruct res_store = new ResultStruct();
                float[] dev_A, dev_B, dev_C;
                float[] C = new float[rs_A.Rows * rs_B.Cols];

                gpu.Lock();
                if ((rs_A.Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                    dev_A = gpu.CopyToDevice<float>(RangeRepository.RangeRepositoryStore[rs_A.Value.ToString()].Values as float[]);
                else
                    dev_A = gpu.CopyToDevice<float>(rs_A.Value as float[]);
                if ((rs_B.Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                    dev_B = gpu.CopyToDevice<float>(RangeRepository.RangeRepositoryStore[rs_B.Value.ToString()].Values as float[]);
                else
                    dev_B = gpu.CopyToDevice<float>(rs_B.Value as float[]);
                dev_C = gpu.Allocate<float>(rs_A.Rows * rs_B.Cols);

                gpu_blas.GEMM(rs_A.Rows, rs_A.Cols, rs_B.Cols, 1, dev_A, dev_B, 0, dev_C);
                gpu.CopyFromDevice(dev_C, C);

                gpu.Free(dev_A);
                gpu.Free(dev_B);
                gpu.Free(dev_C);
                gpu.Unlock();

                res_store.Rows = rs_A.Rows;
                res_store.Cols = rs_B.Cols;
                res_store.Type = ResultEnum.Array | ResultEnum.Float;
                res_store.Value = C;
                res[i] = res_store;
            });
            
            return res;
        }

        private List<ResultStruct> Fn_Vlookup(int N, List<List<ResultStruct>> parameters, int num_parameters)
        {
            List<ResultStruct> res = new List<ResultStruct>();
            res.AddRange(new ResultStruct[N]);
            ResultStruct rs_temp = (ResultStruct)(parameters[1])[0]; //this code works assuming that parallelisation is done on unique ranges only
            dynamic search_space;
            int[] sorted_map = null;
            search_space = new string[rs_temp.Rows];

            //if the first parameter is a mapped range
            if ((rs_temp.Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
            {
                RangeStruct mapped_range = RangeRepository.RangeRepositoryStore[rs_temp.Value.ToString()];
                Array.Copy(mapped_range.Values as string[], search_space, rs_temp.Rows);
                sorted_map = mapped_range.Sorted_Map;
            }
            else if ((rs_temp.Type & ResultEnum.Array) == ResultEnum.Array)
            {
                for (int i = 0; i < rs_temp.Rows; i++)
                    search_space[i] = (rs_temp.Value as string[])[i];

                //sort the search space
                sorted_map = new int[rs_temp.Rows];
                for (int i = 0; i < sorted_map.Length; i++)
                    sorted_map[i] = i;

                Array.Sort(search_space, sorted_map, StringComparer.Ordinal);
            }

            int[] results = new int[N];

            //for each parallel signature n in N
            Parallel.For(0, N, i => 
            {
                bool range_lookup = true;

                //if the 4th parameter exists and is true
                if (num_parameters == 4)
                    if (((ResultStruct)(parameters[3])[i]).Value.ToString() == "FALSE") //NOTE: when false, only exact matches are searched for. When true, it returns the first value less than the searched term
                        range_lookup = false;

                ResultStruct rs_search_value = (ResultStruct)(parameters[0])[i];
                ResultStruct output_res = new ResultStruct();

                int result_column = Convert.ToInt32(((ResultStruct)(parameters[2])[i]).Value); //get the 3rd parameter (which column to get result from)
                if (result_column >= 1 && rs_temp.Cols >= result_column)
                {
                    //search through the 2nd parameter's 1st column for the 1st parameter's value
                    results[i] = Tools.BinarySearch<string>(search_space, rs_search_value.Value.ToString(), sorted_map);

                    //deal with not-found cases
                    if (results[i] < 0)
                    {
                        if (range_lookup) //if approximate
                        {
                            if (~results[i] == results.Length + 1)
                            {
                                //#N/A error here
                                output_res.Type = ResultEnum.Error | ResultEnum.String;
                                output_res.Value = "#N/A";
                            }
                            else
                            {
                                output_res.Type = ResultEnum.Unit | ResultEnum.String;
                                results[i] = ~results[i] - 1;
                            }
                        }
                        else //if exact
                        {
                            //#N/A error here
                            output_res.Type = ResultEnum.Error | ResultEnum.String;
                            output_res.Value = "#N/A";
                        }
                    }
                }
                else
                {
                    results[i] = -1; //an error marker for later
                    if (result_column < 1)
                    {
                        output_res.Type = ResultEnum.Error | ResultEnum.String;
                        output_res.Value = "#VALUE!";
                        
                    }
                    else if (result_column > rs_temp.Cols)
                    {
                        output_res.Type = ResultEnum.Error | ResultEnum.String;
                        output_res.Value = "#REF!";
                    }
                }

                //get values from column
                if (results[i] >= 0)
                {
                    output_res.Cols = 1; //vlookups can only be individual results (vlookup by definition only returns 1 row)
                    output_res.Rows = 1;
                    if ((rs_temp.Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                    {
                        //since the cell types can be mixed for a column (i.e. some can be integers, others strings) we need to look at each value specifically
                        int value_pos = rs_temp.Rows * (result_column - 1) + sorted_map[results[i]];
                        float value_out;
                        RangeStruct mapped_range = RangeRepository.RangeRepositoryStore[rs_temp.Value.ToString()];
                        if (float.TryParse((mapped_range.Values as string[])[value_pos],out value_out))
                        {
                            output_res.Type = ResultEnum.Unit | ResultEnum.Float;
                            output_res.Value = value_out;
                        }
                        else
                        {
                            output_res.Type = ResultEnum.Unit | ResultEnum.String;
                            output_res.Value = (mapped_range.Values as string[])[value_pos].ToString();
                        }
                    }
                    else if ((rs_temp.Type & ResultEnum.Array) == ResultEnum.Array)
                        output_res.Value = (parameters[1])[rs_temp.Rows * (result_column - 1) + sorted_map[results[i]]];
                    else
                        throw new ApplicationException("Error in VLOOKUP, input parameter is not an array or mapped range!");
                }
                res[i] = output_res;
            });
            return res;
        }

        private static List<ResultStruct> FnA_Transpose(int N, List<List<ResultStruct>> parameters, int num_parameters)
        {
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);
            Parallel.For(0, N, i => 
            {
                ResultStruct res_store = (ResultStruct)((parameters[0])[i]);
                float[] A = new float[res_store.Cols * res_store.Rows];
                int temp = res_store.Cols;
                res_store.Cols = res_store.Rows;
                res_store.Rows = temp;

                for (int j = 0; j < res_store.Cols * res_store.Rows; j++)
			    {
			        A[(int)(j % res_store.Cols) * res_store.Rows + (int)(j / res_store.Cols)] = (res_store.Value as float[])[j];
			    }
                res_store.Value = A;
                res[i] = res_store;
            });
            
            return res;
        }

        private static List<ResultStruct> Fn_Sumproduct(int N, List<List<ResultStruct>> parameters, int num_parameters)
        {
            //SUMPRODUCT first multiplies all parameters and then sums them. The answer is only one number and if it's an array, just replicated.
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            //for each n in N
            Parallel.For(0, N, i => 
            {
                bool is_valid = true;
                //if all parameters are the same size (row & col)
                for (int j = 1; j < num_parameters; j++)
                    if ((((ResultStruct)(parameters[0])[i]).Type != ((ResultStruct)(parameters[j])[i]).Type) ||
                        (((ResultStruct)(parameters[0])[i]).Rows != ((ResultStruct)(parameters[j])[i]).Rows) ||
                        (((ResultStruct)(parameters[0])[i]).Cols != ((ResultStruct)(parameters[j])[i]).Cols))
                    {
                        is_valid = false;
                        break;
                    }
                ResultStruct rs_temp = new ResultStruct();
                //if they're not return a #VALUE! error
                if (!is_valid)
                {
                    rs_temp.Type = ResultEnum.Error;
                    rs_temp.Value = "#VALUE!";
                }
                else
                {
                    int temp_acc = ((ResultStruct)(parameters[0])[i]).Cols * ((ResultStruct)(parameters[0])[i]).Rows;
                    rs_temp.Type = ResultEnum.Float | ResultEnum.Unit;
                    rs_temp.Rows = 1;
                    rs_temp.Cols = 1;
                    //if num_parameters = 2
                    if (num_parameters == 2)
                    {
                        //do a gpu_blas.dot()
                        gpu.Lock();
                        float[] dev_A = gpu.CopyToDevice<float>(((ResultStruct)(parameters[0])[i]).Value as float[]);
                        float[] dev_B = gpu.CopyToDevice<float>(((ResultStruct)(parameters[1])[i]).Value as float[]);
                        rs_temp.Value = gpu_blas.DOT(dev_A, dev_B);
                        gpu.Free(dev_A);
                        gpu.Free(dev_B);
                        gpu.Unlock();        
                    }
                    else
                    {
                        gpu.Lock();
                        float[] dev_B = gpu.CopyToDevice<float>(((ResultStruct)(parameters[0])[i]).Value as float[]);
                        float[] dev_A = gpu.Allocate<float>(temp_acc);
                        //for each parameter > 1
                        for (int j = 1; j < num_parameters; j++)
                        {
                            //do a multiplication
                            gpu.CopyToDevice<float>(((ResultStruct)(parameters[j])[i]).Value as float[], dev_A);
                            gpu.Launch(temp_acc, 1, "op_mul", dev_A, dev_B);
                        }
                        float[] B = new float[temp_acc];
                        gpu.CopyFromDevice<float>(dev_B, B);
                        gpu.Free(dev_B);
                        gpu.Free(dev_A);
                        gpu.Unlock();
                        //return the sum of the final product
                        for (int j = 1; j < temp_acc; j++)
                            B[0] += B[j];
                        rs_temp.Value = B[0];
                    }
                }
                res[i] = rs_temp;
            });
            return res;
        }
        
        private List<ResultStruct> FnA_Variance(int N, List<List<ResultStruct>> parameters, int num_parameters, bool population = true) //TODO: needs fixing
        {
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            Parallel.For(0, N, i => 
            {
                int num = 0;
                float temp_float = 0;
                float mean;
                float[] B;
                float[] dev_A;
                float[] dev_B;
                //get values from all parameters and sum them
                for (int j = 0; j < num_parameters; j++)
                {
                    num += parameters[j][i].Cols * parameters[j][i].Rows;
                    if ((parameters[j][i].Type & ResultEnum.Array) == ResultEnum.Array)
                    {
                        int value_size = parameters[j][i].Cols * parameters[j][i].Rows;
                        B = new float[value_size];
                        for (int k = 0; k < value_size; k++)
                            B[k] = 1;

                        gpu.Lock();
                        dev_A = gpu.CopyToDevice<float>(parameters[j][i].Value as float[]);
                        dev_B = gpu.CopyToDevice<float>(B);
                        temp_float += gpu_blas.DOT(dev_A, dev_B);
                        gpu.Free(dev_A);
                        gpu.Free(dev_B);
                        gpu.Unlock();
                    }
                    else if ((parameters[j][i].Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                    {
                        int value_size = parameters[j][i].Cols * parameters[j][i].Rows;
                        B = new float[value_size];
                        for (int k = 0; k < value_size; k++)
                            B[k] = 1;

                        gpu.Lock();
                        dev_A = gpu.CopyToDevice<float>(RangeRepository.RangeRepositoryStore[parameters[j][i].Value.ToString()].Values as float[]);
                        dev_B = gpu.CopyToDevice<float>(B);
                        temp_float += gpu_blas.DOT(dev_A, dev_B);
                        gpu.Free(dev_A);
                        gpu.Free(dev_B);
                        gpu.Unlock();
                    }
                    else //must be a unit
                        temp_float += (float)parameters[j][i].Value;
                }
                //get the mean
                mean = temp_float / num;

                //get the squared differences
                float[] A = new float[num];
                B = new float[num];
                int counter = 0;
                for (int j = 0; j < num_parameters; j++)
                {
                    int value_size = parameters[j][i].Cols * parameters[j][i].Rows;
                    if ((parameters[j][i].Type & ResultEnum.Array) == ResultEnum.Array)
                        Parallel.For(0, value_size, k => 
                        {
                            A[counter + value_size] = (parameters[j][i].Value as float[])[k];
                        });
                    else if ((parameters[j][i].Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                        Parallel.For(0, value_size, k => 
                        {
                            A[counter + value_size] = (RangeRepository.RangeRepositoryStore[parameters[j][i].Value.ToString()].Values as float[])[k];
                        });
                    else
                        A[counter] = (float)parameters[j][i].Value;
                    counter += value_size;
                }
                for (int j = 0; j < num; j++)
                    B[j] = mean;

                gpu.Lock();
                dev_A = gpu.CopyToDevice<float>(A);
                dev_B = gpu.CopyToDevice<float>(B);
                gpu.Launch(num, 1, "op_sub_squared", dev_A, dev_B);
                gpu.CopyFromDevice<float>(dev_B, B);
                gpu.Free(dev_A);
                gpu.Unlock();

                //get the sum of squares
                for (int j = 0; j < num; j++)
                    A[j] = 1;

                gpu.Lock();
                dev_A = gpu.CopyToDevice<float>(A);
                temp_float = gpu_blas.DOT(B, A);
                gpu.Free(dev_A);
                gpu.Free(dev_B);
                gpu.Unlock();

                temp_float /= population ? num : num - 1;
                ResultStruct temp_res = new ResultStruct();
                temp_res.Cols = 1;
                temp_res.Rows = 1;
                temp_res.Type = ResultEnum.Unit | ResultEnum.Float;
                temp_res.Value = temp_float;
                
                res[i] = temp_res;
            });
            return res;
        }
        
        private List<ResultStruct> Fn_Power(int N, List<List<ResultStruct>> parameters, int num_parameters)
        {
            List<ResultStruct> res = new List<ResultStruct>(N);
            res.AddRange(new ResultStruct[N]);

            Parallel.For(0, N, i => 
            {
                ResultStruct rs_temp = new ResultStruct();

                //if input is string, throw error
                if (((parameters[0][i].Type & ResultEnum.String) == ResultEnum.String) || ((parameters[1][i].Type & ResultEnum.String) == ResultEnum.String))
                {
                    rs_temp.Cols = 1;
                    rs_temp.Rows = 1;
                    rs_temp.Type = ResultEnum.Error | ResultEnum.String;
                    rs_temp.Value = "#VALUE!";
                    return;
                }
                
                //if both units, no need to use GPU
                if (((parameters[0][i].Type & ResultEnum.Unit) == ResultEnum.Unit) && ((parameters[1][i].Type & ResultEnum.Unit) == ResultEnum.Unit))
                {
                    rs_temp.Cols = 1;
                    rs_temp.Rows = 1;
                    rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
                    rs_temp.Value = Math.Pow((float)parameters[0][i].Value, (float)parameters[1][i].Value);
                }
                else
                {
                    float[] arr_left;
                    float[] arr_right;

                    arr_left = parameters[0][i].Value as float[];
                    arr_right = parameters[1][i].Value as float[];
                    rs_temp.Cols = Math.Max(parameters[0][i].Cols, parameters[1][i].Cols);
                    rs_temp.Rows = Math.Max(parameters[0][i].Rows, parameters[1][i].Rows);
                    rs_temp.Type = ResultEnum.Array | ResultEnum.Float;

                    if (arr_left.Length < arr_right.Length)
                    {
                        float[] temp_arr = arr_left;
                        arr_left = new float[arr_right.Length];
                        Array.Copy(temp_arr, arr_left, temp_arr.Length);
                    }
                    else if (arr_left.Length > arr_right.Length)
                    {
                        float[] temp_arr = arr_right;
                        arr_right = new float[arr_left.Length];
                        Array.Copy(temp_arr, arr_right, temp_arr.Length);
                    }

                    gpu.Lock();
                    float[] dev_A = gpu.CopyToDevice<float>(arr_left);
                    float[] dev_B = gpu.CopyToDevice<float>(arr_right);

                    gpu.Launch(N, 1, "op_exp", dev_A, dev_B);
                    gpu.Synchronize();
                    gpu.CopyFromDevice<float>(dev_B, arr_right);

                    gpu.Free(dev_A);
                    gpu.Free(dev_B);
                    gpu.Unlock();

                    rs_temp.Value = arr_right;
                }
            });

            return res;
        }
        
        private List<ResultStruct> Fn_Prod(int N, List<List<ResultStruct>> parameters, int num_arguments)
        {
            List<ResultStruct> output = new List<ResultStruct>(N);
            output.AddRange(new ResultStruct[N]);

            float[] prod_accum = new float[N];
            float[] temp_array = new float[N];
            for (int i = 0; i < N; i++)
                prod_accum[i] = 1;

            gpu.Lock();
            float[] dev_A = gpu.Allocate<float>(N);
            float[] dev_B = gpu.Allocate<float>(N);
            float[] dev_res = gpu.Allocate<float>(N);
            gpu.Unlock();

            for (int j = 0; j < num_arguments; j++)
            {
                //for each n in N
                Parallel.For(0, N, i =>
                {
                    if ((parameters[j][i].Type & ResultEnum.Array) == ResultEnum.Array)
                    {
                        temp_array[i] = 1;
                        for (int k = 0; k < parameters[j][i].Rows * parameters[j][i].Cols; k++)
                            temp_array[i] *= (parameters[j][i].Value as float[])[k];
                    }
                    else if ((parameters[j][i].Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                    {
                        temp_array[i] = 1;
                        for (int k = 0; k < parameters[j][i].Rows * parameters[j][i].Cols; k++)
                            temp_array[i] *= (RangeRepository.RangeRepositoryStore[parameters[j][i].Value.ToString()].Values as float[])[k];
                    }
                    else
                        temp_array[i] = (float)parameters[j][i].Value;
                });

                if (j > 0)
                {
                    gpu.Lock();
                    dev_A = gpu.CopyToDevice<float>(temp_array);
                    dev_B = gpu.CopyToDevice<float>(prod_accum);
                    gpu_blas.GEMV(N, 1, 1, dev_A, dev_B, 0, dev_res); //if this doesn't work, make N the second argument.
                    gpu.Synchronize();
                    gpu.CopyFromDevice<float>(dev_res, prod_accum);
                    gpu.Unlock();
                }

            }
            gpu.Lock();
            gpu.Free(dev_A);
            gpu.Free(dev_B);
            gpu.Free(dev_res);
            gpu.Unlock();
            
            return output;
        }

        private List<ResultStruct> Fn_Sum(int N, List<List<ResultStruct>> parameters, int num_arguments)
        {
            /* Steps:
             * 1 - Add up each individual argument of the SUM using dot product (example, ranges)
             * 2 - Add up all arguments for all parallel equations using GEMV
             * */

            gpu.Lock();
            float[,] X = new float[num_arguments, N];
            float[,] dev_X = gpu.Allocate<float>(X);

            float[] V = new float[num_arguments];
            float[] dev_V = gpu.Allocate<float>(V);

            float[] res = new float[N];
            float[] dev_res = gpu.Allocate<float>(N);
            gpu.Unlock();

            List<ResultStruct> output = new List<ResultStruct>(N);
            output.AddRange(new ResultStruct[N]);
            int[] counter = new int[N];

            //pre-process any ranges
            //for each argument (num_arguments)
            for (int i = 0; i < num_arguments; i++)
            {
                Parallel.For(0, N, j =>
                {
                    ResultStruct param_unit = (ResultStruct)parameters[i][j];
                    //if a range, sum using DOT
                    if ((param_unit.Type & ResultEnum.Array) == ResultEnum.Array || (param_unit.Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                    {
                        int size = param_unit.Rows * param_unit.Cols;
                        float[] Y = new float[size];
                        float temp_res;

                        Parallel.For(0, size, k =>
                        {
                            Y[k] = 1;
                        });

                        gpu.Lock();
                        float[] dev_Y = gpu.Allocate<float>(Y);
                        float[] dev_Z = gpu.Allocate<float>(size);

                        gpu.CopyToDevice(Y, dev_Y); //~ 35 us for 30 elements
                        if ((param_unit.Type & ResultEnum.Array) == ResultEnum.Array)
                            gpu.CopyToDevice(param_unit.Value as float[], dev_Z);
                        else
                            gpu.CopyToDevice(RangeRepository.RangeRepositoryStore[param_unit.Value.ToString()].Values as float[], dev_Z);

                        temp_res = gpu_blas.DOT(dev_Y, dev_Z); //~ 1.3 ms for 30 elements

                        gpu.Free(dev_Y);
                        gpu.Free(dev_Z);
                        gpu.Unlock();

                        X[i, j] = temp_res;
                    }
                    else
                        X[i, j] = (float)param_unit.Value;
                });
                V[i] = 1;
            }

            if (num_arguments * N > 1) //don't do GEMV if we've only got one input
            {
                gpu.Lock();
                //sum all N (independently) using GEMV
                gpu.CopyToDevice(X, dev_X);
                gpu.CopyToDevice(V, dev_V);
                float[] cast_dev_X = gpu.Cast(dev_X, num_arguments * N);

                gpu_blas.GEMV(N, num_arguments, 1, cast_dev_X, dev_V, 0, dev_res);

                gpu.CopyFromDevice(dev_res, res);

                gpu.Free(dev_res);
                gpu.Free(dev_X);
                gpu.Free(dev_V);
                gpu.Unlock();
            }
            else
                res[0] = X[0, 0];

            for (int i = 0; i < N; i++)
            {
                ResultStruct formatted_res = new ResultStruct();
                formatted_res.Value = res[i];
                formatted_res.Rows = 1;
                formatted_res.Cols = 1;
                formatted_res.Type = ResultEnum.Unit | ResultEnum.Float;
                output[i] = formatted_res;
            }

            return output;
        }

        private List<ResultStruct> Fn_Average(int N, List<List<ResultStruct>> parameters, int num_arguments)
        {
            /* Steps:
             * 1 - Add up each individual argument of the SUM using dot product (example, ranges)
             * 2 - Add up all arguments for all parallel equations using GEMV
             * */

            gpu.Lock();
            float[,] X = new float[num_arguments, N];
            float[,] dev_X = gpu.Allocate<float>(X);

            float[] V = new float[num_arguments];
            float[] dev_V = gpu.Allocate<float>(V);

            float[] res = new float[N];
            float[] dev_res = gpu.Allocate<float>(N);
            gpu.Unlock();

            List<ResultStruct> output = new List<ResultStruct>(N);
            output.AddRange(new ResultStruct[N]);
            int[] counter = new int[N];

            //pre-process any ranges
            //for each argument (num_arguments)
            for (int i = 0; i < num_arguments; i++)
            {
                Parallel.For(0, N, j =>
                {
                    ResultStruct param_unit = (ResultStruct)parameters[i][j];
                    //if a range, sum using DOT
                    if ((param_unit.Type & ResultEnum.Array) == ResultEnum.Array || (param_unit.Type & ResultEnum.Mapped_Range) == ResultEnum.Mapped_Range)
                    {
                        int size = param_unit.Rows * param_unit.Cols;
                        float[] Y = new float[size];
                        float temp_res;

                        Parallel.For(0, size, k =>
                        {
                            Y[k] = 1;
                        });

                        gpu.Lock();
                        float[] dev_Y = gpu.Allocate<float>(Y);
                        float[] dev_Z = gpu.Allocate<float>(size);

                        gpu.CopyToDevice(Y, dev_Y); //~ 35 us for 30 elements
                        if ((param_unit.Type & ResultEnum.Array) == ResultEnum.Array)
                            gpu.CopyToDevice(param_unit.Value as float[], dev_Z);
                        else
                            gpu.CopyToDevice(RangeRepository.RangeRepositoryStore[param_unit.Value.ToString()].Values as float[], dev_Z);

                        temp_res = gpu_blas.DOT(dev_Y, dev_Z); //~ 1.3 ms for 30 elements

                        gpu.Free(dev_Y);
                        gpu.Free(dev_Z);
                        gpu.Unlock();

                        X[i, j] = temp_res;
                        counter[j] += size;
                    }
                    else
                    {
                        X[i, j] = (float)param_unit.Value;
                        counter[j]++;
                    }
                });
                V[i] = 1;
            }

            if (num_arguments * N > 1) //don't do GEMV if we've only got one input
            {
                gpu.Lock();
                //sum all N (independently) using GEMV
                gpu.CopyToDevice(X, dev_X);
                gpu.CopyToDevice(V, dev_V);
                float[] cast_dev_X = gpu.Cast(dev_X, num_arguments * N);

                gpu_blas.GEMV(N, num_arguments, 1, cast_dev_X, dev_V, 0, dev_res);

                gpu.CopyFromDevice(dev_res, res);

                gpu.Free(dev_res);
                gpu.Free(dev_X);
                gpu.Free(dev_V);
                gpu.Unlock();
            }
            else
                res[0] = X[0, 0];

            for (int i = 0; i < N; i++)
            {
                ResultStruct formatted_res = new ResultStruct();
                formatted_res.Value = res[i] / counter[i];
                formatted_res.Rows = 1;
                formatted_res.Cols = 1;
                formatted_res.Type = ResultEnum.Unit | ResultEnum.Float;
                output[i] = formatted_res;
            }

            return output;
        }

        private List<ResultStruct> Fn_Random(int N)
        {
            float[] res = new float[N];

            gpu.Lock();
            float[] dev_r = gpu.Allocate<float>(N);

            gpu.Synchronize();
            gpu_rand.SetPseudoRandomGeneratorSeed((ulong)DateTime.Now.Ticks);
            gpu_rand.GenerateUniform(dev_r, N);            
            gpu.Synchronize();
            gpu.CopyFromDevice(dev_r, res);
            gpu.Free(dev_r);
            gpu.Unlock();

            List<ResultStruct> output = new List<ResultStruct>();
            output.AddRange(new ResultStruct[N]);
            Parallel.For(0, N, i => 
            {
                ResultStruct rs_temp = new ResultStruct();
                rs_temp.Type = ResultEnum.Unit | ResultEnum.Float;
                rs_temp.Value = res[i];
                rs_temp.Cols = 1;
                rs_temp.Rows = 1;
                output[i] = rs_temp;
            });
            
            return output;
        }

        #endregion

        #region CUDA functions

        [Cudafy]
        private static void CUDA_bitonic_sort_float(GThread thread, float[] values, int[] swaps)
        {
            //from C:\Users\Bob\Documents\Visual Studio 2013\Projects\Cudafy_Source\Cudafy\3p\cuda.net3.0.0_win\examples\bitonic
            int tid = thread.get_global_id(0);
            float flt_temp;
            int int_tmp;

            thread.SyncThreads();

            if (tid < values.Length)
            {
                for (int k = 2; k <= values.Length; k *= 2)
                {
                    for (int j = k / 2; j > 0; j /= 2)
                    {
                        int ixj = tid ^ j;

                        if (ixj > tid)
                        {
                            if ((tid & k) == 0)
                            {
                                if (values[tid] > values[ixj])
                                {
                                    //swap values
                                    flt_temp = values[tid];
                                    values[tid] = values[ixj];
                                    values[ixj] = flt_temp;
                                    //swap swaps
                                    int_tmp = swaps[tid];
                                    swaps[tid] = swaps[ixj];
                                    swaps[ixj] = int_tmp;
                                }
                            }
                            else
                            {
                                if (values[tid] < values[ixj])
                                {
                                    //swap values
                                    flt_temp = values[tid];
                                    values[tid] = values[ixj];
                                    values[ixj] = flt_temp;
                                    //swap swaps
                                    int_tmp = swaps[tid];
                                    swaps[tid] = swaps[ixj];
                                    swaps[ixj] = int_tmp;
                                }
                            }
                        }
                        thread.SyncThreads();
                    }
                }
            }
        }

        [Cudafy]
        private static void CUDA_search_range(GThread thread, string keyword, string[] search_space, int[] res)
        {
            int tid = thread.get_global_id(0);
            if (tid < search_space.Length)
                if (keyword == search_space[tid])
                {
                    //AtomicFunctions.atomicMin(thread, ref res[0], tid);
                    thread.atomicMin(ref res[0], tid);
                }
        }

        [Cudafy]
        private static void op_sub_squared(GThread thread, float[] a, float[] b)
        {
            int tid = thread.blockIdx.x;
            if (tid < a.Length)
                b[tid] = (float)Math.Pow(a[tid] - b[tid], 2);
        }

        [Cudafy]
        private static void op_mul(GThread thread, float[] a, float[] b)
        {
            int tid = thread.blockIdx.x;
            if (tid < a.Length)
                b[tid] = a[tid] * b[tid];
        }

        [Cudafy]
        private static void op_div(GThread thread, float[] a, float[] b)
        {
            int tid = thread.blockIdx.x;
            if (tid < a.Length)
            {
                if (b[tid] == 0)
                    b[tid] = float.MaxValue;
                else
                    b[tid] = a[tid] / b[tid];
            }
        }

        [Cudafy]
        private static void op_exp(GThread thread, float[] a, float[] b)
        {
            int tid = thread.blockIdx.x;
            if (tid < a.Length)
                b[tid] = (float)Math.Pow(a[tid], b[tid]);
        }
        
        [Cudafy]
        private static void op_onesided_add(GThread thread, float a, float[] b)
        {
            int tid = thread.blockIdx.x;
            if (tid < b.Length)
                b[tid] = b[tid] + a;
        }

        [Cudafy]
        private static void op_onesided_sub(GThread thread, float a, float[] b, bool swap = false)
        {
            int tid = thread.blockIdx.x;
            if (tid < b.Length)
                if (swap)
                    b[tid] = b[tid] - a;
                else
                    b[tid] = a - b[tid];
        }

        [Cudafy]
        private static void op_onesided_exp(GThread thread, float a, float[] b, bool swap = false)
        {
            int tid = thread.blockIdx.x;
            if (tid < b.Length)
                if (swap)
                    b[tid] = (float)Math.Pow(b[tid], a);
                else
                    b[tid] = (float)Math.Pow(a, b[tid]);
        }

        [Cudafy]
        private static void op_onesided_div(GThread thread, float a, float[] b, bool swap = false)
        {
            int tid = thread.blockIdx.x;
            if (tid < b.Length)
                if (swap)
                    b[tid] = b[tid] / a;
                else
                    b[tid] = a / b[tid];
        }

        [Cudafy]
        private static void vba_op_power_transform(GThread thread, float[] a, float lambda, float GM)
        {
            int tid = thread.blockIdx.x;
            if (tid < a.Length)
            {
                if (lambda == 0)
                    a[tid] = (float)Math.Log(a[tid]) * GM;
                else
                    a[tid] = (a[tid] - 1) / (lambda * (float)Math.Pow(GM, lambda - 1));
            }
        }

        [Cudafy]
        private static void vba_op_black_scholes_call(GThread thread, float[] spot_price, float[] maturity_time, float[] strike_price, float[] risk_free_rate, float[] volatility)
        {
            int tid = thread.blockIdx.x;
            if (tid < spot_price.Length)
            {
                float d1 = (1 / (volatility[tid] * (float)Math.Sqrt(maturity_time[tid]))) * (float)(Math.Log(spot_price[tid] / strike_price[tid]) + (risk_free_rate[tid] + Math.Pow(volatility[tid], 2) / 2) * (maturity_time[tid]));
                float d2 = d1 - volatility[tid] * (float)Math.Sqrt(maturity_time[tid]); //NOTE: can be optimised by storing already calculated bits

                float cdf1 = 0;
                float cdf2 = 0;

                GThread.InsertCode("{0} = 0.5 * (1 + erf({1}/sqrt(2.0f)));", cdf1, d1); //in the future the square roots may be precomputed, but currently it's blazing fast.
                GThread.InsertCode("{0} = 0.5 * (1 + erf({1}/sqrt(2.0f)));", cdf2, d2); //important to have explicit floats

                float res = cdf1 * spot_price[tid] - cdf2 * strike_price[tid] * (float)Math.Exp(-1 * risk_free_rate[tid] * maturity_time[tid]);
                spot_price[tid] = res;
            }
        }

        [Cudafy]
        private static void vba_op_black_scholes_put(GThread thread, float[] spot_price, float[] maturity_time, float[] strike_price, float[] risk_free_rate, float[] volatility)
        {
            int tid = thread.blockIdx.x;
            if (tid < spot_price.Length)
            {
                float d1 = (1 / (volatility[tid] * (float)Math.Sqrt(maturity_time[tid]))) * (float)(Math.Log(spot_price[tid] / strike_price[tid]) + (risk_free_rate[tid] + Math.Pow(volatility[tid], 2) / 2) * (maturity_time[tid]));
                float d2 = d1 - volatility[tid] * (float)Math.Sqrt(maturity_time[tid]);

                float cdf1 = 0;
                float cdf2 = 0;

                GThread.InsertCode("{0} = 0.5 * (1 + erf(-1 * {1}/sqrt(2.0f)));", cdf1, d1); //important to have explicit floats
                GThread.InsertCode("{0} = 0.5 * (1 + erf(-1 * {1}/sqrt(2.0f)));", cdf2, d2);
                
                float res = cdf2 * strike_price[tid] * (float)Math.Exp(-1 * risk_free_rate[tid] * maturity_time[tid]) - cdf1 * spot_price[tid];
                spot_price[tid] = res;
            }
        }

        [Cudafy]
        private static void vba_op_monte_carlo_wiener_european(GThread thread, float[] spot_price, float[] maturity_time, float[] risk_free_rate, float[] volatility, int time_samples, float[] gaussianRandoms, float[] strike_price, int call = 0)
        {
            int tid = thread.blockIdx.x;
            if (tid < spot_price.Length)
            {
                float time = 0;
                float payoff = 0;
                float delta_t = maturity_time[tid] / time_samples;
                float price = spot_price[tid];
                int step = 0;
                while (time < maturity_time[tid])
                {
                    time += delta_t;
                    price *= (float)Math.Exp((risk_free_rate[tid] - 0.5 * (float)Math.Pow(volatility[tid], 2)) * delta_t + volatility[tid] * (float)Math.Sqrt(delta_t) * gaussianRandoms[tid * time_samples + step]);
                    step++;
                    if (call == 0)
                        payoff += (float)Math.Exp(-1 * risk_free_rate[tid] * maturity_time[tid]) * (float)Math.Max(price - strike_price[tid], 0);
                    else
                        payoff += (float)Math.Exp(-1 * risk_free_rate[tid] * maturity_time[tid]) * (float)Math.Max(strike_price[tid] - price, 0);
                }
                spot_price[tid] = payoff / time_samples; //mean of payoff for European option output
            }
        }

        [Cudafy]
        private static void vba_op_stutzer_index(GThread thread, float[] instrument_ret, float[] benchmark_ret, float[] theta)
        {
            int tid = thread.blockIdx.x;            

            if (tid < theta.Length)
            {
                float sum = 0;

                for (int i = 0; i < instrument_ret.Length; i++)
                {
                    sum += (float)Math.Exp(theta[tid] * (instrument_ret[i] - benchmark_ret[i]));
                }

                theta[tid] = -1 * (float)Math.Log(sum / instrument_ret.Length);

                //NOTE: this can be improved by using a reduce function to atomically save the maximum value and position
            }
        }

        #endregion
    }
}
