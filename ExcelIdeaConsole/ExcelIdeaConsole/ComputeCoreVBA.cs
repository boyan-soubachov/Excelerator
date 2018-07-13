/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * 
 * This class acts as the interface for VBA functions called and their computation with CUDA/CPU.
 * 
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelIdea
{
    class ComputeCoreVBA
    {
        public ComputeCoreVBA()
        {
            if (!Program.CUDA_Core.core_up)
            {
                throw new SystemException("=> ERROR: When initialising VBA compute core, CUDA core is not up.");
            }
        }

        public List<ResultStruct> ComputeOperator(string operation, List<List<ResultStruct>> parameters, int N, int num_arguments)
        {

            switch (operation)
            {
                case "RANDN":
                    return Program.CUDA_Core.VBACompute("RANDN", parameters, N, num_arguments);
                case "BLACK-SCHOLES-PUT":
                    return Program.CUDA_Core.VBACompute("BLACK-SCHOLES-PUT", parameters, N, num_arguments);
                case "BLACK-SCHOLES-CALL":
                    return Program.CUDA_Core.VBACompute("BLACK-SCHOLES-CALL", parameters, N, num_arguments);
                case "MC-WIENER-CALL-PRICE":
                    return Program.CUDA_Core.VBACompute("MC-WIENER-CALL-PRICE", parameters, N, num_arguments);
                case "POWER-TRANSFORM":
                    return Program.CUDA_Core.VBACompute("POWER-TRANSFORM", parameters, N, num_arguments);
                case "RANDOM":
                    return Program.CUDA_Core.Compute("RAND", parameters, N, num_arguments);
                default:
                    break;
            }
            return null;
        }
    }
}
