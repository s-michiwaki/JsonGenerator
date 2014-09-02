using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonGenerator.Repository
{
    public class AnyActionResult
    {
        public RepositoryResultCode ResultCode { get; private set; }

        public string Message { get; private set; }

        public static AnyActionResult Success(string message = "")
        {
            return new AnyActionResult
            {
                ResultCode = RepositoryResultCode.Success,
                Message = message,
            };
        }

        public static AnyActionResult Failed(string message = "")
        {
            return new AnyActionResult
            {
                ResultCode = RepositoryResultCode.Failed,
                Message = message,
            };
        }
    }

    public enum RepositoryResultCode
    {
        Success = 1,
        Failed = 2,
    }
}
