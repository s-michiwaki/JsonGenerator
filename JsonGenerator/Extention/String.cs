using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public static class StringExtention
{
    public static bool IsNullOrEmpty(this string value)
    {
        return value == null || value == "";
    }
}
