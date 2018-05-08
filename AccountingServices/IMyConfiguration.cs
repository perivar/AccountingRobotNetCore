using System.Collections;
using System.Collections.Generic;

namespace AccountingServices
{
    public interface IMyConfiguration
    {
         string GetValue(string key);         
    }
}