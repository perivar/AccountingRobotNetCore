using System.Collections;
using System.Collections.Generic;

namespace AccountingServices.Helpers
{
    public interface IMyConfiguration
    {
         string GetValue(string key);         
    }
}