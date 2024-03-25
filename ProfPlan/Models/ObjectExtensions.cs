using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlan.Models
{
    public static class ObjectExtensions
    {
        public static T? ToNullable<T>(this object value) where T : struct
        {
            if (value == DBNull.Value || value == null)
            {
                return null;
            }
            else
            {
                try
                {
                    return (T)Convert.ChangeType(value, typeof(T));
                }
                catch (InvalidCastException)
                {
                    return null;
                }
            }
        }
    }
}
