using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace NpoiMadeSimple.Excel.Typed
{
    public class Sheet<T> : Sheet where T : class, new()
    {
        Dictionary<Type, Action<T, PropertyInfo, string>> setPropertyValue = new Dictionary<Type, Action<T, PropertyInfo, string>>{
            {typeof(bool)   ,(item, prop, value)=> { prop.SetValue(item, bool.Parse(value) ); }  },
            {typeof(byte)   ,(item, prop, value)=> { prop.SetValue(item, byte.Parse(value) ); }  },
            {typeof(sbyte)  ,(item, prop, value)=> { prop.SetValue(item, sbyte.Parse(value) ); }  },
            {typeof(char)   ,(item, prop, value)=> { prop.SetValue(item, char.Parse(value) ); }  },
            {typeof(decimal),(item, prop, value)=> { prop.SetValue(item, decimal.Parse(value) ); }  },
            {typeof(double) ,(item, prop, value)=> { prop.SetValue(item, double.Parse(value) ); }  },
            {typeof(float)  ,(item, prop, value)=> { prop.SetValue(item, float.Parse(value) ); }  },
            {typeof(int)    ,(item, prop, value)=> { prop.SetValue(item, int.Parse(value) ); }  },
            {typeof(uint)   ,(item, prop, value)=> { prop.SetValue(item, uint.Parse(value) ); }  },
            {typeof(long)   ,(item, prop, value)=> { prop.SetValue(item, long.Parse(value) ); }  },
            {typeof(ulong)  ,(item, prop, value)=> { prop.SetValue(item, ulong.Parse(value) ); }  },
            {typeof(short)  ,(item, prop, value)=> { prop.SetValue(item, short.Parse(value) ); }  },
            {typeof(ushort) ,(item, prop, value)=> { prop.SetValue(item, ushort.Parse(value) ); }  },
            {typeof(DateTime) ,(item, prop, value)=> { prop.SetValue(item, DateTime.Parse(value)); }  },
            {typeof(string) ,(item, prop, value)=> { prop.SetValue(item, value); }  }

        };

        public Sheet(ISheet sheet) : this(sheet, 0) { }

        public Sheet(ISheet sheet, int headerRowNumber) : base(sheet, headerRowNumber) { }

        public IEnumerable<T> AllItens()
        {
            Type type = typeof(T);

            foreach (Row row in this)
            {
                T item = new T();
                foreach (PropertyInfo prop in type.GetProperties())
                {
                    var att = prop.GetCustomAttribute<SheetColumnAttribute>();
                    if (att == null)
                        continue;

                    var cell = row[att.GetName()];

                    try
                    {
                        setPropertyValue[prop.PropertyType](item, prop, cell.ToString());
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error parsing value \"{cell}\" at \"{cell.GetAddress()}\" as {prop.PropertyType}", ex);
                    }
                }
                yield return item;
            }
        }

    }
}
