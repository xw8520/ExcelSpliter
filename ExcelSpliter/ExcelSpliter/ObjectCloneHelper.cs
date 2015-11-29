using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSpliter
{
    public class ObjectCloneHelper
    {
        public static T CloneOf<T>(T serializableObject)
        {
            object objCopy = null;
            MemoryStream stream = new MemoryStream();
            BinaryFormatter binFormatter = new BinaryFormatter();
            binFormatter.Serialize(stream, serializableObject);
            stream.Position = 0;
            objCopy = (T)binFormatter.Deserialize(stream);
            stream.Close();
            return (T)objCopy;
        }
    }
}
