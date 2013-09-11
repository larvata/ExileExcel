using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExileExcel
{
    public interface IExileData
    {
        bool IsKeyPairMatch(Dictionary<int, string> inputKeyPair);
        string GetClassDescription();
        Dictionary<string, string> GetNameAttributePair();

    }
}
