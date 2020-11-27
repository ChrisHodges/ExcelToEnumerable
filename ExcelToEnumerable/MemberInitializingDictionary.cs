using System.Collections.Generic;

namespace ExcelToEnumerable
{
    internal class MemberInitializingDictionary<TKey, TValue> : Dictionary<TKey, TValue> where TValue : new()
    {
        public new TValue this[TKey index]
        {
            get
            {
                if (!ContainsKey(index))
                {
                    base[index] = new TValue();
                }

                return base[index];
            }
            set => base[index] = value;
        }
    }
}