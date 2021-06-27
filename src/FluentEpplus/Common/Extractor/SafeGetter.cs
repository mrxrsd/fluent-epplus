using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace FluentEpplus.Common.Extractor
{
    public class SafeGetter<T>
    {
        private readonly bool _validChain;
        private T _value;

        public T Value
        {
            get
            {
                if (ValidChain())
                {
                    return _value;
                }

                return default(T);
            }
            private set
            {
                _value = value;
            }
        }

        public T GetValueOrDefault(T @default = default(T))
        {
            if (!ValidChain())
            {
                return @default;
            }

            return Value;
        }

        public SafeGetter(T value, bool validChain)
        {
            _validChain = validChain;
            Value = value;
        }

        public bool ValidChain()
        {
            return _validChain;
        }
    }
}
