using Microsoft.VisualBasic.CompilerServices;

namespace FluentEpplus.Common
{
    public interface IExtractor
    {
        object GetValue(object targetObject);
    }

    public interface IExtractor<T> : IExtractor
    {
        object GetValue(T targetObject);
    }
}