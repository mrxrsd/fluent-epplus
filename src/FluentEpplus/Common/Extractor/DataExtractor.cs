using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Text;

namespace FluentEpplus.Common.Extractor
{
    public interface IDataExtractor
    {
        object GetValue(object fromData);
    }

    public class DataExtractor<TDto,TValue> : IDataExtractor
    {
        private readonly Func<TDto, TValue> _compiledFunction;
        private readonly Func<TDto, SafeGetter<TValue>> _safeCompiledFunction;

        public DataExtractor(Expression<Func<TDto, TValue>> propertyExpression)
        {
            if (propertyExpression.Body.ToString().Split('.').Length > 2)
            {
                _safeCompiledFunction = CreateSafePath(propertyExpression);
            }
            else
            {
                _compiledFunction = propertyExpression.Compile();
            }
        }

        protected Func<Y, SafeGetter<T>> CreateSafePath<Y, T>(Expression<Func<Y, T>> propertyExpression)
        {
            var transform = (Expression<Func<Y, SafeGetter<T>>>) new NullVisitor<T>().Visit(propertyExpression);
            return transform.Compile();
        }

        public object GetValue(object data)
        {
            return (_compiledFunction != null) ? GetUnsafeValue(data) : GetSafeValue(data);
        }

        protected virtual object GetUnsafeValue(object data)
        {
            return _compiledFunction.Invoke((TDto) data);
        }

        protected virtual object GetSafeValue(object data)
        {
            return _safeCompiledFunction((TDto) data).Value;
        }
    }
}
