using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace FluentEpplus.Common.Extractor
{
    public class DecoratorFormatExtractor<TDto,TValue> : IDataExtractor
    {
        private readonly Func<TValue, object> _compiledFunction;
        private readonly Func<TValue, SafeGetter<object>> _safeCompiledFunction;
        private readonly IDataExtractor _baseExtractor;

        public DecoratorFormatExtractor(IDataExtractor baseExtractor, Expression<Func<TValue, object>> formatExpression)
        {
            _baseExtractor = baseExtractor;
            if (formatExpression.Body.ToString().Split('.').Length > 2)
            {
                _safeCompiledFunction = CreateSafePath(formatExpression);
            }
            else
            {
                _compiledFunction = formatExpression.Compile();
            }
        }

        protected Func<Y, SafeGetter<T>> CreateSafePath<Y, T>(Expression<Func<Y, T>> propertyExpression)
        {
            var transform = (Expression<Func<Y, SafeGetter<T>>>)new NullVisitor<T>().Visit(propertyExpression);
            return transform.Compile();
        }

        public object GetValue(object data)
        {
            return (_compiledFunction != null) ? GetUnsafeValue(data) : GetSafeValue(data);
        }

        protected virtual object GetUnsafeValue(object data)
        {
            return _compiledFunction.Invoke((TValue)data);
        }

        protected virtual object GetSafeValue(object data)
        {
            return _safeCompiledFunction((TValue)data).Value;
        }
    }
}
