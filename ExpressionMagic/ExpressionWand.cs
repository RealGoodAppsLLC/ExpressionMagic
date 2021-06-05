using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;

namespace RealGoodApps.ExpressionMagic
{
    /// <summary>
    /// Magic functions for creating, augmenting, and combining expressions.
    /// </summary>
    public static class ExpressionWand
    {
        /// <summary>
        /// Compile and invoke an expression on a parameter.
        /// </summary>
        /// <param name="expression">The expression to compile and invoke.</param>
        /// <param name="parameter">The parameter to pass as the parameter to the compiled expression.</param>
        /// <param name="alwaysThrow">If this parameter is true, then the function will throw even if the exception is an null reference or argument null exception.</param>
        /// <typeparam name="TParameter">The parameter type.</typeparam>
        /// <typeparam name="TResult">The result type.</typeparam>
        /// <returns>The result of the function invocation.</returns>
        [SuppressMessage(
            "ReSharper",
            "ParameterOnlyUsedForPreconditionCheck.Global",
            Justification = "This parameter is used to control whether or not an exception should be thrown.")]
        public static TResult Invoke<TParameter, TResult>(
            this Expression<Func<TParameter, TResult>> expression,
            TParameter parameter,
            bool alwaysThrow = false)
        {
            try
            {
                var func = expression.Compile();
                return func.Invoke(parameter);
            }
            catch (Exception ex) when (ex is NullReferenceException || ex is ArgumentNullException)
            {
                // This allows us to return default/null when the expression generates a typically-displayed exception.
                // You can toggle this behavior with the alwaysThrow parameter.
                if (alwaysThrow)
                {
                    throw;
                }

                return default;
            }
        }

        /// <summary>
        /// Generate a predicate from an existing predicate and the result of "Or" with another predicate.
        /// </summary>
        /// <param name="left">The existing predicate.</param>
        /// <param name="right">The predicate to "Or" the existing predicate against.</param>
        /// <typeparam name="TParameter">The type of the parameter that the predicate is performed on.</typeparam>
        /// <returns>A lambda expression which is the result of the combination of expressions.</returns>
        public static Expression<Func<TParameter, bool>> Or<TParameter>(
            this Expression<Func<TParameter, bool>> left,
            Expression<Func<TParameter, bool>> right)
        {
            var orElseExpression = Expression.OrElse(
                left.Body,
                ReplaceParameter(right.Body, right.Parameters[0], left.Parameters[0]));

            return Expression.Lambda<Func<TParameter, bool>>(
                orElseExpression,
                left.Parameters);
        }

        /// <summary>
        /// Generate a predicate from an existing predicate and the result of "And" with another predicate.
        /// </summary>
        /// <param name="expr1">The existing predicate.</param>
        /// <param name="expr2">The predicate to "And" the existing predicate against.</param>
        /// <typeparam name="TParameter">The type of the parameter that the predicate is performed on.</typeparam>
        /// <returns>A lambda expression which is the result of the combination of expressions.</returns>
        public static Expression<Func<TParameter, bool>> And<TParameter>(
            this Expression<Func<TParameter, bool>> expr1,
            Expression<Func<TParameter, bool>> expr2)
        {
            var andAlsoExpression = Expression.AndAlso(
                expr1.Body,
                ReplaceParameter(expr2.Body, expr2.Parameters[0], expr1.Parameters[0]));

            return Expression.Lambda<Func<TParameter, bool>>(
                andAlsoExpression,
                expr1.Parameters);
        }

        /// <summary>
        /// Generate a predicate that inverts the value of an existing predicate.
        /// </summary>
        /// <param name="expr">The existing predicate.</param>
        /// <typeparam name="TParameter">The type of the parameter that the predicate is performed on.</typeparam>
        /// <returns>A lambda expression which is the result of the inverted expression.</returns>
        public static Expression<Func<TParameter, bool>> Not<TParameter>(
            this Expression<Func<TParameter, bool>> expr)
        {
            var notExpression = Expression.Not(expr.Body);
            return Expression.Lambda<Func<TParameter, bool>>(notExpression, expr.Parameters[0]);
        }

        /// <summary>
        /// Generate a predicate that takes an expression and compares the result against a value.
        /// </summary>
        /// <param name="expr">The existing expression.</param>
        /// <param name="comparison">The value to compare against the expression's result.</param>
        /// <typeparam name="TParameter">The type of the parameter that the expression is performed on.</typeparam>
        /// <typeparam name="TValue">The type of the value to compare.</typeparam>
        /// <returns>A lambda expression which is the result of the comparison.</returns>
        public static Expression<Func<TParameter, bool>> IsEqual<TParameter, TValue>(
            this Expression<Func<TParameter, TValue>> expr,
            TValue comparison)
        {
            var equalExpression = Expression.Equal(expr.Body, Expression.Constant(comparison, typeof(TValue)));
            return Expression.Lambda<Func<TParameter, bool>>(equalExpression, expr.Parameters[0]);
        }

        /// <summary>
        /// Generate a predicate that takes an expression and compares the result against a value.
        /// </summary>
        /// <param name="expr">The existing expression.</param>
        /// <param name="comparison">The value to compare against the expression's result.</param>
        /// <typeparam name="TParameter">The type of the parameter that the expression is performed on.</typeparam>
        /// <typeparam name="TValue">The type of the value to compare.</typeparam>
        /// <returns>A lambda expression which is the result of the comparison.</returns>
        public static Expression<Func<TParameter, bool>> IsEqual<TParameter, TValue>(
            this Expression<Func<TParameter, TValue>> expr,
            Expression<Func<TParameter, TValue>> comparison)
        {
            var body = Expression.Equal(
                expr.Body,
                ReplaceParameter(comparison.Body, comparison.Parameters[0], expr.Parameters[0]));

            return Expression.Lambda<Func<TParameter, bool>>(body, expr.Parameters[0]);
        }

        /// <summary>
        /// Create a ternary (if, then, else) expression using a "test", "if-true", and "if-false" expression.
        /// </summary>
        /// <param name="test">The test expression.</param>
        /// <param name="ifTrue">The expression to use if the test expression returns true.</param>
        /// <param name="ifFalse">The expression to use if the test expression returns false.</param>
        /// <typeparam name="TParameter">The type of the parameter.</typeparam>
        /// <typeparam name="TResult">The type of the result.</typeparam>
        /// <returns>The resulting ternary expression.</returns>
        public static Expression<Func<TParameter, TResult>> CreateIfThenElse<TParameter, TResult>(
            Expression<Func<TParameter, bool>> test,
            Expression<Func<TParameter, TResult>> ifTrue,
            Expression<Func<TParameter, TResult>> ifFalse)
        {
            var ifThenElse = Expression.Condition(
                test.Body,
                ReplaceParameter(ifTrue.Body, ifTrue.Parameters[0], test.Parameters[0]),
                ReplaceParameter(ifFalse.Body, ifFalse.Parameters[0], test.Parameters[0]));

            return Expression.Lambda<Func<TParameter, TResult>>(ifThenElse, test.Parameters[0]);
        }

        /// <summary>
        /// Pipe the result of an expression into another expression, resulting in an expression that takes in a
        /// normalized parameter.
        /// </summary>
        /// <param name="source">The expression which will have it's result piped into the second expression.</param>
        /// <param name="target">The receiver expression of the piped expression result.</param>
        /// <typeparam name="TSource">The parameter type, which will be used to normalize the final expression.</typeparam>
        /// <typeparam name="TPipedResult">The result type which will be piped into the other expression.</typeparam>
        /// <typeparam name="TResult">The result type of the receiver expression and final result.</typeparam>
        /// <returns>The resulting expression.</returns>
        /// <exception cref="InvalidOperationException">If re-writing the second expression fails.</exception>
        public static Expression<Func<TSource, TResult>> Pipe<TSource, TPipedResult, TResult>(
            this Expression<Func<TSource, TPipedResult>> source,
            Expression<Func<TPipedResult, TResult>> target)
        {
            var secondExpressionBodyRewritten = ReplaceParameter(
                target.Body,
                target.Parameters[0],
                source.Body);

            return Expression.Lambda<Func<TSource, TResult>>(secondExpressionBodyRewritten, source.Parameters[0]);
        }

        /// <summary>
        /// Create an expression and return it.
        /// </summary>
        /// <param name="expr">The expression lambda.</param>
        /// <typeparam name="TParameter">The type of parameter for the lambda.</typeparam>
        /// <typeparam name="TResult">The type of result for the lambda.</typeparam>
        /// <returns>An instance of the expression.</returns>
        public static Expression<Func<TParameter, TResult>> Create<TParameter, TResult>(Expression<Func<TParameter, TResult>> expr) => expr;

        private static Expression ReplaceParameter(
            Expression expression,
            ParameterExpression searched,
            Expression replace)
        {
            var bodyWithReplacedParameter = new ParameterReplaceVisitor(searched, replace).Visit(expression);

            if (bodyWithReplacedParameter == null)
            {
                throw new InvalidOperationException("Unable to re-write expression with specified parameter.");
            }

            return bodyWithReplacedParameter;
        }

        private sealed class ParameterReplaceVisitor : ExpressionVisitor
        {
            private readonly ParameterExpression _searched;
            private readonly Expression _replaced;

            public ParameterReplaceVisitor(ParameterExpression searched, Expression replaced)
            {
                _searched = searched ?? throw new ArgumentNullException(nameof(searched));
                _replaced = replaced ?? throw new ArgumentNullException(nameof(replaced));
            }

            protected override Expression VisitParameter(ParameterExpression node)
            {
                return node == _searched ? _replaced : base.VisitParameter(node);
            }
        }
    }
}
