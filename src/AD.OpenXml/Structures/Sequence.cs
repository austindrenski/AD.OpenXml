using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc />
    /// <summary>
    /// Represents a thread-safe sequence of unsigned integers beginning with 1.
    /// </summary>
    [PublicAPI]
    public class Sequence : IEnumerable<string>
    {
        /// <summary>
        /// The current enumerator.
        /// </summary>
        [NotNull] private readonly Enumerator _enumerator;

        /// <summary>
        /// Initializes a new sequence with the specified format template.
        /// </summary>
        /// <param name="template">The string format template.</param>
        public Sequence([NotNull] string template = "{0}") => _enumerator = new Enumerator(template);

        /// <summary>
        /// Returns the next value of the sequence.
        /// </summary>
        /// <returns>
        /// The next value in the sequence.
        /// </returns>
        [NotNull]
        [MustUseReturnValue]
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public string NextValue()
        {
            bool _ = _enumerator.MoveNext();
            return _enumerator.Current;
        }

        /// <inheritdoc />
        IEnumerator<string> IEnumerable<string>.GetEnumerator() => _enumerator;

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => _enumerator;

        /// <inheritdoc />
        public override string ToString() => $"(Current: {_enumerator.Current})";

        /// <inheritdoc />
        private class Enumerator : IEnumerator<string>
        {
            [NotNull] private readonly object _lock = new object();
            [NotNull] private readonly string _template;
            private uint _counter;

            /// <inheritdoc />
            public string Current => string.Format(_template, _counter);

            /// <inheritdoc />
            object IEnumerator.Current => Current;

            /// <summary>
            /// Initializes a new sequence with the specified format template.
            /// </summary>
            /// <param name="template">The string format template.</param>
            public Enumerator([NotNull] string template) => _template = template;

            /// <inheritdoc />
            public bool MoveNext()
            {
                lock (_lock)
                {
                    return ++_counter != 0;
                }
            }

            /// <inheritdoc />
            void IEnumerator.Reset() => _counter = default;

            /// <inheritdoc />
            void IDisposable.Dispose() => _counter = default;
        }
    }
}