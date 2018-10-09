using System;
using System.Collections;
using System.Collections.Generic;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEnumerable{T}" />
    /// <inheritdoc cref="IEnumerator{T}" />
    /// <summary>
    /// Represents a sequence of unsigned integers beginning with 1 and formatted to a specified template.
    /// </summary>
    [PublicAPI]
    public class Sequence : IEnumerable<string>, IEnumerator<string>
    {
        [NotNull] readonly object _lock = new object();
        [NotNull] readonly string _template;
        uint _counter;

        /// <inheritdoc />
        [NotNull]
        public string Current => string.Format(_template, _counter);

        /// <inheritdoc />
        [NotNull]
        object IEnumerator.Current => Current;

        /// <summary>
        /// Initializes a new sequence with the specified format template.
        /// </summary>
        /// <param name="template">The string format template.</param>
        public Sequence([CanBeNull] string template = "{0}") => _template = template ?? "{0}";

        /// <summary>
        /// Returns the next value of the sequence.
        /// </summary>
        /// <returns>
        /// The next value in the sequence.
        /// </returns>
        [NotNull]
        [MustUseReturnValue]
        public string NextValue()
        {
            bool _ = MoveNext();
            return Current;
        }

        /// <inheritdoc />
        [MustUseReturnValue]
        public bool MoveNext()
        {
            lock (_lock)
            {
                return ++_counter != 0;
            }
        }

        /// <inheritdoc />
        [NotNull]
        IEnumerator<string> IEnumerable<string>.GetEnumerator() => this;

        /// <inheritdoc />
        [NotNull]
        IEnumerator IEnumerable.GetEnumerator() => this;

        /// <inheritdoc />
        void IEnumerator.Reset() => _counter = default;

        /// <inheritdoc />
        void IDisposable.Dispose() => _counter = default;

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString() => $"(Current: {Current})";
    }
}