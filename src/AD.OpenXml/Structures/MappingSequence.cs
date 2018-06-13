using System;
using System.Collections.Generic;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc />
    /// <summary>
    /// Represents a sequence of unsigned integers beginning with 1 and formatted to a specified template.
    /// Values can be mapped to keys for later retrieval.
    /// </summary>
    public class MappingSequence : Sequence
    {
        /// <summary>
        /// The mapping of keys and sequence values.
        /// </summary>
        [NotNull] private readonly Dictionary<string, string> _map = new Dictionary<string, string>();

        /// <inheritdoc />
        public MappingSequence([CanBeNull] string template = "{0}") : base(template)
        {
        }

        /// <summary>
        /// Returns the sequence value associated with the specified key, or the next value in the sequence.
        /// </summary>
        /// <param name="key">The key locate an existing sequence value or to associate with the next sequence value.</param>
        /// <returns>
        /// The value associated with the key, or the next value in the sequence.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [NotNull]
        [MustUseReturnValue]
        public string GetOrNextValue([NotNull] string key)
        {
            if (key is null)
                throw new ArgumentNullException(nameof(key));

            return TryGetValue(key, out string value) ? value : (_map[key] = NextValue());
        }

        /// <summary>
        /// Returns the next sequence value and associates it with the specified key.
        /// </summary>
        /// <param name="key">The key to associate with the next sequence value.</param>
        /// <returns>
        /// The next value in the sequence.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [NotNull]
        [MustUseReturnValue]
        public string NextValue([NotNull] string key)
        {
            if (key is null)
                throw new ArgumentNullException(nameof(key));

            return _map[key] = NextValue();
        }

        /// <summary>
        /// Returns the sequence value mapped to the specified key..
        /// </summary>
        /// <param name="key">The key to lookup in the sequence.</param>
        /// <returns>
        /// True if the key was found; otherwise false.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        /// <exception cref="ArgumentException">The key is not mapped in the sequence.</exception>
        [Pure]
        [NotNull]
        public string GetValue([NotNull] string key)
        {
            if (key is null)
                throw new ArgumentNullException(nameof(key));
            if (!_map.ContainsKey(key))
                throw new ArgumentException("The key is not mapped in the sequence.");

            return _map[key];
        }

        /// <summary>
        /// Attempts to locate the sequence value mapped to the specified key.
        /// </summary>
        /// <param name="key">The key to lookup in the sequence.</param>
        /// <param name="value">The value if the key is mapped; otherwise null.</param>
        /// <returns>
        /// True if the key was found; otherwise false.
        /// </returns>
        [Pure]
        [ContractAnnotation("=> true, value: notnull; => false, value: null")]
        public bool TryGetValue([NotNull] string key, [CanBeNull] out string value)
        {
            if (key is null)
                throw new ArgumentNullException(nameof(key));

            return _map.TryGetValue(key, out value);
        }
    }
}