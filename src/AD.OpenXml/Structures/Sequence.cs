using System.Runtime.CompilerServices;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <summary>
    /// Represents a thread-safe sequence of unsigned integers beginning with 1.
    /// </summary>
    [PublicAPI]
    public class Sequence
    {
        [NotNull] private readonly object _lock = new object();

        private uint _counter;

        /// <summary>
        /// Returns the next value of the sequence.
        /// </summary>
        /// <returns>
        /// The next value in the sequence.
        /// </returns>
        [MustUseReturnValue]
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public uint NextValue()
        {
            lock (_lock)
            {
                return ++_counter;
            }
        }

        /// <inheritdoc />
        public override string ToString() => $"Sequence: {_counter}.";
    }
}