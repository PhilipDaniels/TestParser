using System;
using System.Globalization;
using System.Runtime.Serialization;

namespace TestParser.Core
{
    /// <summary>
    /// This class demonstrates the canonical form of an application-specific exception.
    /// If you add further members you should ensure that they are serialized properly.
    /// See http://stackoverflow.com/questions/94488/what-is-the-correct-way-to-make-a-custom-net-exception-serializable?lq=1
    /// </summary>
    [Serializable]
    public class ParseException : Exception
    {
        /// <summary>
        /// Construct a new ParseException.
        /// </summary>
        public ParseException()
        {
        }

        /// <summary>
        /// Construct a new ParseException.
        /// </summary>
        /// <param name="message">Message to use.</param>
        public ParseException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Construct a new ParseException with formatted message.
        /// </summary>
        /// <param name="format">Format string.</param>
        /// <param name="args">Optional arguments for message.</param>
        public ParseException(string format, params object[] args)
            : base(String.Format(CultureInfo.InvariantCulture, format, args))
        {
        }

        /// <summary>
        /// Construct a new ParseException.
        /// </summary>
        /// <param name="message">Message to use.</param>
        /// <param name="innerException">Inner exception.</param>
        public ParseException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Construct a new ParseException with formatted message.
        /// </summary>
        /// <param name="format">Format string.</param>
        /// <param name="innerException">Inner exception.</param>
        /// <param name="args">Optional arguments for message.</param>
        public ParseException(string format, Exception innerException, params object[] args)
            : base(String.Format(CultureInfo.InvariantCulture, format, args), innerException)
        {
        }

        /// <summary>
        /// Construct a new ParseException using a serialization context.
        /// </summary>
        /// <param name="info">Serialization info.</param>
        /// <param name="context">Streaing context.</param>
        protected ParseException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }

}
