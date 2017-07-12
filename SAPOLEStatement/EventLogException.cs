using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SAPOLEStatement
{
    public class EventLogException : ApplicationException
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public EventLogException()
        {
        }

        /// <summary>
        /// Constructor accepting a message. Overloads the Exception constructor.
        /// </summary>
        /// <param name="message">The message to be included with the exception</param>
        public EventLogException(String message)
            : base(message)
        {
        }

        /// <summary>
        /// Constructor accepting a message and an exception. Overloads the Exception 
        /// constructor.
        /// </summary>
        /// <param name="message">The message to be included with the exception</param>
        /// <param name="inner">The exception to be stored as the inner exception. 
        /// This is usually the exception that was originally raised and is being replaced
        /// by this exception.</param>
        public EventLogException(String message, Exception inner)
            : base(message, inner)
        {
        }
    }

}
