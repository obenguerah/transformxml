using System;
using System.Collections.Generic;
using System.Text;

namespace AndrewTweddle.Tools.TransformXml
{
    [Serializable]
    class CommandLineSwitchException : Exception
    {
        public CommandLineSwitchException()
            : base()
        {
        }

        public CommandLineSwitchException(string message)
            : base(message)
        {
        }

        public CommandLineSwitchException(string message,
            Exception innerException)
            : base(message, innerException)
        {
        }

        protected CommandLineSwitchException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
        }
    }
}
