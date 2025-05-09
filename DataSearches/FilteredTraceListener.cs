using System.Diagnostics;
using System.IO;

namespace DataSearches
{
    internal class FilteredTraceListener : TextWriterTraceListener
    {
        public FilteredTraceListener(StreamWriter writer) : base(writer) { }

        public override void WriteLine(string message)
        {
            if (message?.StartsWith("Warning: Serialization methods should not be called from the GUI thread") == true)
                return;

            base.WriteLine(message);
        }
    }
}