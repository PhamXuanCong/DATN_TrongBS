using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ProjectApp.ModelFromCad
{
    public static class ComRunningObject
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CLSIDFromProgIDEx(string progId, out Guid clsid);

        [DllImport("oleaut32.dll")]
        public static extern int GetActiveObject(ref Guid clsid, IntPtr reserved, out IntPtr punk);

        public static object GetActiveObjectByProgId(string progId)
        {
            if (string.IsNullOrWhiteSpace(progId))
                throw new ArgumentNullException(nameof(progId));

            int hr = CLSIDFromProgIDEx(progId, out var clsid);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);

            hr = GetActiveObject(ref clsid, IntPtr.Zero, out var punk);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);

            try
            {
                return Marshal.GetObjectForIUnknown(punk);
            }
            finally
            {
                Marshal.Release(punk);
            }
        }
    }
}
