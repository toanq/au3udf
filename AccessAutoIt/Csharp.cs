using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AutoItFun
{
    // I saw other people doing something similar
    // Check http://www.autoitscript.com/forum/topic/110379-autoitobject-udf/page__view__findpost__p__884486
    class Program
    {
        static void Main(string[] args)
        {
            object myAu3Object = AutoItHelper.GetObject("AutoIt.Application");
            if (myAu3Object == null) return;
            Type myAu3Type = myAu3Object.GetType();

            myAu3Type.InvokeMember("Call", BindingFlags.InvokeMethod, null, myAu3Object, new object[] { "ToolTip", "Some cool text" });
            myAu3Type.InvokeMember("Call", BindingFlags.InvokeMethod, null, myAu3Object, new object[] { "Beep", 500, 700 });
            myAu3Type.InvokeMember("Call", BindingFlags.InvokeMethod, null, myAu3Object, new object[] { "Sleep", 1000 });
            myAu3Type.InvokeMember("Call", BindingFlags.InvokeMethod, null, myAu3Object, new object[] { "ToolTip", "" });

            Int32 iRet = (Int32)myAu3Type.InvokeMember("Call", BindingFlags.InvokeMethod, null, myAu3Object, new object[] { "MsgBox", 4 + 48 + 262144, "?", "Kill Server?" });
            if (iRet == 6)
            {
                myAu3Type.InvokeMember("Quit", BindingFlags.InvokeMethod, null, myAu3Object, null);
            }
        }
    }
}


//...
namespace AutoItFun
{
    //////////////////////////////////////////////////
    ///
    /// <author>Nicholas Paldino</author>
    ///
    //////////////////////////////////////////////////
    // bent over by trancexx for this demo

    public class AutoItHelper
    {
        public static Guid IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        [StructLayout(LayoutKind.Sequential)]
        public struct BIND_OPTS
        {
            [MarshalAs(UnmanagedType.U4)]
            public int cbStruct;
            [MarshalAs(UnmanagedType.U4)]
            public int grfFlags;
            [MarshalAs(UnmanagedType.U4)]
            public int grfMode;
            [MarshalAs(UnmanagedType.U4)]
            public int dwTickCountDeadline;
        }

        [DllImport("ole32.dll")]
        public extern static int CoGetObject([MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.Struct)] ref BIND_OPTS bindOptions, ref Guid iid, out IntPtr ppv);

        public static object GetObject(string moniker)
        {
            // Set the structure up for BIND_OPTS.
            BIND_OPTS pobjBindOpts = new BIND_OPTS();
            // Set the size of the structure.
            pobjBindOpts.cbStruct = Marshal.SizeOf(typeof(BIND_OPTS));

            // Leave the rest. Define the interface pointer.
            IntPtr pintObject = IntPtr.Zero;

            // Make the call.
            int pintResult = CoGetObject(moniker, ref pobjBindOpts, ref IUnknown, out pintObject);

            // Check for errors or whatever here... (layman's touch by trancexx)
            if (pintResult != 0) return null;

            // Marshal to an object.
            object pobjRetVal = Marshal.GetObjectForIUnknown(pintObject);

            // Release the reference on the object (there are two because of the calls to CoGetObject and GetObjectForIUnknown.
            Marshal.Release(pintObject);

            // That's all folks.
            return pobjRetVal;
        }
    }
}
