using System;
using System.Runtime.InteropServices;
using System.Threading;
using Windows.Win32.Foundation;
using Windows.Win32.Media.Audio;
using Windows.Win32.System.Com;
//using static Windows.Win32.PInvoke;

namespace ProcLoopback;

public class ActivateAudioInterfaceAsyncCallback : IActivateAudioInterfaceCompletionHandler
{
    private ManualResetEvent m_bActivateCompleted = new ManualResetEvent(false);
    public ActivateAudioInterfaceAsyncCallback()
    {
    }

    public void Wait()
    {
        m_bActivateCompleted.WaitOne();
        m_bActivateCompleted.Reset();
    }

    void ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation)
    {
        m_bActivateCompleted.Set();
    }

    public void QueryInterface(ref Guid riid, out IntPtr ppvObject)
    {
        if (riid == typeof(IAgileObject).GUID)
        {
            ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IActivateAudioInterfaceCompletionHandler));
        }
        else
        {
            ppvObject = IntPtr.Zero;
            throw new COMException("No such interface supported", (int)HRESULT.E_NOINTERFACE);
        }
    }

    public uint AddRef()
    {
        return 1;
    }

    public uint Release()
    {
        return 0;
    }

    void IActivateAudioInterfaceCompletionHandler.ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation)
    {
        throw new NotImplementedException();
    }
}

public class ProcessLoopbackCapture
{
}