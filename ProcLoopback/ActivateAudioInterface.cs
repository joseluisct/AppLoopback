using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wasapi.CoreAudioApi.Interfaces;
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace NAudio.Wasapi.CoreAudioApi;

internal class ActivateAudioInterfaceCompletionHandler<T> : IActivateAudioInterfaceCompletionHandler, IAgileObject
{
    private Action<T> initializeAction;
    private TaskCompletionSource<T> tcs = new TaskCompletionSource<T>();

    public ActivateAudioInterfaceCompletionHandler(Action<T> initializeAction)
    {
        this.initializeAction = initializeAction;
    }

    public void ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation)
    {
        // First get the activation results, and see if anything bad happened then
        activateOperation.GetActivateResult(out int hr, out object unk);
        if (hr != 0)
        {
            tcs.TrySetException(Marshal.GetExceptionForHR(hr, new IntPtr(-1)));
            return;
        }

        var pAudioClient = (T)unk;
        // Next try to call the client's (synchronous, blocking) initialization method.
        try
        {
            initializeAction(pAudioClient);
            tcs.SetResult(pAudioClient);
        }
        catch (Exception ex)
        {
            tcs.TrySetException(ex);
        }
    }

    public TaskAwaiter<T> GetAwaiter()
    {
        return tcs.Task.GetAwaiter();
    }
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("94ea2b94-e9cc-49e0-c0ff-ee64ca8f5b90")]
internal interface IAgileObject
{
}

internal class ActivateAudioInterfaceCompletionHandler1 : IActivateAudioInterfaceCompletionHandler, IAgileObject
{
    private Action<IAudioClient> initializeAction;
    private TaskCompletionSource<IAudioClient> tcs = new TaskCompletionSource<IAudioClient>();

    public ActivateAudioInterfaceCompletionHandler1(Action<IAudioClient> initializeAction)
    {
        this.initializeAction = initializeAction;
    }

    public void ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation)
    {
        // First get the activation results, and see if anything bad happened then
        activateOperation.GetActivateResult(out int hr, out object unk);
        if (hr != 0)
        {
            tcs.TrySetException(Marshal.GetExceptionForHR(hr, new IntPtr(-1)));
            return;
        }

        var pAudioClient = (IAudioClient)unk;
        // Next try to call the client's (synchronous, blocking) initialization method.
        try
        {
            initializeAction(pAudioClient);
            tcs.SetResult(pAudioClient);
        }
        catch (Exception ex)
        {
            tcs.TrySetException(ex);
        }
    }

    public TaskAwaiter<IAudioClient> GetAwaiter()
    {
        return tcs.Task.GetAwaiter();
    }
}

internal struct AudioClientActivationParams
{
    public AudioClientActivationType ActivationType;
    public AudioClientProcessLoopbackParams ProcessLoopbackParams;
}

internal enum AudioClientActivationType
{
    //
    // Resumen:
    //     AUDIOCLIENT_ACTIVATION_TYPE_DEFAULT Default activation.
    Default,
    //
    // Resumen:
    //     AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK Process loopback activation, allowing
    //     for the inclusion or exclusion of audio rendered by the specified process and
    //     its child processes.
    ProcessLoopback
}

internal struct AudioClientProcessLoopbackParams
{
    //
    // Resumen:
    //     AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS The ID of the process for which the render
    //     streams, and the render streams of its child processes, will be included or excluded
    //     when activating the process loopback stream.
    public uint TargetProcessId;
    public ProcessLoopbackMode ProcessLoopbackMode;
}

internal enum ProcessLoopbackMode
{
    //
    // Resumen:
    //     PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE Render streams from the specified
    //     process and its child processes are included in the activated process loopback
    //     stream.
    IncludeTargetProcessTree,
    //
    // Resumen:
    //     PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE Render streams from the specified
    //     process and its child processes are excluded from the activated process loopback
    //     stream.
    ExcludeTargetProcessTree
}