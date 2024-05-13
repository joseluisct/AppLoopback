using System.Runtime.InteropServices;
using System.Threading;
using Windows.Win32.Foundation;
using Windows.Win32.Media.Audio;
using Windows.Win32.System.Com.StructuredStorage;
using Windows.Win32.System.Variant;
using static Windows.Win32.PInvoke;

namespace AudioPrueba;

public enum DeviceState
{
    Uninitialized,
    Error,
    Initialized,
    Starting,
    Capturing,
    Stopping,
    Stopped,
};

public sealed partial class CLoopbackCapture(int processId, bool includeProcessTree = true)
{
    /*const uint MF_SDK_VERSION = 0x0002;
    const uint MF_API_VERSION = 0x0070;
    const uint MF_VERSION = MF_SDK_VERSION << 16 | MF_API_VERSION;
    const uint MFSTARTUP_NOSOCKET = 0x1;
    HRESULT S_OK = (HRESULT)0;*/
    uint m_dwQueueID = 0;
    uint m_cbHeaderSize = 0;
    uint m_cbDataSize = 0;
    DeviceState m_DeviceState = DeviceState.Uninitialized;
    IAudioClient m_AudioClient;
    readonly AutoResetEvent m_SampleReadyEvent = new(false);
    //readonly AutoResetEvent m_hActivateCompleted = new(false);
    readonly AutoResetEvent m_hCaptureStopped = new(false);
    string m_outputFileName;

    HRESULT SetDeviceStateErrorIfFailed(HRESULT hr)
    {
        if (hr < 0) m_DeviceState = DeviceState.Error;
        return hr;
    }

    HRESULT InitializeLoopbackCapture()
    {
        // Create events for sample ready or user stop
        //m_SampleReadyEvent.create(wil::EventOptions::None).ThrowOnFailure();
        // Initialize MF
        MFStartup(MF_VERSION, MFSTARTUP_LITE).ThrowOnFailure();
        // Register MMCSS work queue
        //uint dwTaskID = 0;
        //MFLockSharedWorkQueue("Capture", 0, ref dwTaskID, out uint m_dwQueueID).ThrowOnFailure();
        // Set the capture event work queue to use the MMCSS queue
        //m_xSampleReady.SetQueueID(m_dwQueueID);
        // Create the completion event as auto-reset
        //m_hActivateCompleted.create(wil::EventOptions::None).ThrowOnFailure();
        // Create the capture-stopped event as auto-reset
        //m_hCaptureStopped.create(wil::EventOptions::None).ThrowOnFailure();
        return HRESULT.S_OK;
    }

    ~CLoopbackCapture()
    {
        /*if (m_dwQueueID != 0)
        {
            MFUnlockWorkQueue(m_dwQueueID);
        }*/
    }

    unsafe IAudioClient ActivateAudioInterface()
    {
        const string VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK = @"VAD\Process_Loopback";
        ActivateAudioInterfaceCompletionHandler<IAudioClient> m_hActivateCompleted = new();
        ActivationParameters @params = new()
        {
            Type = ActivationType.ProcessLoopback,
            ProcessLoopbackParams = new()
            {
                TargetProcessId = processId,
                ProcessLoopbackMode = includeProcessTree ? LoopbackMode.IncludeProcessTree : LoopbackMode.ExcludeProcessTree
            }
        };

        PROPVARIANT propVariant = new()
        {
            Anonymous =
            {
                Anonymous =
                {
                    vt = VARENUM.VT_BLOB,
                    Anonymous =
                    {
                        blob =
                        {
                            cbSize = (uint)sizeof(ActivationParameters),
                            pBlobData = (byte*)(&@params)
                        }
                    }
                }
            }
        };

        ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, typeof(IAudioClient).GUID, propVariant, m_hActivateCompleted, out var resultHandler).ThrowOnFailure();
        m_hActivateCompleted.WaitForCompletion();
        HRESULT hresult = default;
        resultHandler.GetActivateResult(&hresult, out var result);
        SetDeviceStateErrorIfFailed(hresult);
        return (IAudioClient)result;
    }

    HRESULT CreateWAVFile()
    {
            /*m_hFile.reset(CreateFile(m_outputFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
            RETURN_LAST_ERROR_IF(!m_hFile);
            // Create and write the WAV header
            // 1. RIFF chunk descriptor
            DWORD header[] = {
                                FCC('RIFF'),        // RIFF header
                                0,                  // Total size of WAV (will be filled in later)
                                FCC('WAVE'),        // WAVE FourCC
                                FCC('fmt '),        // Start of 'fmt ' chunk
                                sizeof(m_CaptureFormat) // Size of fmt chunk
            };
            DWORD dwBytesWritten = 0;
            WriteFile(m_hFile.get(), header, sizeof(header), &dwBytesWritten, NULL);
            m_cbHeaderSize += dwBytesWritten;

            // 2. The fmt sub-chunk
            WI_ASSERT(m_CaptureFormat.cbSize == 0);
            RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_CaptureFormat, sizeof(m_CaptureFormat), &dwBytesWritten, NULL));
            m_cbHeaderSize += dwBytesWritten;
            // 3. The data sub-chunk
            DWORD data[] = { FCC('data'), 0 };  // Start of 'data' chunk
            RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), data, sizeof(data), &dwBytesWritten, NULL));
            m_cbHeaderSize += dwBytesWritten;*/
            return HRESULT.S_OK;
    }

    public int StartCaptureAsync(string outputFileName)
    {
        m_outputFileName = outputFileName;
        //auto resetOutputFileName = wil::scope_exit([&] { m_outputFileName = nullptr; });

        InitializeLoopbackCapture().ThrowOnFailure();
        ActivateAudioInterface();

        // We should be in the initialzied state if this is the first time through getting ready to capture.
        if (m_DeviceState == DeviceState.Initialized)
        {
            m_DeviceState = DeviceState.Starting;
            //m_SampleReadyEvent.Set();
            OnStartCapture();
            //return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStartCapture, nullptr);
        }
        return (int)HRESULT.S_OK;
    }

    HRESULT OnStartCapture()
    {
        m_AudioClient.Start();
        m_DeviceState = DeviceState.Capturing;
        //m_SampleReadyEvent.Set();
        return HRESULT.S_OK;
    }

    HRESULT StopCaptureAsync()
    {
        if(m_DeviceState != DeviceState.Capturing && m_DeviceState != DeviceState.Error) return (HRESULT)(-1);
        m_DeviceState = DeviceState.Stopping;
        //RETURN_IF_FAILED(MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStopCapture, nullptr));
        // Wait for capture to stop
        m_hCaptureStopped.WaitOne();
        return HRESULT.S_OK;
    }

    HRESULT OnStopCapture()
    {
        // Stop capture by cancelling Work Item
        // Cancel the queued work item (if any)
       /* if (0 != m_SampleReadyKey)
        {
            MFCancelWorkItem(m_SampleReadyKey);
            m_SampleReadyKey = 0;
        }*/

        m_AudioClient.Stop();
        m_SampleReadyEvent.Reset();
        //return FinishCaptureAsync();
        return HRESULT.S_OK;
    }

    HRESULT FinishCaptureAsync()
    {
        // We should be flushing when this is called
        //return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xFinishCapture, null);
        return HRESULT.S_OK;
    }

    HRESULT OnFinishCapture()
    {
        // FixWAVHeader will set the DeviceStateStopped when all async tasks are complete
        //HRESULT hr = FixWAVHeader();
        HRESULT hr = HRESULT.S_OK;
        m_DeviceState = DeviceState.Stopped;
        m_hCaptureStopped.Set();
        return hr;
    }

    HRESULT OnSampleReady()
    {
        /*if (SUCCEEDED(OnAudioSampleRequested()))
        {
            // Re-queue work item for next sample
            if (m_DeviceState == DeviceState.Capturing)
            {
                // Re-queue work item for next sample
                return MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);
            }
        }
        else
        {
            m_DeviceState = DeviceState.Error;
        }*/
        return HRESULT.S_OK;
    }
}

public sealed class ActivateAudioInterfaceCompletionHandler<T> : IActivateAudioInterfaceCompletionHandler where T : class
{
    readonly AutoResetEvent _completionEvent = new(false);
    void IActivateAudioInterfaceCompletionHandler.ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation) => _completionEvent.Set();
    public void WaitForCompletion() => _completionEvent.WaitOne();
}

/// <summary>
/// Specifies the activation parameters for a call to <see cref="Helper.ActivateAudioInterfaceAsync(String, Guid, ByRef PropVariant, ByRef IActivateAudioInterfaceCompletionHandler)"/>.
/// </summary>
[StructLayout(LayoutKind.Sequential)]
public struct ActivationParameters
{
    /// <summary>
    /// A member of the <see cref="ActivationType">AUDIOCLIENT_ACTIVATION_TYPE</see> specifying the type of audio interface activation. <br/>
    /// Currently default activation and loopback activation are supported.
    /// </summary>
    public ActivationType Type;
    /// <summary>
    /// A <see cref="ProcessLoopbackParams">AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS</see> specifying the loopback parameters for the audio interface activation.
    /// </summary>
    public ProcessLoopbackParams ProcessLoopbackParams;
}

/// <summary>
/// Specifies the activation type for an <see cref="ActivationParameters">AUDIOCLIENT_ACTIVATION_PARAMS</see> structure passed into a call to <see cref="Helper.ActivateAudioInterfaceAsync(String, Guid, ByRef PropVariant, ByRef IActivateAudioInterfaceCompletionHandler)" />.
/// </summary>
public enum ActivationType
{
    /// <summary>
    /// Default activation.
    /// </summary>
    Default,
    /// <summary>
    /// Process loopback activation, allowing for the inclusion or exclusion of audio rendered by the specified process and its child processes. <br/>
    /// For sample code that demonstrates the process loopback capture scenario, see the Application Loopback <see href="https://docs.microsoft.com/en-us/samples/microsoft/windows-classic-samples/applicationloopbackaudio-sample/">API Capture Sample.</see>
    /// </summary>
    ProcessLoopback
}

[StructLayout(LayoutKind.Sequential)]
public struct ProcessLoopbackParams
{
    /// <summary>
    /// The ID of the process for which the render streams, and the render streams of its child processes, will be included or excluded when activating the process loopback stream.
    /// </summary>
    public int TargetProcessId;
    /// <summary>
    /// A value from the <see cref="LoopbackMode">PROCESS_LOOPBACK_MODE</see> enumeration specifying whether the render streams for the process and child processes specified in the TargetProcessId field should be included or excluded when activating the audio interface. <br />
    /// For sample code that demonstrates the process loopback capture scenario, see the <see href="https://docs.microsoft.com/en-us/samples/microsoft/windows-classic-samples/applicationloopbackaudio-sample/">Application Loopback API Capture Sample</see>.
    /// </summary>
    public LoopbackMode ProcessLoopbackMode;
}

/// <summary>
/// Specifies the loopback mode for an <see cref="ActivationParameters">AUDIOCLIENT_ACTIVATION_PARAMS</see> structure passed into a call to <see cref="Helper.ActivateAudioInterfaceAsync(String, Guid, ByRef PropVariant, ByRef IActivateAudioInterfaceCompletionHandler)"/>.
/// </summary>
public enum LoopbackMode
{
    /// <summary>
    /// Render streams from the specified process and its child processes are included in the activated process loopback stream.
    /// </summary>
    IncludeProcessTree,
    /// <summary>
    /// Render streams from the specified process and its child processes are excluded from the activated process loopback stream.
    /// </summary>
    ExcludeProcessTree
}