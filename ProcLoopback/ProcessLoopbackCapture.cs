using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;

public enum eCaptureState
{
    READY = 0,
    CAPTURING,
    PAUSED
}

public enum eCaptureError
{
    NONE = 0,
    PARAM,
    STATE,
    NOT_AVAILABLE,
    FORMAT,
    PROCESSID,
    DEVICE,
    ACTIVATION,
    INITIALIZE,
    SERVICE,
    START,
    STOP,
    EVENT,
    INTERFACE
}

[StructLayout(LayoutKind.Sequential)]
public struct WAVEFORMATEX
{
    public ushort wFormatTag;
    public ushort nChannels;
    public uint nSamplesPerSec;
    public uint nAvgBytesPerSec;
    public ushort nBlockAlign;
    public ushort wBitsPerSample;
    public ushort cbSize;
}

[StructLayout(LayoutKind.Sequential)]
public struct AUDIOCLIENT_ACTIVATION_PARAMS
{
    public int ActivationType;
    public ProcessLoopbackParams ProcessLoopbackParams;
}

[StructLayout(LayoutKind.Sequential)]
public struct ProcessLoopbackParams
{
    public int ProcessLoopbackMode;
    public uint TargetProcessId;
}

[StructLayout(LayoutKind.Sequential)]
public struct PROPVARIANT
{
    public ushort vt;
    public ushort wReserved1;
    public ushort wReserved2;
    public ushort wReserved3;
    public IntPtr blob;
    public uint cbSize;
}

[ComImport]
[Guid("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAudioClient
{
    int Initialize(uint ShareMode, uint StreamFlags, long hnsBufferDuration, long hnsPeriodicity, ref WAVEFORMATEX pFormat, IntPtr AudioSessionGuid);
    int GetService(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
    int Start();
    int Stop();
    int Reset();
    int SetEventHandle(IntPtr eventHandle);
}

[ComImport]
[Guid("C8ADBD64-E71E-48a0-A4DE-185C395CD317")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAudioCaptureClient
{
    int GetBuffer(out IntPtr ppData, out uint pNumFramesToRead, out uint pdwFlags, out ulong pu64DevicePosition, out ulong pu64QPCPosition);
    int ReleaseBuffer(uint NumFramesRead);
}

[ComImport]
[Guid("41D949AB-9862-444A-80F6-C261334DA5EB")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IActivateAudioInterfaceAsyncOperation
{
    int GetActivateResult(out int activateResult, [MarshalAs(UnmanagedType.IUnknown)] out object activatedInterface);
    void Release();
}

[ComImport]
[Guid("41D949AB-9862-444A-80F6-C261334DA5EB")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IActivateAudioInterfaceCompletionHandler
{
    int ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation);
}

public class ActivateAudioInterfaceAsyncCallback : IActivateAudioInterfaceCompletionHandler
{
    private AutoResetEvent activateCompletedEvent = new AutoResetEvent(false);

    public void Wait()
    {
        activateCompletedEvent.WaitOne();
    }

    public int ActivateCompleted(IActivateAudioInterfaceAsyncOperation activateOperation)
    {
        activateCompletedEvent.Set();
        return 0; // S_OK
    }
}

public class ProcessLoopbackCapture
{
    private IAudioClient audioClient;
    private IAudioCaptureClient audioCaptureClient;
    private IntPtr sampleReadyEvent;
    private WAVEFORMATEX captureFormat;
    private eCaptureState captureState;
    private Thread mainAudioThread;
    private bool runAudioThreads;
    private int lastError;
    private bool captureFormatInitialized;
    private uint processId;
    private bool processInclusive;
    private Action<List<byte>, List<byte>, object> callbackFunc;
    private object callbackFuncUserData;
    private uint callbackInterval;
    private List<byte> audioData;
    private double maxExecutionTime;
    private uint mainThreadBytesToSkip;

    public ProcessLoopbackCapture()
    {
        captureState = eCaptureState.READY;
        runAudioThreads = false;
        audioData = new List<byte>();
        maxExecutionTime = 0.0;
        mainThreadBytesToSkip = 0;
    }

    public eCaptureError SetCaptureFormat(uint sampleRate, ushort bitDepth, ushort channelCount, ushort formatTag = 1)
    {
        if (captureState != eCaptureState.READY) return eCaptureError.STATE;
        if (sampleRate < 1000) return eCaptureError.PARAM;
        if (bitDepth == 0 || bitDepth > 32 || (bitDepth % 8) != 0) return eCaptureError.PARAM;
        if (channelCount < 1 || channelCount > 1024) return eCaptureError.PARAM;

        captureFormat = new WAVEFORMATEX
        {
            wFormatTag = formatTag,
            nChannels = channelCount,
            nSamplesPerSec = sampleRate,
            wBitsPerSample = bitDepth,
            nBlockAlign = (ushort)(bitDepth / 8 * channelCount),
            nAvgBytesPerSec = sampleRate * (uint)(bitDepth / 8 * channelCount)
        };

        captureFormatInitialized = true;
        return eCaptureError.NONE;
    }

    public bool CopyCaptureFormat(out WAVEFORMATEX format)
    {
        if (!captureFormatInitialized)
        {
            format = new WAVEFORMATEX();
            return false;
        }

        format = captureFormat;
        return true;
    }

    public eCaptureError SetTargetProcess(uint processId, bool inclusive = true)
    {
        if (captureState != eCaptureState.READY) return eCaptureError.STATE;
        if (processId == 0) return eCaptureError.PARAM;

        this.processId = processId;
        this.processInclusive = inclusive;
        return eCaptureError.NONE;
    }

    public eCaptureError SetCallback(Action<List<byte>, List<byte>, object> callbackFunc, object userData = null)
    {
        if (captureState != eCaptureState.READY) return eCaptureError.STATE;

        this.callbackFunc = callbackFunc;
        this.callbackFuncUserData = userData;
        return eCaptureError.NONE;
    }

    public eCaptureError SetCallbackInterval(uint interval)
    {
        if (captureState != eCaptureState.READY) return eCaptureError.STATE;

        callbackInterval = interval < 1 ? 1 : interval;
        return eCaptureError.NONE;
    }

    public eCaptureState GetState()
    {
        return captureState;
    }

    public eCaptureError StartCapture()
    {
        if (captureState != eCaptureState.READY) return eCaptureError.STATE;
        if (!captureFormatInitialized) return eCaptureError.FORMAT;
        if (processId == 0) return eCaptureError.PROCESSID;

        AUDIOCLIENT_ACTIVATION_PARAMS activationParams = new AUDIOCLIENT_ACTIVATION_PARAMS
        {
            ActivationType = 1, // AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK
            ProcessLoopbackParams = new ProcessLoopbackParams
            {
                ProcessLoopbackMode = processInclusive ? 1 : 0, // PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE or PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE
                TargetProcessId = processId
            }
        };

        PROPVARIANT propVariant = new PROPVARIANT
        {
            vt = 65, // VT_BLOB
            cbSize = (uint)Marshal.SizeOf(activationParams),
            blob = Marshal.AllocHGlobal(Marshal.SizeOf(activationParams))
        };
        Marshal.StructureToPtr(activationParams, propVariant.blob, false);

        IActivateAudioInterfaceAsyncOperation activationOperation = null;
        ActivateAudioInterfaceAsyncCallback asyncCallback = new ActivateAudioInterfaceAsyncCallback();
        lastError = ActivateAudioInterfaceAsync("VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK", typeof(IAudioClient).GUID, ref propVariant, asyncCallback, out activationOperation);
        if (lastError != 0)
        {
            Reset();
            return eCaptureError.DEVICE;
        }

        asyncCallback.Wait();
        lastError = activationOperation.GetActivateResult(out lastError, out object activatedInterface);
        activationOperation.Release();
        if (lastError != 0)
        {
            Reset();
            return eCaptureError.ACTIVATION;
        }

        audioClient = (IAudioClient)activatedInterface;
        lastError = audioClient.Initialize(1, 0x00020000 | 0x00040000 | 0x08000000 | 0x10000000, 0, 0, ref captureFormat, IntPtr.Zero);
        if (lastError != 0)
        {
            Reset();
            return eCaptureError.INITIALIZE;
        }

        Guid IID_IAudioCaptureClient = typeof(IAudioCaptureClient).GUID;
        lastError = audioClient.GetService(ref IID_IAudioCaptureClient, out object audioCaptureClientObj);
        if (lastError != 0)
        {
            Reset();
            return eCaptureError.SERVICE;
        }
        audioCaptureClient = (IAudioCaptureClient)audioCaptureClientObj;

        sampleReadyEvent = CreateEvent(IntPtr.Zero, false, false, null);
        lastError = audioClient.SetEventHandle(sampleReadyEvent);
        if (lastError != 0 || sampleReadyEvent == IntPtr.Zero)
        {
            Reset();
            return eCaptureError.EVENT;
        }

        lastError = audioClient.Start();
        if (lastError != 0)
        {
            Reset();
            return eCaptureError.START;
        }

        StartThreads(0.0);
        captureState = eCaptureState.CAPTURING;
        return eCaptureError.NONE;
    }

    public eCaptureError StopCapture()
    {
        if (captureState == eCaptureState.READY) return eCaptureError.STATE;
        Reset();
        return eCaptureError.NONE;
    }

    public eCaptureError PauseCapture()
    {
        if (captureState != eCaptureState.CAPTURING) return eCaptureError.STATE;
        captureState = eCaptureState.PAUSED;
        lastError = audioClient.Stop();
        if (lastError != 0) return eCaptureError.STOP;
        StopThreads();
        return eCaptureError.NONE;
    }

    public eCaptureError ResumeCapture(double initialDurationToSkip = 0.1)
    {
        if (captureState != eCaptureState.PAUSED) return eCaptureError.STATE;
        captureState = eCaptureState.CAPTURING;
        ResetEvent(sampleReadyEvent);
        lastError = audioClient.Start();
        if (lastError != 0) return eCaptureError.START;
        StartThreads(initialDurationToSkip);
        return eCaptureError.NONE;
    }

    public int GetLastErrorResult()
    {
        return lastError;
    }

    public double GetMaxExecutionTime()
    {
        return maxExecutionTime;
    }

    public void ResetMaxExecutionTime()
    {
        maxExecutionTime = 0.0;
    }

    public eCaptureError GetQueueSize(out int size)
    {
        size = 0;
        return eCaptureError.NOT_AVAILABLE;
    }

    private void Reset()
    {
        StopThreads();
        if (captureState == eCaptureState.CAPTURING)
        {
            audioClient.Stop();
        }
        if (audioCaptureClient != null)
        {
            Marshal.ReleaseComObject(audioCaptureClient);
            audioCaptureClient = null;
        }
        if (audioClient != null)
        {
            audioClient.Reset();
            Marshal.ReleaseComObject(audioClient);
            audioClient = null;
        }
        if (sampleReadyEvent != IntPtr.Zero)
        {
            CloseHandle(sampleReadyEvent);
            sampleReadyEvent = IntPtr.Zero;
        }
        captureState = eCaptureState.READY;
    }

    private void StartThreads(double initialDurationToSkip)
    {
        if (runAudioThreads) return;
        if (initialDurationToSkip < 0.0) initialDurationToSkip = 0.0;
        mainThreadBytesToSkip = (uint)(captureFormat.nSamplesPerSec * initialDurationToSkip) * captureFormat.nBlockAlign;
        runAudioThreads = true;
        mainAudioThread = new Thread(ProcessMainToCallback);
        mainAudioThread.Start();
    }

    private void StopThreads()
    {
        if (!runAudioThreads) return;
        runAudioThreads = false;
        mainAudioThread.Join();
        mainAudioThread = null;
        audioData.Clear();
    }

    private void ProcessMainToCallback()
    {
        uint taskIndex = 0;
        IntPtr taskHandle = AvSetMmThreadCharacteristics("Pro Audio", ref taskIndex);
        IntPtr pData = IntPtr.Zero;
        uint framesAvailable;
        uint bytesAvailable;
        uint captureFlags;
        while (runAudioThreads)
        {
            if (WaitForSingleObject(sampleReadyEvent, 50) == 0)
            {
                if (!runAudioThreads) break;
                var tickStart = DateTime.Now;
                while (audioCaptureClient.GetBuffer(out pData, out framesAvailable, out captureFlags, out _, out _) == 0)
                {
                    bytesAvailable = framesAvailable * captureFormat.nBlockAlign;
                    for (uint i = 0; i < bytesAvailable; ++i)
                    {
                        if (mainThreadBytesToSkip != 0)
                        {
                            --mainThreadBytesToSkip;
                        }
                        else
                        {
                            audioData.Add(Marshal.ReadByte(pData, (int)i));
                        }
                    }
                    audioCaptureClient.ReleaseBuffer(framesAvailable);
                }
                if (audioData.Count > 0)
                {
                    callbackFunc?.Invoke(audioData, audioData, callbackFuncUserData);
                    audioData.Clear();
                }
                var duration = (DateTime.Now - tickStart).TotalMilliseconds;
                if (duration > maxExecutionTime) maxExecutionTime = duration;
            }
        }
        if (taskHandle != IntPtr.Zero) AvRevertMmThreadCharacteristics(taskHandle);
    }

    [DllImport("api-ms-win-core-synch-l1-2-0.dll", SetLastError = true)]
    private static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string lpName);

    [DllImport("api-ms-win-core-com-l1-1-0.dll", SetLastError = true)]
    private static extern int ActivateAudioInterfaceAsync(string deviceInterfacePath, Guid riid, ref PROPVARIANT activationParams, IActivateAudioInterfaceCompletionHandler completionHandler, out IActivateAudioInterfaceAsyncOperation activationOperation);

    [DllImport("avrt.dll", SetLastError = true)]
    private static extern IntPtr AvSetMmThreadCharacteristics(string taskName, ref uint taskIndex);

    [DllImport("avrt.dll", SetLastError = true)]
    private static extern bool AvRevertMmThreadCharacteristics(IntPtr taskHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool ResetEvent(IntPtr hEvent);
}