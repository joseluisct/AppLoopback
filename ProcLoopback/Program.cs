using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public class Program
{
    private static ProcessLoopbackCapture g_LoopbackCapture = new ProcessLoopbackCapture();
    private static object g_AudioDataLock = new object();
    private static List<byte> g_Data = new List<byte>();

    private const bool WRITE_RAW_FILE = false;
    private const uint DEFAULT_SAMPLE_RATE = 44100U;
    private const ushort DEFAULT_BIT_DEPTH = 16;
    private const ushort DEFAULT_CHANNEL_COUNT = 2;

    public static void Main(string[] args)
    {
        if (CoInitializeEx(IntPtr.Zero, COINIT_MULTITHREADED) != 0)
        {
            Console.WriteLine("Failed to init COM");
            Environment.Exit(1);
        }

        uint sampleRate = DEFAULT_SAMPLE_RATE;
        ushort bitDepth = DEFAULT_BIT_DEPTH;
        ushort channelCount = DEFAULT_CHANNEL_COUNT;

        if (args.Length >= 1)
        {
            if (!uint.TryParse(args[0], out sampleRate))
            {
                sampleRate = DEFAULT_SAMPLE_RATE;
            }
        }

        if (args.Length >= 2)
        {
            if (!ushort.TryParse(args[1], out bitDepth))
            {
                bitDepth = DEFAULT_BIT_DEPTH;
            }
        }

        if (args.Length >= 3)
        {
            if (!ushort.TryParse(args[2], out channelCount))
            {
                channelCount = DEFAULT_CHANNEL_COUNT;
            }
        }

        Console.WriteLine($"Sample Rate: {sampleRate}");
        Console.WriteLine($"Bit Depth : {bitDepth}");
        Console.WriteLine($"Channels  : {channelCount}");
        Console.WriteLine();

        bool runApplication = true;
        while (runApplication)
        {
            g_Data.Clear();

            uint processId = 0;
            do
            {
                Console.WriteLine("Enter the Process Name to listen to (incl. .exe):");
                string processName = Console.ReadLine();
                if (string.IsNullOrEmpty(processName)) continue;

                List<uint> processIds = FindParentProcessIDs(processName);
                if (processIds.Count > 0) processId = processIds[0];
            }
            while (processId == 0);

            Console.WriteLine($"PID: {processId}");

            g_LoopbackCapture.SetCaptureFormat(sampleRate, bitDepth, channelCount, 1);
            g_LoopbackCapture.SetTargetProcess(processId, true);
            g_LoopbackCapture.SetCallback(OnDataCapture);
            g_LoopbackCapture.SetCallbackInterval(40);

            eCaptureError eError = g_LoopbackCapture.StartCapture();
            if (eError != eCaptureError.NONE)
            {
                int hr = g_LoopbackCapture.GetLastErrorResult();
                Console.WriteLine();
                Console.WriteLine($"ERROR ({(int)eError}): {GetErrorText(eError)}");
                Console.WriteLine($"HR: 0x{hr:X}");
                Console.WriteLine($"HR Text: {new System.ComponentModel.Win32Exception(hr).Message}");
                Console.WriteLine();
                continue;
            }

            Console.WriteLine("Capturing audio.");
            Console.WriteLine("Press Enter to stop and save.");
            Console.WriteLine("Type \"discard\" to stop without saving.");
            Console.WriteLine("Type \"pause\" to pause or resume capture.");
            Console.WriteLine("Type \"hang\" to simulate a long hang in the callback.");
            Console.WriteLine("Type \"exit\" to exit the application.");

            while (true)
            {
                string input = Console.ReadLine();
                if (input.Equals("discard", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
                if (input.Equals("pause", StringComparison.OrdinalIgnoreCase))
                {
                    if (g_LoopbackCapture.GetState() == eCaptureState.CAPTURING)
                    {
                        eCaptureError pauseError = g_LoopbackCapture.PauseCapture();
                        if (pauseError != eCaptureError.NONE)
                        {
                            int hr = g_LoopbackCapture.GetLastErrorResult();
                            Console.WriteLine();
                            Console.WriteLine($"ERROR ({(int)pauseError}): {GetErrorText(pauseError)}");
                            Console.WriteLine($"HR: 0x{hr:X}");
                            Console.WriteLine($"HR Text: {new System.ComponentModel.Win32Exception(hr).Message}");
                            Console.WriteLine();
                        }
                        else
                        {
                            Console.WriteLine("Capture paused");
                        }
                    }
                    else if (g_LoopbackCapture.GetState() == eCaptureState.PAUSED)
                    {
                        eCaptureError resumeError = g_LoopbackCapture.ResumeCapture();
                        if (resumeError != eCaptureError.NONE)
                        {
                            int hr = g_LoopbackCapture.GetLastErrorResult();
                            Console.WriteLine();
                            Console.WriteLine($"ERROR ({(int)resumeError}): {GetErrorText(resumeError)}");
                            Console.WriteLine($"HR: 0x{hr:X}");
                            Console.WriteLine($"HR Text: {new System.ComponentModel.Win32Exception(hr).Message}");
                            Console.WriteLine();
                        }
                        else
                        {
                            Console.WriteLine("Capture resumed");
                        }
                    }
                    continue;
                }
                if (input.Equals("hang", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Hanging Callback Thread for 10 seconds ...");
                    lock (g_AudioDataLock)
                    {
                        Thread.Sleep(10000);
                    }
                    Console.WriteLine("Done.");
                    continue;
                }
                if (input.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    runApplication = false;
                    break;
                }
                if (string.IsNullOrEmpty(input))
                {
                    g_LoopbackCapture.StopCapture();
                    string fileName = $"out-{Environment.TickCount}.wav";
                    Console.WriteLine($"Saving Audio to \"{fileName}\" ...");
                    g_LoopbackCapture.CopyCaptureFormat(out WAVEFORMATEX format);
                    WriteWavFile(fileName, g_Data, format);
                    Console.WriteLine("Done");
                    break;
                }
            }
            g_LoopbackCapture.StopCapture();
        }

        CoUninitialize();
    }

    private static void OnDataCapture(List<byte> data, List<byte> _, object userData)
    {
        lock (g_AudioDataLock)
        {
            g_Data.AddRange(data);
        }
    }

    private static void WriteWavFile(string fileName, List<byte> data, WAVEFORMATEX format)
    {
        using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
        {
            using (BinaryWriter writer = new BinaryWriter(fs))
            {
                if (!WRITE_RAW_FILE)
                {
                    writer.Write(Encoding.ASCII.GetBytes("RIFF"));
                    writer.Write(0); // Placeholder for file size
                    writer.Write(Encoding.ASCII.GetBytes("WAVE"));
                    writer.Write(Encoding.ASCII.GetBytes("fmt "));
                    writer.Write(Marshal.SizeOf(format));
                    writer.Write(format.wFormatTag);
                    writer.Write(format.nChannels);
                    writer.Write(format.nSamplesPerSec);
                    writer.Write(format.nAvgBytesPerSec);
                    writer.Write(format.nBlockAlign);
                    writer.Write(format.wBitsPerSample);
                    writer.Write(format.cbSize);
                    writer.Write(Encoding.ASCII.GetBytes("data"));
                    writer.Write(data.Count);
                    writer.Write(data.ToArray());

                    long fileSize = fs.Length;
                    fs.Seek(4, SeekOrigin.Begin);
                    writer.Write((int)(fileSize - 8));
                    fs.Seek(40, SeekOrigin.Begin);
                    writer.Write(data.Count);
                }
                else
                {
                    writer.Write(data.ToArray());
                }
            }
        }
    }

    private static List<uint> FindParentProcessIDs(string executableName)
    {
        List<uint> processIdList = new List<uint>();
        foreach (Process process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(executableName)))
        {
            processIdList.Add((uint)process.Id);
        }
        return processIdList;
    }

    private static string GetErrorText(eCaptureError error)
    {
        return error switch
        {
            eCaptureError.NONE => "Success",
            eCaptureError.PARAM => "Invalid parameter",
            eCaptureError.STATE => "Invalid operation for current state",
            eCaptureError.NOT_AVAILABLE => "Feature not available",
            eCaptureError.FORMAT => "CaptureFormat is invalid or not initialized",
            eCaptureError.PROCESSID => "ProcessId is invalid (0/not set)",
            eCaptureError.DEVICE => "Failed to get device",
            eCaptureError.ACTIVATION => "Failed to activate device",
            eCaptureError.INITIALIZE => "Failed to init device",
            eCaptureError.SERVICE => "Failed to get interface pointer via service",
            eCaptureError.START => "Failed to start capture",
            eCaptureError.STOP => "Failed to stop capture",
            eCaptureError.EVENT => "Failed to create and set event",
            eCaptureError.INTERFACE => "Failed to call Windows interface function",
            _ => "Unknown"
        };
    }

    [DllImport("ole32.dll")]
    private static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [DllImport("ole32.dll")]
    private static extern void CoUninitialize();

    private const uint COINIT_MULTITHREADED = 0x0;
    private const uint COINIT_APARTMENTTHREADED = 0x2;
}