using NAudio.CoreAudioApi;
using NAudio.Wave;
using ProcLoopback;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

public class Program
{
    static WasapiCapture captureDevice = null;
    static WaveFileWriter writer;

    private static void CleanupCapture()
    {
        if (captureDevice != null)
        {
            captureDevice.DataAvailable -= CaptureDevice_DataAvailable;
            captureDevice.RecordingStopped -= CaptureDevice_RecordingStopped;
            captureDevice.Dispose();
            captureDevice = null;
        }
    }

    private static void CaptureDevice_RecordingStopped(object sender, StoppedEventArgs e)
    {
        writer.Dispose();
        writer = null;
        CleanupCapture();
    }

    private static void CaptureDevice_DataAvailable(object sender, NAudio.Wave.WaveInEventArgs e)
    {
        if(e.BytesRecorded > 0)
        {
            writer.Write(e.Buffer, 0, e.BytesRecorded);
        }
    }

    public static async Task Main(string[] args)
    {
        if (args.Length != 3)
        {
            Usage();
            return;
        }
        Process process = null;
        bool isPid = int.TryParse(args[0], out int processId);
        if (!isPid)
        {
            var handle = Process.GetProcessesByName(args[0]).FirstOrDefault().Handle;
            process = ParentProcessUtilities.GetParentProcess(handle);
        }
        else process = ParentProcessUtilities.GetParentProcess(processId);

        bool includeProcessTree;
        if (args[1] == "includetree")
        {
            includeProcessTree = true;
        }
        else if (args[1] == "excludetree")
        {
            includeProcessTree = false;
        }
        else
        {
            Usage();
            return;
        }

        string outputFileName = args[2];
        try
        {
            if (captureDevice == null)
            {
                captureDevice = await WasapiCapture.CreateForProcessCaptureAsync(process.Id, includeProcessTree);
                captureDevice.DataAvailable += CaptureDevice_DataAvailable;
                captureDevice.RecordingStopped += CaptureDevice_RecordingStopped;
                var file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, outputFileName);
                writer = new WaveFileWriter(file, captureDevice.WaveFormat);
                captureDevice.StartRecording();
                Console.WriteLine($"Process: {process.ProcessName} - PID: {process.Id}");
                Console.WriteLine("Capturing 10 seconds of audio.");

                await Task.Delay(10000);
                captureDevice.StopRecording();
                Console.WriteLine("Finished.");
                await Task.Delay(1000);
            }
            else
            {
                captureDevice.StopRecording(); // WAV File will be completed in recording stopped
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to start capture\n0x{ex.HResult:X}: {ex.Message}");
            CleanupCapture();
        }    
    }

    static void Usage()
    {
        Console.WriteLine(
            "Usage: ApplicationLoopback <pid|app.exe> <includetree|excludetree> <outputfilename>\n" +
            "\n" +
            "<pid> is the process ID to capture or exclude from capture\n" +
            "<app.exe> is the process name to capture or exclude from capture\n" +
            "includetree includes audio from that process and its child processes\n" +
            "excludetree includes audio from all processes except that process and its child processes\n" +
            "<outputfilename> is the WAV file to receive the captured audio (10 seconds)\n" +
            "\n" +
            "Examples:\n" +
            "\n" +
            "ApplicationLoopback 1234 includetree CapturedAudio.wav\n" +
            "\n" +
            "  Captures audio from parent process of 1234 and its children.\n" +
            "\n" +
            "ApplicationLoopback msedge.exe includetree CapturedAudio.wav\n" +
            "\n" +
            "  Captures audio from parent process of msedge and its children.\n" +
            "\n" +
            "ApplicationLoopback 1234 excludetree CapturedAudio.wav\n" +
            "\n" +
            "  Captures audio from all processes except parent process of 1234 and its children.\n");
    }
}