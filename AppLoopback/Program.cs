using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace AudioPrueba;

class Program
{
    //[MTAThread]
    static async Task Main(string[] args)
    {
        if (args.Length != 3)
        {
            Usage();
            return;
        }

        int processId = int.Parse(args[0]);
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
        //var loopbackCapture = new CLoopbackCapture(processId, includeProcessTree);
        try
        {
            StartCapture(processId, includeProcessTree, outputFileName);
            Console.WriteLine("Capturing 10 seconds of audio.");
            await Task.Delay(10000);
            StopCapture();
            Console.WriteLine("Finished.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to start capture\n0x{ex.HResult:X}: {ex.Message}");
        }
    }

    static void Usage()
    {
        Console.WriteLine(
            "Usage: ApplicationLoopback <pid> <includetree|excludetree> <outputfilename>\n" +
            "\n" +
            "<pid> is the process ID to capture or exclude from capture\n" +
            "includetree includes audio from that process and its child processes\n" +
            "excludetree includes audio from all processes except that process and its child processes\n" +
            "<outputfilename> is the WAV file to receive the captured audio (10 seconds)\n" +
            "\n" +
            "Examples:\n" +
            "\n" +
            "ApplicationLoopback 1234 includetree CapturedAudio.wav\n" +
            "\n" +
            "  Captures audio from process 1234 and its children.\n" +
            "\n" +
            "ApplicationLoopback 1234 excludetree CapturedAudio.wav\n" +
            "\n" +
            "  Captures audio from all processes except process 1234 and its children.\n");
    }

    [DllImport("Wrapper.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern int StopCapture();
    [DllImport("Wrapper.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern int StartCapture(int processId, bool include, string output);
}