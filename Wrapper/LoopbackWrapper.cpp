#include "pch.h"
#include "../NativeLib/LoopbackCapture.h"
#include <string>

extern "C" {
    CLoopbackCapture obj;
    __declspec(dllexport) int StopCapture()
    {
        return obj.StopCaptureAsync();
    }

    __declspec(dllexport) int StartCapture(int processId, bool include, const char* output)
    {
        const size_t cSize = strlen(output) + 1;
        wchar_t* wc = new wchar_t[cSize];
        size_t tmp = 0;
        mbstowcs_s(&tmp, wc, cSize, output, cSize);
        return obj.StartCaptureAsync(processId, include, wc);
    }
}