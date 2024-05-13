#pragma once
extern "C" {
    __declspec(dllexport) int add(int a, int b)
    {
        SampleArithmetic obj;
        return obj.Add(a, b);
    }

    __declspec(dllexport) int subtract(int a, int b)
    {
        SampleArithmetic obj;
        return obj.Subtract(a, b);
    }

    __declspec(dllexport) int multiply(int a, int b)
    {
        SampleArithmetic obj;
        return obj.Multiply(a, b);
    }
}
