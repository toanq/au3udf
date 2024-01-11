#include <windows.h>

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID reserved)
{
    switch (reason) {
        case DLL_PROCESS_ATTACH:
        case DLL_PROCESS_DETACH:
        case DLL_THREAD_ATTACH:
        case DLL_THREAD_DETACH:
            break;
    }

    return 1;
}

int __declspec(dllexport) add(int a, int b)
{
    return a + b;
}

DWORD __stdcall test(void* param)
{
    MessageBoxA(NULL, "title", "content", 0);
    return 0;
}

void __declspec(dllexport) newThread()
{
    CreateThread(NULL, 0, test, NULL, 0, NULL);
}

int __declspec(dllexport) externalAdd(int max)
{
    int i, s = 0;
    for (i=0;i<max;i++)
        s+=i;

    return s;
}