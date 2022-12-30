// Hi there.
// This file is a regrettable amalgamation of ancient Windows APIs that, for better or worse, works.
// This code is terrible, but making it better is honestly not something I am interested in doing,
// since at the end of the day, it will still involve low-level Win32 nonsense anyway.

// You have been warned.
// -----

// Helpful: https://hotcashew.com/2013/10/extracting-icons/

#include <Windows.h>
#include <olectl.h>
#include <ShObjIdl.h> //For IShellItemImageFactory
#include <Shlwapi.h>
#include <stdio.h>

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "Shlwapi")
#pragma comment(lib, "Gdiplus")

void DisplayUsage()
{
    wprintf(L"Usage:\n");
    wprintf(L"ImageFactorySample.exe <size> <path-to-input> <path-to-output>\n");
    wprintf(L"e.g. ImageFactorySample.exe 256 C:/some/path/program.exe C:/some/path/program.png\n");
}

// https://stackoverflow.com/a/24645420

#include <gdiplus.h>
using namespace Gdiplus;

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
    UINT  num = 0;          // number of image encoders
    UINT  size = 0;         // size of the image encoder array in bytes

    ImageCodecInfo* pImageCodecInfo = NULL;

    GetImageEncodersSize(&num, &size);
    if (size == 0)
        return -1;  // Failure

    pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
    if (pImageCodecInfo == NULL)
        return -1;  // Failure

    GetImageEncoders(num, size, pImageCodecInfo);

    // List all available encoders:
    // for (UINT j = 0; j < num; ++j)
    //     wprintf(L"%ls\n", pImageCodecInfo[j].MimeType);

    for (UINT j = 0; j < num; ++j)
    {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0)
        {
            *pClsid = pImageCodecInfo[j].Clsid;
            free(pImageCodecInfo);
            return j;  // Success
        }
    }
    free(pImageCodecInfo);
    return -1;  // Failure
}

bool WriteToImageFile(Bitmap* image, LPCWSTR filePath, LPCWSTR formatMimeType)
{
    CLSID myClsId;
    int retVal = GetEncoderClsid(formatMimeType, &myClsId);
    if (retVal == -1) return false;

    image->Save(filePath, &myClsId, nullptr);

    return true;
}

int wmain(int argc, wchar_t *argv[])
{
    if (argc != 4)
    {
        DisplayUsage();
    }
    else
    {
        int nSize = _wtoi(argv[1]);
        if (!nSize)
        {
            DisplayUsage();
        }
        else
        {
            GdiplusStartupInput gdiplusStartupInput;
            ULONG_PTR gdiplusToken;
            GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);

            LPCWSTR pwszError{};

            HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
            if (SUCCEEDED(hr))
            {
                // Getting the IShellItemImageFactory interface pointer for the file.
                IShellItemImageFactory *pImageFactory;
                hr = SHCreateItemFromParsingName(argv[2], nullptr, IID_PPV_ARGS(&pImageFactory));
                if (SUCCEEDED(hr))
                {
                    SIZE size = { nSize, nSize };

                    HBITMAP hbitmap;
                    hr = pImageFactory->GetImage(size, SIIGBF_BIGGERSIZEOK|SIIGBF_ICONONLY|SIIGBF_SCALEUP, &hbitmap);
                    if (SUCCEEDED(hr))
                    {
                        // https://stackoverflow.com/questions/69344394/is-there-a-way-to-save-a-png-from-bitmap-with-transparency-gdi

                        //Creating nessecery palettes for gdi, so i can use Bitmap::FromHBITMAP
                        PALETTEENTRY pat = { 255,255,255,0 };
                        LOGPALETTE logpalt = { 0, 1, pat };
                        HPALETTE hpalette = CreatePalette(&logpalt);

                        //Getting data from the bitmap and creating new bitmap from it
                        auto bitmap = Bitmap::FromHBITMAP(hbitmap, hpalette);
                        Rect rect(0, 0, bitmap->GetWidth(), bitmap->GetHeight());
                        BitmapData bd;
                        bitmap->LockBits(&rect, ImageLockModeRead, bitmap->GetPixelFormat(), &bd);
                        Bitmap* bitmapWithAlpha = new Bitmap(bd.Width, bd.Height, bd.Stride, PixelFormat32bppARGB, (BYTE*)bd.Scan0);
                        bitmap->UnlockBits(&bd);

                        if (!WriteToImageFile(bitmapWithAlpha, argv[3], L"image/png"))
                        {
                            wprintf(L"Failed to write to file :(\n");
                        }

                        delete bitmapWithAlpha;
                        delete bitmap;
                        DeleteObject(hbitmap);
                    }
                    else
                    {
                        pwszError = L"IShellItemImageFactory::GetImage failed with error code %x";
                    }

                    pImageFactory->Release();
                }
                else
                {
                    pwszError = L"SHCreateItemFromParsingName failed with error %x";
                }

                CoUninitialize();
            }
            else
            {
                pwszError = L"CoInitializeEx failed with error code %x";
            }

            if (FAILED(hr) && pwszError)
            {
                wprintf(pwszError, hr);
            }

            GdiplusShutdown(gdiplusToken);
        }
    }

    return 0;
}
