// 
// These are functions to interface with the MadCodeHook library,
// which can hook any API calls on the same process or system wide.
// Includes event handlers, so that managed code can set handlers
// and listen for when an API function is called.
// http://www.madshi.net/madCodeHookDescription.htm
//

#include <windows.h>
#include "madCHook.h"

//
// BltBlt listeners
//

typedef void (__stdcall *BitBltEventHandler)(HDC hdcDest, int nXDest, int nYDest,
						   int nWidth, int nHeight, HDC hdcSrc, int nXSrc, int nYSrc, DWORD rowOp);

BitBltEventHandler bitBltListener;

extern "C" __declspec(dllexport) void WINAPI SetBitBltListener(PVOID functionPointer){
	bitBltListener=(BitBltEventHandler)functionPointer;
}

BOOL (WINAPI *BitBltNextHook)(HDC hdcDest, int nXDest, int nYDest,
	int nWidth, int nHeight, HDC hdcSrc, int nXSrc, int nYSrc, DWORD dwRop);

BOOL WINAPI BitBltHookProc(HDC hdcDest, int nXDest, int nYDest,
						   int nWidth, int nHeight, HDC hdcSrc, int nXSrc, int nYSrc, DWORD dwRop){	
	
	//int  *p=NULL;
	//(*p)++;

	// Call original
    BOOL result = BitBltNextHook(hdcDest, nXDest, nYDest, nWidth, nHeight, hdcSrc, nXSrc,
		nYSrc, dwRop);
	// Notify listeners after the original has been called
	if (bitBltListener!=NULL)
		(*bitBltListener)(hdcDest, nXDest, nYDest, nWidth, nHeight, hdcSrc, nXSrc,
		nYSrc, dwRop);

	return result;
}

//
// GetWindowDC listeners
//

typedef void (__stdcall *GetWindowDCEventHandler)(HWND window, HDC result);

GetWindowDCEventHandler getWindowDCListener;

extern "C" __declspec(dllexport) void WINAPI SetGetWindowDCListener(PVOID functionPointer){
	getWindowDCListener=(GetWindowDCEventHandler)functionPointer;
}

HDC(WINAPI *GetWindowDCNextHook)(HWND window);

HDC WINAPI GetWindowDCHookProc(HWND window){
	// Call original
    HDC result = GetWindowDCNextHook(window);
	// Notify listeners
	if (getWindowDCListener!=NULL)
		(*getWindowDCListener)(window, result);

	return result;
}


//
// ScrollDC Listeners
//

typedef void (__stdcall *ScrollDCEventHandler)(HDC hdcDest, int dx, int dy,
						   const RECT* lprcScroll, const RECT* lprcClip, HRGN hrgnUpdate, LPRECT lprcUpdate);

ScrollDCEventHandler scrollDCListener;

extern "C" __declspec(dllexport) void WINAPI SetScrollDCListener(PVOID functionPointer){
	scrollDCListener=(ScrollDCEventHandler)functionPointer;
}
BOOL (WINAPI *ScrollDCNextHook)(HDC hdc, int dx, int dy,
	const RECT* lprcScroll, const RECT* lprcClip, HRGN hrgnUpdate, LPRECT lprcUpdate);

BOOL WINAPI ScrollDCHookProc(HDC hdc, int dx, int dy,
							  const RECT* lprcScroll, const RECT* lprcClip, HRGN hrgnUpdate, LPRECT lprcUpdate){
	BOOL result = ScrollDCNextHook(hdc,dx,dy,lprcScroll,lprcClip,hrgnUpdate,lprcUpdate);
	if (scrollDCListener!=NULL)
		(*scrollDCListener)(hdc,dx,dy,lprcScroll,lprcClip,hrgnUpdate,lprcUpdate);
	return result;
}

//
// InvalidateRect Listeners
//

typedef void (__stdcall *InvalidateRectEventHandler)(HWND hwnd, CONST RECT* lpRect, BOOL bErase);

InvalidateRectEventHandler invalidateRectListener;

extern "C" __declspec(dllexport) void WINAPI SetInvalidateRectListener(PVOID functionPointer){
	invalidateRectListener=(InvalidateRectEventHandler)functionPointer;
}
BOOL (WINAPI *InvalidateRectNextHook)(HWND hwnd, CONST RECT* lpRect, BOOL bErase);

BOOL WINAPI InvalidateRectHookProc(HWND hwnd, CONST RECT* lpRect, BOOL bErase){
	BOOL result = InvalidateRectNextHook(hwnd, lpRect, bErase);
	if (invalidateRectListener!=NULL)
		(*invalidateRectListener)(hwnd, lpRect, bErase);
	return result;
}

// Install hooks
extern "C" __declspec(dllexport) BOOL WINAPI Hook(){
	int result1 = HookAPI("gdi32.dll", "BitBlt", BitBltHookProc, (PVOID*) &BitBltNextHook);
	int result2 = HookAPI("user32.dll", "GetWindowDC", GetWindowDCHookProc, (PVOID*) &GetWindowDCNextHook);
	int result3 = HookAPI("user32.dll","ScrollDC",ScrollDCHookProc, (PVOID*) &ScrollDCNextHook);
	int result4 = HookAPI("user32.dll","InvalidateRect",InvalidateRectHookProc, (PVOID*) &InvalidateRectNextHook);
	// -1 = success
	return (result1==-1 && result2==-1 && result3==-1);	
}

// Uninstall hooks
extern "C" __declspec(dllexport) BOOL WINAPI UnHook(){
	int result1 = UnhookAPI((PVOID*) &BitBltNextHook);
	int result2 = UnhookAPI((PVOID*) &GetWindowDCNextHook);
	int result3 = UnhookAPI((PVOID*) &ScrollDCNextHook);
	int result4 = UnhookAPI((PVOID*) &InvalidateRectNextHook);
	// -1 = success
	return (result1==-1 && result2==-1 && result3==1 && result4==1);	
}


/*
 * This is the original method included in the sample. Use for reference.
/*

/*
// variable for the "next hook", which we then call in the callback function
// it must have *exactly* the same parameters and calling convention as the
// original function
// besides, it's also the parameter that you need to undo the code hook again
UINT (WINAPI *WinExecNextHook)(LPCSTR lpCmdLine, UINT uCmdShow);

// this function is our hook callback function, which will receive
// all calls to the original SomeFunc function, as soon as we've hooked it
// the hook function must have *exactly* the same parameters and calling
// convention as the original function
UINT WINAPI WinExecHookProc(LPCSTR lpCmdLine, UINT uCmdShow)
{
  // check the input parameters and ask whether the call should be executed
  if (MessageBox(0, lpCmdLine, "Execute?", MB_YESNO | MB_ICONQUESTION) == IDYES)
    // it shall be executed, so let's do it
    return WinExecNextHook(lpCmdLine, uCmdShow);
  else
    // we don't execute the call, but we should at least return a valid value
    return ERROR_ACCESS_DENIED;
}
int WINAPI WinMain(HINSTANCE hInstance,
                   HINSTANCE hPrevInstance,
                   LPSTR     lpCmdLine,
                   int       nCmdShow)
{
  // InitializeMadCHook is needed only if you're using the static madCHook.lib
  InitializeMadCHook();

  // we install our hook on the API...
  // please note that in this demo the hook only effects our own process!
  HookAPI("kernel32.dll", "WinExec", WinExecHookProc, (PVOID*) &WinExecNextHook);

  // now call the original (but hooked) API
  // as a result of the hook the user will receive our messageBox etc
  WinExec("notepad.exe", SW_SHOWNORMAL);
  // we like clean programming, don't we?
  // so we cleanly unhook again
  UnhookAPI((PVOID*) &WinExecNextHook);

  // FinalizeMadCHook is needed only if you're using the static madCHook.lib
  FinalizeMadCHook();

  return true;
}
*/