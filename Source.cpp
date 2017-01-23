#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "shlwapi")
#pragma comment(lib, "comctl32")
#pragma comment(lib, "version")

#include <windows.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <stdio.h>

HWND hList;
TCHAR szClassName[] = TEXT("Window");

BOOL IsTargetFile(LPCWSTR lpszFilePath)
{
	WCHAR szExtList[] = L"*.exe;*.dll;";
	LPCWSTR seps = L";";
	WCHAR *next;
	LPWSTR token = wcstok_s(szExtList, seps, &next);
	while (token != NULL)
	{
		if (PathMatchSpecW(lpszFilePath, token))
		{
			return TRUE;
		}
		token = wcstok_s(0, seps, &next);
	}
	return FALSE;
}

VOID TargetFileCount(LPCWSTR lpInputPath)
{
	WCHAR szFullPattern[MAX_PATH];
	WIN32_FIND_DATAW FindFileData;
	HANDLE hFindFile;
	PathCombineW(szFullPattern, lpInputPath, L"*");
	hFindFile = FindFirstFileW(szFullPattern, &FindFileData);
	if (hFindFile != INVALID_HANDLE_VALUE)
	{
		do
		{
			if (FindFileData.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY)
			{
				if (lstrcmpW(FindFileData.cFileName, L"..") != 0 &&
					lstrcmpW(FindFileData.cFileName, L".") != 0)
				{
					PathCombineW(szFullPattern, lpInputPath, FindFileData.cFileName);
					TargetFileCount(szFullPattern);
				}
			}
			else
			{
				PathCombineW(szFullPattern, lpInputPath, FindFileData.cFileName);
				if (IsTargetFile(szFullPattern))
				{
					SendMessageW(hList, LB_ADDSTRING, 0, (LPARAM)szFullPattern);
				}
			}
		} while (FindNextFileW(hFindFile, &FindFileData));
		FindClose(hFindFile);
	}
}

struct VS_VERSIONINFO
{
	WORD                wLength;
	WORD                wValueLength;
	WORD                wType;
	WCHAR               szKey[1];
	WORD                wPadding1[1];
	VS_FIXEDFILEINFO    Value;
	WORD                wPadding2[1];
	WORD                wChildren[1];
};

struct
{
	WORD wLanguage;
	WORD wCodePage;
} *lpTranslate;

BOOL UpdateVersionInfo(LPCWSTR lpszFilePath, LPCWSTR lpszVersion)
{
	BOOL bResult = FALSE;
	DWORD dwHandle;
	const DWORD dwSize = GetFileVersionInfoSizeW(lpszFilePath, &dwHandle);
	if (dwSize > 0)
	{
		LPBYTE lpBuffer = (LPBYTE)GlobalAlloc(0, dwSize);
		if (lpBuffer)
		{
			if (GetFileVersionInfoW(lpszFilePath, 0, dwSize, lpBuffer) != FALSE)
			{
				VS_VERSIONINFO*pVerInfo = (VS_VERSIONINFO*)lpBuffer;
				LPBYTE pOffsetBytes = (LPBYTE)&pVerInfo->szKey[lstrlenW(pVerInfo->szKey) + 1];
#define roundoffs(a,b,r) (((LPBYTE) (b) - (LPBYTE) (a) + ((r) - 1)) & ~((r) - 1))
#define roundpos(a,b,r) (((LPBYTE) (a)) + roundoffs(a,b,r))
				VS_FIXEDFILEINFO*pFixedInfo = (VS_FIXEDFILEINFO*)roundpos(pVerInfo, pOffsetBytes, 4);
#undef roundpos
#undef roundoffs
				int nMajorVersion = 0, nMinorVersion = 0, nBuildNumber = 0, nRevision = 0;
				swscanf_s(lpszVersion, L"%d.%d.%d.%d", &nMajorVersion, &nMinorVersion, &nBuildNumber, &nRevision);
				pFixedInfo->dwFileVersionMS = MAKELONG(nMinorVersion, nMajorVersion);
				pFixedInfo->dwFileVersionLS = MAKELONG(nRevision, nBuildNumber);
				const HANDLE hResource = BeginUpdateResourceW(lpszFilePath, FALSE);
				if (hResource)
				{
					UINT uLen;
					if (VerQueryValueW(lpBuffer, L"\\VarFileInfo\\Translation", (LPVOID *)&lpTranslate, &uLen) != FALSE)
					{
						if (UpdateResourceW(hResource, MAKEINTRESOURCEW(16), MAKEINTRESOURCEW(VS_VERSION_INFO), lpTranslate->wLanguage, lpBuffer, dwSize) != FALSE)
						{
							if (EndUpdateResourceW(hResource, FALSE))
							{
								bResult = TRUE;
							}
						}
					}
				}
			}
			GlobalFree(lpBuffer);
		}
	}
	return bResult;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit, hProgress;
	switch (msg)
	{
	case WM_CREATE:
		InitCommonControls();
		hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"),
			TEXT("1.0.0.1"),
			WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL,
			10, 10, 1024, 32, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hEdit, EM_LIMITTEXT, 0, 0);
		hProgress = CreateWindow(TEXT("msctls_progress32"), 0, WS_VISIBLE | WS_CHILD,
			0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hList = CreateWindow(TEXT("LISTBOX"), 0, WS_CHILD | WS_VISIBLE | WS_VSCROLL |
			WS_HSCROLL | LBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, hWnd, 0,
			((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_SIZE:
		MoveWindow(hEdit, 10, 10, LOWORD(lParam) - 20, 32, 1);
		MoveWindow(hProgress, 10, 50, LOWORD(lParam) - 20, 32, 1);
		MoveWindow(hList, 10, 90, LOWORD(lParam) - 20, HIWORD(lParam) - 100, 1);
		break;
	case WM_DROPFILES:
	{
		WCHAR szFilePath[MAX_PATH];
		const UINT iFileNum = DragQueryFileW((HDROP)wParam, -1, NULL, 0);
		for (UINT i = 0; i < iFileNum; ++i)
		{
			DragQueryFileW((HDROP)wParam, i, szFilePath, MAX_PATH);
			if (PathIsDirectoryW(szFilePath))
			{
				TargetFileCount(szFilePath);
			}
			else if (IsTargetFile(szFilePath))
			{
				SendMessageW(hList, LB_ADDSTRING, 0, (LPARAM)szFilePath);
			}
		}
		DragFinish((HDROP)wParam);
		SendMessageW(hProgress, PBM_SETRANGE32, 0, SendMessage(hList, LB_GETCOUNT, 0, 0));
		SendMessageW(hProgress, PBM_SETSTEP, 1, 0);
		SendMessageW(hProgress, PBM_SETPOS, 0, 0);
		DWORD dwSize = GetWindowTextLengthW(hEdit);
		LPWSTR lpszVersion = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR)*(dwSize + 1));
		GetWindowTextW(hEdit, lpszVersion, dwSize + 1);
		while (SendMessageW(hList, LB_GETCOUNT, 0, 0))
		{
			SendMessageW(hList, LB_GETTEXT, 0, (LPARAM)szFilePath);
			UpdateVersionInfo(szFilePath, lpszVersion);
			SendMessageW(hList, LB_DELETESTRING, 0, 0);
			SendMessageW(hProgress, PBM_STEPIT, 0, 0);
		}
		GlobalFree(lpszVersion);
		MessageBoxW(hWnd, L"更新が完了しました。", L"確認", 0);
	}
	break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("ドロップされたファイルまたはフォルダに含まれるEXE、DLLのファイルバージョンを書き換える"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}