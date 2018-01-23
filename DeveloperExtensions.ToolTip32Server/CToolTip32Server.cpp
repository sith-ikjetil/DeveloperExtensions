// CToolTip32Server.cpp : Implementation of CToolTip32Server

#include "stdafx.h"
#include "CToolTip32Server.h"
#include <exdisp.h>
#include "oleauto.h"
#include <windows.h>
#include <comadmin.h>
#include <tlhelp32.h>
#include <commctrl.h>
#include <vector>
using namespace std;

#include "../Include/itsoftware.h"
#include "../Include/itsoftware-com.h"

wchar_t *BigNumToString(LONG lNum, wchar_t *wcsBuffer);
wchar_t *BigNumToStringF(float fNum, wchar_t *wcsBuffer);

//STDAPI DllRegisterServerType
typedef HRESULT(WINAPI* DllRegisterServerType)(void);

//STDAPI DllUnregisterServer(void)
typedef HRESULT(WINAPI* DllUnregisterServerType)(void);

// CToolTip32Server
STDMETHODIMP CToolTip32Server::GetToolTip(/*[in]*/ BSTR filename, /*[out, retval]*/ BSTR* tooltip)
{
	CComBSTR bstrFileName(filename);

	wchar_t buffer[1024];
	wcscpy_s(buffer, 1024, bstrFileName.operator LPWSTR());
	_wcslwr_s(buffer, wcslen(buffer) + 1);

	size_t length = wcslen(buffer);
	
	bool bIsTLB = false;
	if (wcscmp(&buffer[length - 3], L"tlb") == 0)
	{
		bIsTLB = true;
	}

	CComBSTR bstrInfo("[General]\n  Type: [");
	SHFILEINFOW sfi;

	SHGetFileInfo(bstrFileName.operator LPWSTR(), 0, &sfi, sizeof(sfi), SHGFI_TYPENAME);
	bstrInfo += sfi.szTypeName;
	bstrInfo.Append("]\n  Size: [");

	// get info		
	HANDLE hFile = CreateFile(bstrFileName.operator LPWSTR(), 0, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile != NULL) {
		DWORD dwFileSizeLow;
		DWORD dwFileSizeHigh;
		dwFileSizeLow = GetFileSize(hFile, &dwFileSizeHigh);
		CloseHandle(hFile);

		wchar_t wcsFileSizeKB[100];
		float fFileSizeLow = static_cast<float>(dwFileSizeLow);
		fFileSizeLow /= 1024;
		BigNumToStringF(fFileSizeLow, wcsFileSizeKB);

		wchar_t wcsFileSizeLow[100];
		BigNumToString(dwFileSizeLow, wcsFileSizeLow);

		bstrInfo += wcsFileSizeKB;
		bstrInfo.Append(" KB (");
		bstrInfo += wcsFileSizeLow;
		bstrInfo.Append(" bytes)]\n");
	}
	else {
		bstrInfo.Append("[0 KB (0 bytes)]\n  ");
	}

	bstrInfo.Append("  Machine: [32-bit]\n  ");

	// COM/COM+:
	bstrInfo.Append("\n[COM/COM+]\n  ");
	bstrInfo.Append(" COM Server: ");

	if (IsDllCOMServer(bstrFileName.operator LPWSTR())) {
		bstrInfo.Append("[Yes]\n  ");

		/*vector<CComBSTR> vProgIDs(ContainsProgID(m_wcsToolTipFilename));
		if ( vProgIDs.size() > 0 ) {

		bstrInfo.Append(" ProgIDs: ");
		for ( long l = 0; l < vProgIDs.size(); l++ ) {
		bstrInfo.Append(" [");
		bstrInfo.AppendBSTR(vProgIDs[l]);
		bstrInfo.Append("]\n");
		if ( l+1 < vProgIDs.size() ) {
		bstrInfo.Append("       ");
		}
		}
		}
		*/
		bstrInfo.Append(" Configured: ");
		CComBSTR bstrName;
		bstrName = IsRegistredInCOMPlus(bstrFileName.operator LPWSTR());
		if (bstrName.Length() > 0) {
			bstrInfo.Append("[Yes] in [");
			bstrInfo.AppendBSTR(bstrName);
			bstrInfo.Append("]\n  ");
		}
		else {
			bstrInfo.Append("[No]\n  ");
		}

		// new typelib info
		bstrInfo.Append(" Classes: ");
		CComPtr<ITypeLib> pITypeLib;
		HRESULT hr = ::LoadTypeLibEx(bstrFileName.operator LPWSTR(), REGKIND_NONE, &pITypeLib);
		if (FAILED(hr)) {
			bstrInfo.Append(" <err>  ");
		}
		else {
			int iCount = 0;
			if (pITypeLib != NULL) {
				UINT uCount = pITypeLib->GetTypeInfoCount();
				for (UINT i = 0; i < uCount; i++) {
					CComPtr<ITypeInfo> pITypeInfo;
					hr = pITypeLib->GetTypeInfo(i, &pITypeInfo);
					if (SUCCEEDED(hr) && pITypeInfo != NULL) {
						TYPEATTR *pTypeAttr;
						hr = pITypeInfo->GetTypeAttr(&pTypeAttr);
						if (SUCCEEDED(hr) && pTypeAttr != NULL) {
							if (pTypeAttr->typekind == TKIND_COCLASS) {
								CComBSTR bstrName;
								hr = pITypeLib->GetDocumentation(i, &bstrName, NULL, NULL, NULL);
								if (SUCCEEDED(hr)) {
									if (iCount > 0) {
										bstrInfo.Append(", ");
									}
									bstrInfo.Append("[");
									bstrInfo.AppendBSTR(bstrName);
									bstrInfo.Append("]");
									iCount++;
								}
							}
							pITypeInfo->ReleaseTypeAttr(pTypeAttr);
						}
					}// if
				}// for
			}// if ( pITypeLib != NULL ) {
		}
		bstrInfo.Append("\n");
	}
	else {
		bstrInfo.Append("[No]\n  ");

		if (bIsTLB)
		{
			// new typelib info
			bstrInfo.Append(" Classes: ");
			CComPtr<ITypeLib> pITypeLib;
			HRESULT hr = ::LoadTypeLibEx(bstrFileName.operator LPWSTR(), REGKIND_NONE, &pITypeLib);
			if (FAILED(hr)) {
				bstrInfo.Append(" <err>  ");
			}
			else {
				int iCount = 0;
				if (pITypeLib != NULL) {
					UINT uCount = pITypeLib->GetTypeInfoCount();
					for (UINT i = 0; i < uCount; i++) {
						CComPtr<ITypeInfo> pITypeInfo;
						hr = pITypeLib->GetTypeInfo(i, &pITypeInfo);
						if (SUCCEEDED(hr) && pITypeInfo != NULL) {
							TYPEATTR *pTypeAttr;
							hr = pITypeInfo->GetTypeAttr(&pTypeAttr);
							if (SUCCEEDED(hr) && pTypeAttr != NULL) {
								if (pTypeAttr->typekind == TKIND_COCLASS) {
									CComBSTR bstrName;
									hr = pITypeLib->GetDocumentation(i, &bstrName, NULL, NULL, NULL);
									if (SUCCEEDED(hr)) {
										if (iCount > 0) {
											bstrInfo.Append(", ");
										}
										bstrInfo.Append("[");
										bstrInfo.AppendBSTR(bstrName);
										bstrInfo.Append("]");
										iCount++;
									}
								}
								pITypeInfo->ReleaseTypeAttr(pTypeAttr);
							}
						}// if
					}// for
				}// if ( pITypeLib != NULL ) {
			}
			bstrInfo.Append("\n");
		}
	}

	if (!bIsTLB) {
		bstrInfo.Append("\n[Advanced]\n  ");
		bstrInfo.Append("  Base address: ");

		DWORD dwBaseAddress = GetBaseAddressOf(bstrFileName.operator LPWSTR());
		wchar_t wcsBaseAddress[40];
		wsprintf(wcsBaseAddress, L"[%#08x]", dwBaseAddress);
		bstrInfo += wcsBaseAddress;
	}
	bstrInfo.Append("\n\n    Copyright (c) 2001-2017 by Kjetil Kristoffer Solberg.\n    All rights reserved.");

	// set tooltip
	/**ppwszTip = (LPWSTR)CoTaskMemAlloc((bstrInfo.Length() + 1) * sizeof(WCHAR));
	if (*ppwszTip)
		wcscpy_s((*(ppwszTip)), bstrInfo.Length() * sizeof(WCHAR), bstrInfo.operator LPWSTR());
		*/
	*tooltip = bstrInfo.Detach();

	return S_OK;
}

// This function accepts a number and converts it to a
// string, inserting commas where appropriate.
wchar_t *BigNumToString(LONG lNum, wchar_t *wcsBuffer) {

	wchar_t wcsNum[100];
	wsprintf(wcsNum, TEXT("%d"), lNum);
	NUMBERFMT nf;
	nf.NumDigits = 0;
	nf.LeadingZero = FALSE;
	nf.Grouping = 3;
	nf.lpDecimalSep = L".";
	nf.lpThousandSep = L",";
	nf.NegativeOrder = 0;
	GetNumberFormat(LOCALE_USER_DEFAULT, 0, wcsNum, &nf, wcsBuffer, 100);
	return(wcsBuffer);
}

wchar_t *BigNumToStringF(float fNum, wchar_t *wcsBuffer) {

	wchar_t wcsNum[100];
	_snwprintf_s(wcsNum, 100, L"%f", fNum);
	NUMBERFMT nf;
	nf.NumDigits = 1;
	nf.LeadingZero = FALSE;
	nf.Grouping = 3;
	nf.lpDecimalSep = L".";
	nf.lpThousandSep = L",";
	nf.NegativeOrder = 0;
	GetNumberFormat(LOCALE_USER_DEFAULT, 0, wcsNum, &nf, wcsBuffer, 100);
	return(wcsBuffer);
}

DWORD CToolTip32Server::GetBaseAddressOf(wchar_t *pwcs)
{
	HMODULE hModule = ::LoadLibraryEx(pwcs,NULL, DONT_RESOLVE_DLL_REFERENCES);
	if (hModule == NULL)
	{		
		return 0;
	}

	PVOID pvModulePreferredBaseAddr(NULL);
	IMAGE_DOS_HEADER idh = { 0 };
	IMAGE_NT_HEADERS inth = { 0 };

	HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetCurrentProcessId());
	if (hSnapshot == INVALID_HANDLE_VALUE)
		return 0;

	MODULEENTRY32 me = { 0 };
	me.dwSize = sizeof(me);
	BOOL bRet = Module32First(hSnapshot, &me);
	if (me.hModule == hModule) {
		Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr, &idh, sizeof(idh), NULL);
		if (idh.e_magic == IMAGE_DOS_SIGNATURE) {
			Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr + idh.e_lfanew, &inth, sizeof(inth), NULL);
			if (inth.Signature == IMAGE_NT_SIGNATURE) {
				pvModulePreferredBaseAddr = (PVOID)inth.OptionalHeader.ImageBase;
			}
		}
	}
	else {
		while (Module32Next(hSnapshot, &me)) {
			if (me.hModule == hModule) {
				Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr, &idh, sizeof(idh), NULL);
				if (idh.e_magic == IMAGE_DOS_SIGNATURE) {
					Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr + idh.e_lfanew, &inth, sizeof(inth), NULL);
					if (inth.Signature == IMAGE_NT_SIGNATURE) {
						pvModulePreferredBaseAddr = (PVOID)inth.OptionalHeader.ImageBase;
					}
				}
			}
		}
	}


	FreeLibrary(hModule);
	return (DWORD)pvModulePreferredBaseAddr;
}

CComBSTR CToolTip32Server::IsRegistredInCOMPlus(wchar_t *pwcs)
{
	CComBSTR bstrInprocServer;
	bstrInprocServer += pwcs;
	bstrInprocServer.ToLower();

	CComPtr<ICOMAdminCatalog> pAdminCatalog;
	HRESULT hr = pAdminCatalog.CoCreateInstance(CComBSTR("COMAdmin.COMAdminCatalog.1"));//CLSID_COMAdminCatalog );
	if (FAILED(hr)) {
		return NULL;
	}

	CComPtr<IDispatch> pDispatch;
	hr = pAdminCatalog->GetCollection(CComBSTR("Applications"), &pDispatch);
	if (FAILED(hr)) {
		return NULL;
	}

	CComQIPtr<ICatalogCollection> pCollection;
	pCollection = pDispatch;

	pCollection->Populate();

	long lCount(0);
	hr = pCollection->get_Count(&lCount);

	long lComponents(0);
	for (long l = 0; l < lCount; l++) {
		CComPtr<IDispatch> pDispatch;
		hr = pCollection->get_Item(l, &pDispatch);
		if (FAILED(hr)) {
			return NULL;
		}

		CComQIPtr<ICatalogObject> pObject;
		pObject = pDispatch;
		pDispatch = NULL;

		CComVariant vtName;
		hr = pObject->get_Value(CComBSTR("Name"), &vtName);
		if (FAILED(hr))
			return NULL;

		CComVariant vtKey;
		hr = pObject->get_Key(&vtKey);
		if (FAILED(hr))
			return NULL;

		hr = pCollection->GetCollection(CComBSTR("Components"), vtKey, &pDispatch);
		CComQIPtr<ICatalogCollection> pCollectionC;
		pCollectionC = pDispatch;
		pDispatch = NULL;

		pCollectionC->Populate();

		long lCount(0);
		hr = pCollectionC->get_Count(&lCount);
		for (int i = 0; i < lCount; i++) {
			CComPtr<IDispatch> pDispatch;
			hr = pCollectionC->get_Item(i, &pDispatch);
			if (FAILED(hr)) {
				return NULL;
			}

			CComQIPtr<ICatalogObject> pObject;
			pObject = pDispatch;
			pDispatch = NULL;

			CComVariant vtValue;
			CComBSTR bstrDLL("DLL");
			hr = pObject->get_Value(bstrDLL, &vtValue);
			if (FAILED(hr)) {
				return NULL;
			}

			CComBSTR bstrFile;
			bstrFile.AppendBSTR(vtValue.bstrVal);
			bstrFile.ToLower();

			if (bstrInprocServer == bstrFile) {
				CComBSTR bstr;
				bstr.AppendBSTR(vtName.bstrVal);
				return bstr;
			}
		}
	}// for

	return NULL;
}

vector<CComBSTR> CToolTip32Server::ContainsProgID(wchar_t *pwcs)
{
	vector<CComBSTR> vProgIDs;
	CComPtr<ICOMAdminCatalog> pAdminCatalog;
	HRESULT hr = pAdminCatalog.CoCreateInstance(CComBSTR("COMAdmin.COMAdminCatalog.1"));//CLSID_COMAdminCatalog );
	if (FAILED(hr)) {
		return vProgIDs;
	}

	CComPtr<IDispatch> pDispatch;
	hr = pAdminCatalog->GetCollection(CComBSTR("InprocServers"), &pDispatch);
	if (FAILED(hr)) {
		return vProgIDs;
	}

	CComQIPtr<ICatalogCollection> pCollection;
	pCollection = pDispatch;

	pCollection->Populate();

	long lCount(0);
	hr = pCollection->get_Count(&lCount);

	long lComponents(0);
	CComBSTR bstrSrc;
	bstrSrc += pwcs;
	bstrSrc.ToLower();

	for (long l = 0; l < lCount; l++) {
		CComPtr<IDispatch> pDispatch;
		hr = pCollection->get_Item(l, &pDispatch);
		if (FAILED(hr)) {
			return vProgIDs;
		}

		CComQIPtr<ICatalogObject> pObject;
		pObject = pDispatch;
		pDispatch = NULL;

		CComVariant vtValue;
		CComBSTR bstrDLL("InprocServer32");
		hr = pObject->get_Value(bstrDLL, &vtValue);
		if (FAILED(hr))
			return vProgIDs;

		CComBSTR bstrValue;
		bstrValue.AppendBSTR(vtValue.bstrVal);
		if (bstrValue.Length()) {
			bstrValue.ToLower();
			if (bstrSrc == bstrValue) {
				CComVariant vtProgID;
				CComBSTR bstrProgID("ProgID");
				hr = pObject->get_Value(bstrProgID, &vtProgID);
				CComBSTR bstrRetVal;
				bstrRetVal.AppendBSTR(vtProgID.bstrVal);
				vProgIDs.push_back(bstrRetVal);
			}
		}
	}// for

	return vProgIDs;
}

bool CToolTip32Server::IsDllCOMServer(wchar_t *pwcs)
{
	HMODULE hModule = ::LoadLibraryEx(pwcs,NULL, DONT_RESOLVE_DLL_REFERENCES );
	if (hModule != NULL) {
		DllRegisterServerType fnRegisterServer = (DllRegisterServerType) ::GetProcAddress(hModule, "DllRegisterServer");
		if (fnRegisterServer != NULL) {
			FreeLibrary(hModule);
			return true;
		}
		FreeLibrary(hModule);
	}

	return false;
}