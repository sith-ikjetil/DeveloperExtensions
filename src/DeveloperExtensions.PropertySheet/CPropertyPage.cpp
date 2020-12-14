// CPropertyPage.cpp : Implementation of CCPropertyPage

#include "stdafx.h"
#include "imagehlp.h"
#include "CPropertyPage.h"
#include "../Include/itsoftware.h"
#include "../Include/itsoftware-com.h"
#include "../DeveloperExtensions.ToolTip/DeveloperExtensionsToolTip_i.h"
#include "../DeveloperExtensions.ToolTip32Server/DeveloperExtensionsToolTip32Server_i.h"
#include "../DeveloperExtensions.ToolTip64Server/DeveloperExtensionsToolTip64Server_i.h"
#include <string>

using namespace std;

extern HINSTANCE g_hInstanceDll;

UINT CALLBACK PropPageProc(HWND hWnd, UINT uMsg, LPPROPSHEETPAGE ppsp);
INT_PTR CALLBACK PropPageDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
CComBSTR g_bstrFileName;
CComBSTR g_bstrToolTip;

// IShellExtInit
STDMETHODIMP CCPropertyPage::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pDataObj, HKEY hkeyProgID)
{	
	if (pDataObj == nullptr) {
		return E_FAIL;
	}

	FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stg = { TYMED_HGLOBAL };
	HDROP     hDrop;

	// Look for CF_HDROP data in the data object.
	if (FAILED(pDataObj->GetData(&fmt, &stg)))
	{
		// Nope! Return an "invalid argument" error back to Explorer.
		return E_INVALIDARG;
	}

	// Get a pointer to the actual data.
	hDrop = (HDROP)GlobalLock(stg.hGlobal);

	// Make sure it worked.
	if (NULL == hDrop)
	{
		return E_INVALIDARG;
	}

	// Sanity check - make sure there is at least one filename.
	UINT uNumFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);

	if (0 == uNumFiles)
	{
		GlobalUnlock(stg.hGlobal);
		ReleaseStgMedium(&stg);
		return E_INVALIDARG;
	}

	HRESULT hr = S_OK;

	wchar_t wszFileName[MAX_PATH];
	// Get the name of the first file and store it in our member variable m_szFile.
	if (0 == DragQueryFile(hDrop, 0, wszFileName, MAX_PATH))
	{
		hr = E_INVALIDARG;
	}

	m_bstrFileName = CComBSTR(wszFileName);
	g_bstrFileName = CComBSTR(wszFileName);

	this->Load(wszFileName, 0);
	this->GetInfoTip(0, &g_bstrToolTip);

	GlobalUnlock(stg.hGlobal);
	ReleaseStgMedium(&stg);

	return hr;
}

// CCPropertyPage
STDMETHODIMP CCPropertyPage::AddPages(LPFNADDPROPSHEETPAGE pfnAddPage, LPARAM lParam)
{		
	PROPSHEETPAGE psp = {};
	psp.dwSize = sizeof(psp);	
	psp.hInstance = g_hInstanceDll;
	psp.pfnCallback = PropPageProc;
	psp.pfnDlgProc = PropPageDialogProc;
	psp.dwFlags = PSP_USECALLBACK | PSP_USEHEADERTITLE;
	psp.pszTemplate = MAKEINTRESOURCE(IDD_PROPPAGE);
	
	HPROPSHEETPAGE hPropSheet = CreatePropertySheetPage(&psp);			

	if (hPropSheet)
	{	
		if (pfnAddPage(hPropSheet, lParam))
		{			
			this->AddRef();
			return S_OK;
		}
		else
		{			
			DestroyPropertySheetPage(hPropSheet);
		}
	}
	else
	{
		return E_OUTOFMEMORY;
	}
	
	return E_FAIL;	
}

STDMETHODIMP CCPropertyPage::ReplacePage(UINT uPageID, LPFNADDPROPSHEETPAGE pfnReplacePage, LPARAM lParam)
{
	return S_OK;
}

// PropPageProc
UINT CALLBACK PropPageProc(HWND hwndDlg,	UINT uMsg,	LPPROPSHEETPAGE ppsp)
{
	switch (uMsg)
	{
	case PSPCB_CREATE:		
	{		
		return TRUE;
	}
	case PSPCB_RELEASE:
	{
		// Free any data allocated.
	}	
	default:
		break;
	}

	return FALSE;
}// UINT PropPageProc


 // AboutDialogProc
INT_PTR CALLBACK PropPageDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
	{
		SetDlgItemTextW(hwndDlg, IDC_STATIC, g_bstrToolTip);
		return TRUE;
	}

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:
			EndDialog(hwndDlg, 0);
			return TRUE;
			break;
		default:
			break;
		}
		break;

	default:
		break;
	}

	return FALSE;
}// BOOL AboutDialogProc

STDMETHODIMP CCPropertyPage::GetInfoTip(DWORD dwFlags, LPWSTR *ppwszTip)
{
	*ppwszTip = NULL;
	CComBSTR bstrToolTip;

	if (this->m_fileType == EnumFileType::Dll32)
	{
		CComPtr<IDEToolTip32Server> pIToolTip32Server;
		HRESULT hr = pIToolTip32Server.CoCreateInstance(L"DeveloperExtensions.ToolTip32Server");
		if (FAILED(hr)) {
			return E_FAIL;
		}
		hr = pIToolTip32Server->GetToolTip(this->m_bstrFileName, &bstrToolTip);
		if (FAILED(hr)) {
			return E_FAIL;
		}
	}
	else if (this->m_fileType == EnumFileType::Dll64) {
		CComPtr<IDEToolTip64Server> pIToolTip64Server;
		HRESULT hr = pIToolTip64Server.CoCreateInstance(L"DeveloperExtensions.ToolTip64Server");
		if (FAILED(hr)) {
			return E_FAIL;
		}
		hr = pIToolTip64Server->GetToolTip(this->m_bstrFileName, &bstrToolTip);
		if (FAILED(hr)) {
			return E_FAIL;
		}
	}
	else if (this->m_fileType == EnumFileType::Tlb32) {
		CComPtr<IDEToolTip32Server> pIToolTip32Server;
		HRESULT hr = pIToolTip32Server.CoCreateInstance(L"DeveloperExtensions.ToolTip32Server");
		if (FAILED(hr)) {
			return E_FAIL;
		}
		hr = pIToolTip32Server->GetToolTip(this->m_bstrFileName, &bstrToolTip);
		if (FAILED(hr)) {
			return E_FAIL;
		}
	}
	else if (this->m_fileType == EnumFileType::Tlb64) {
		CComPtr<IDEToolTip64Server> pIToolTip64Server;
		HRESULT hr = pIToolTip64Server.CoCreateInstance(L"DeveloperExtensions.ToolTip64Server");
		if (FAILED(hr)) {
			return E_FAIL;
		}
		hr = pIToolTip64Server->GetToolTip(this->m_bstrFileName, &bstrToolTip);
		if (FAILED(hr)) {
			return E_FAIL;
		}
	}
	else if (this->m_fileType == EnumFileType::Unknwn) {
		return E_FAIL;
	}

	// set tooltip
	*ppwszTip = (LPWSTR)CoTaskMemAlloc((bstrToolTip.Length() + 1) * sizeof(WCHAR));
	if (*ppwszTip)
		wcscpy_s((*(ppwszTip)), bstrToolTip.Length() + 1, bstrToolTip.operator LPWSTR());

	return S_OK;
}


STDMETHODIMP CCPropertyPage::Load(LPCOLESTR wszFile, DWORD dwMode)
{	
	wchar_t buffer[1024];
	wcscpy_s(buffer, 1024, this->m_bstrFileName);
	_wcslwr_s(buffer, wcslen(buffer) + 1);

	size_t length = wcslen(buffer);

	if (wcscmp(&buffer[length - 3], L"dll") == 0)
	{
		CComBSTR bstrFullPath = this->m_bstrFileName.Copy();

		CComBSTR bstrPath;
		this->ExtractPath(bstrFullPath, &bstrPath);

		CComBSTR bstrFileName;
		this->ExtractFilename(bstrFullPath, &bstrFileName);

		wstring wPath(bstrPath.operator LPWSTR());
		wstring wFileName(bstrFileName.operator LPWSTR());

		string sPath;
		sPath.assign(wPath.begin(), wPath.end());

		string sFileName;
		sFileName.assign(wFileName.begin(), wFileName.end());

		PLOADED_IMAGE pLoadedImage = ::ImageLoad(sFileName.c_str(), sPath.c_str());
		if (pLoadedImage == NULL) {
			this->m_fileType = EnumFileType::Unknwn;
			return S_OK;
		}

		if (pLoadedImage->FileHeader->FileHeader.Machine == IMAGE_FILE_MACHINE_I386)
		{
			this->m_fileType = EnumFileType::Dll32;
		}
		else if (pLoadedImage->FileHeader->FileHeader.Machine == IMAGE_FILE_MACHINE_AMD64)
		{
			this->m_fileType = EnumFileType::Dll64;
		}
		else
		{
			this->m_fileType = EnumFileType::Unknwn;
		}

		ImageUnload(pLoadedImage);
	}
	else if (wcscmp(&buffer[length - 3], L"tlb") == 0)
	{
		CComBSTR bstrFullPath = this->m_bstrFileName.Copy();

		CComPtr<ITypeLib> pITypeLib;
		HRESULT hr = ::LoadTypeLibEx(bstrFullPath.operator LPWSTR(), REGKIND_NONE, &pITypeLib);
		if (FAILED(hr)) {
			this->m_fileType = EnumFileType::Unknwn;
			return S_OK;
		}
		TLIBATTR *patt;
		hr = pITypeLib->GetLibAttr(&patt);
		if (FAILED(hr)) {
			this->m_fileType = EnumFileType::Unknwn;
			return S_OK;
		}
		switch (patt->syskind)
		{
		case SYSKIND::SYS_WIN64:
			this->m_fileType = EnumFileType::Tlb64;
			break;

		case SYSKIND::SYS_WIN32:
			this->m_fileType = EnumFileType::Tlb32;
			break;

		case SYSKIND::SYS_WIN16:
			this->m_fileType = EnumFileType::Unknwn;
			break;
		}
		pITypeLib->ReleaseTLibAttr(patt);
	}
	else {
		this->m_fileType = EnumFileType::Unknwn;
	}

	return S_OK;
};


STDMETHODIMP CCPropertyPage::ExtractPath( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* path)
{
	if (SysStringLen(full_path) == 0 || path == NULL) {
		return E_INVALIDARG;//Error(CComBSTR(L"ExtractPath Failed. Invalid Arguments."), IID_INCFPath, E_INVALIDARG);
	}

	long lLength = SysStringLen(full_path);
	long lLastHit(-1);
	for (long l = 0; l < lLength; l++) {
		if (full_path[l] == L'\\' || full_path[l] == L'/') {
			lLastHit = l;
		}
	}

	CComBSTR bstrRetval(L"");
	if (lLastHit != -1) {
		for (long l = 0; l <= lLastHit; l++) {
			bstrRetval.Append((wchar_t)full_path[l]);
		}
		*path = bstrRetval.Detach();
	}
	else {
		*path = bstrRetval.Detach();
	}

	return S_OK;
}

STDMETHODIMP CCPropertyPage::ExtractFilename( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* filename)
{
	if (SysStringLen(full_path) == 0 || filename == NULL) {
		return E_INVALIDARG;//Error(CComBSTR(L"ExtractExtension Failed. Invalid Arguments."), IID_INCFPath, E_INVALIDARG);
	}

	long lLength = SysStringLen(full_path);
	long lLastHit(-1);
	for (long l = 0; l < lLength; l++) {
		if (full_path[l] == L'\\' || full_path[l] == L'/') {
			lLastHit = l;
		}
	}

	CComBSTR bstrRetval(L"");
	if (lLastHit != -1) {
		for (long l = lLastHit + 1; l < lLength; l++) {
			bstrRetval.Append((wchar_t)full_path[l]);
		}
		*filename = bstrRetval.Detach();
	}
	else {
		bstrRetval.AppendBSTR(full_path);
		*filename = bstrRetval.Detach();
	}

	return S_OK;
}