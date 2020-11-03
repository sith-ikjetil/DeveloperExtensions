// ToolTip.cpp : Implementation of CToolTip
#include "stdafx.h"
#include "CToolTip.h"
#include "imagehlp.h"
#include <string>
#include "DeveloperExtensionsToolTip_i.h"
#include "../DeveloperExtensions.ToolTip32Server/DeveloperExtensionsToolTip32Server_i.h"
#include "../DeveloperExtensions.ToolTip64Server/DeveloperExtensionsToolTip64Server_i.h"

extern HINSTANCE g_hInstance;

using namespace std;

// IQueryInfo
STDMETHODIMP CToolTip::GetInfoFlags(DWORD *pdwFlags)
{
	*pdwFlags = 0;
	return E_NOTIMPL;
}

STDMETHODIMP CToolTip::GetInfoTip(DWORD dwFlags, LPWSTR *ppwszTip)
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
		hr = pIToolTip32Server->GetToolTip(CComBSTR(this->m_wcsToolTipFilename), &bstrToolTip);
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
		hr = pIToolTip64Server->GetToolTip(CComBSTR(this->m_wcsToolTipFilename), &bstrToolTip);
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
		hr = pIToolTip32Server->GetToolTip(CComBSTR(this->m_wcsToolTipFilename), &bstrToolTip);
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
		hr = pIToolTip64Server->GetToolTip(CComBSTR(this->m_wcsToolTipFilename), &bstrToolTip);
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


STDMETHODIMP CToolTip::Load(LPCOLESTR wszFile, DWORD dwMode)
{
	wcscpy_s(m_wcsToolTipFilename, 1024, wszFile);
	
	wchar_t buffer[1024];
	wcscpy_s(buffer, 1024, this->m_wcsToolTipFilename);
	_wcslwr_s(buffer, wcslen(buffer) + 1);

	size_t length = wcslen(buffer);

	if (wcscmp(&buffer[length - 3], L"dll") == 0)
	{
		CComBSTR bstrFullPath(this->m_wcsToolTipFilename);

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
		CComBSTR bstrFullPath(this->m_wcsToolTipFilename);

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

STDMETHODIMP CToolTip::GetClassID(LPCLSID pCLSID)
{	
	*pCLSID = CLSID_ToolTip;//memcpy(pCLSID, &CLSID_ToolTip, sizeof(CLSID));
	return S_OK;
}

STDMETHODIMP CToolTip::IsDirty(VOID) { return E_NOTIMPL; }

STDMETHODIMP CToolTip::Save(LPCOLESTR, BOOL) { return E_NOTIMPL; }

STDMETHODIMP CToolTip::SaveCompleted(LPCOLESTR) { return E_NOTIMPL; }

STDMETHODIMP CToolTip::GetCurFile(LPOLESTR FAR*) { return E_NOTIMPL; }


STDMETHODIMP CToolTip::ExtractPath( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* path)
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

STDMETHODIMP CToolTip::ExtractFilename( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* filename)
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