#pragma once
//
// #include
//
#include <windows.h>
#include <atlbase.h>
#include <string>
#include <lmerr.h>

//
// #define
//
#define ERRMSGBUFFERSIZE 256

//
// namespace
//
namespace ItSoftware
{	
	//
	// namespace
	//
	namespace COM
	{
		//
		// using namespace
		//
		using namespace ATL;

		// 
		// ComException class.
		//
		class ComException
		{
		private:
			HRESULT  m_hr;
			CComBSTR m_message;
			CComBSTR m_errorInfoMessage;
			CComBSTR m_systemMessage;

			CComBSTR GetSystemMessage(HRESULT hr)
			{
				wchar_t wszMsg[1024];
				::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, wszMsg, 1024, NULL);
				CComBSTR bstr(wszMsg);
				return bstr;
			}

			CComBSTR GetHResultName(HRESULT hr)
			{
				if (hr == S_OK)
				{
					return CComBSTR(L"S_OK");
				}
				else if (hr == S_FALSE)
				{
					return CComBSTR(L"S_FALSE");
				}
				else if (hr == E_FAIL)
				{
					return CComBSTR(L"E_FAIL");
				}
				else if (hr == E_OUTOFMEMORY)
				{
					return CComBSTR(L"E_OUTOFMEMORY");
				}
				else if (hr == E_HANDLE)
				{
					return CComBSTR(L"E_HANDLE");
				}
				else if (hr == E_INVALIDARG)
				{
					return CComBSTR(L"E_INVALIDARG");
				}
				else if (hr == CLASS_E_CLASSNOTAVAILABLE)
				{
					return CComBSTR(L"CLASS_E_CLASSNOTAVAILABLE");
				}
				else if (hr == E_NOINTERFACE)
				{
					return CComBSTR(L"E_NOINTERFACE");
				}
				return CComBSTR(L"?");
			}

			CComBSTR GetErrorInfo(HRESULT hr)
			{
				CComPtr<IErrorInfo> pIErrorInfo;
				if (SUCCEEDED(::GetErrorInfo(NULL, &pIErrorInfo)))
				{
					if (pIErrorInfo != nullptr)
					{
						CComBSTR bstr;
						pIErrorInfo->GetDescription(&bstr);
						return bstr;
					}
				}
				return CComBSTR(L"");
			}

		public:


			ComException(HRESULT hr)
				: m_hr(hr),
				m_message(CComBSTR(L"")),
				m_systemMessage(GetSystemMessage(hr)),
				m_errorInfoMessage(this->GetErrorInfo(hr))
			{
			}

			ComException(HRESULT hr, CComBSTR message)
				: m_hr(hr),
				m_message(message),
				m_systemMessage(GetSystemMessage(hr)),
				m_errorInfoMessage(this->GetErrorInfo(hr))
			{
			}

			CComBSTR GetMessage()
			{
				return this->m_message;
			}

			CComBSTR GetSystemMessage()
			{
				return this->m_systemMessage;
			}

			CComBSTR ToString()
			{
				std::wstring result;
				result.append(L"## ComException thrown. ##\n");
				result.append(L"\n# HRESULT");
				result.append(L"\n  0x");
				wchar_t wszBuffer[20];
				_ltow_s(this->m_hr, wszBuffer, 16);
				result.append(wszBuffer);
				result.append(L" ("); result.append(GetHResultName(this->m_hr)); result.append(L")");
				result.append(L"\n");
				if (this->m_systemMessage.Length() > 0)
				{
					result.append(L"\n# System\n  "); result.append(this->m_systemMessage.operator LPWSTR());
					if (result.at(result.length() - 1) != L'\n')
					{
						result.append(L"\n");
					}
				}
				if (this->m_message.Length() > 0)
				{
					result.append(L"\n# Message\n  "); result.append(this->m_message.operator LPWSTR());
					result.append(L"\n");
				}
				if (this->m_errorInfoMessage.Length() > 0)
				{
					result.append(L"\n# IErrorInfo\n  "); result.append(this->m_errorInfoMessage.operator LPWSTR());
				}
				return CComBSTR(result.c_str());
			}
		};

		struct ComUtil
		{
			//
			// HR inline function for handling HRESULT codes.
			//

			static void HR(HRESULT hr)
			{
				if (FAILED(hr))
				{
					throw ComException(hr);
				}
			}

			static void HR(HRESULT hr, HRESULT ...)
			{
				if (hr == S_OK)
				{
					return;
				}

				va_list list;
				va_start(list, hr);

				bool failed = true;
				for (;;)
				{
					HRESULT argHR = va_arg(list, HRESULT);
					if (argHR == 0)
					{
						break;
					}
					else if (hr == argHR)
					{
						failed = false;
						break;
					}
				}

				if (failed)
				{
					throw ComException(hr);
				}
			}
			//
			// HR inline function for handling HRESULT codes.
			//

			static void HR(HRESULT hr, CComBSTR message)
			{
				if (FAILED(hr))
				{
					throw ComException(hr, message);
				}
			}

			static void HR(HRESULT hr, CComBSTR message, HRESULT ...)
			{
				if (hr == S_OK)
				{
					return;
				}

				va_list list;
				va_start(list, message);

				bool failed = true;
				for (;;)
				{
					HRESULT argHR = va_arg(list, HRESULT);
					if (argHR == 0)
					{
						break;
					}
					else if (hr == argHR)
					{
						failed = false;
						break;
					}
				}

				if (failed)
				{
					throw ComException(hr, message);
				}
			}
		};
		// 
		// ComApartment enum. COM apartment types.
		//
		enum ComApartment
		{
			SingleThreaded = COINIT_APARTMENTTHREADED,
			MultiThreaded = COINIT_MULTITHREADED,
		};

		//
		// ComRuntime class. Auto CoInitializeEx and CoUninitialize.
		//
		class ComRuntime
		{
		private:
			ComApartment m_apartment;
			bool		 m_isInitialized;
		public:
			ComRuntime(ComApartment apartment)
				: m_apartment(apartment),
				m_isInitialized(false)
			{
				ComUtil::HR(::CoInitializeEx(nullptr, apartment));
				this->m_isInitialized = true;
			}
			~ComRuntime()
			{
				if (this->m_isInitialized)
				{
					::CoUninitialize();
				}
			}
		};
	}// namespace ItSoftware::COM
}// namespace ItSoftware