#pragma once
//
// (i) When including this file, remember to in your Main.cpp file to include itsoftware-win.cpp.
//
// (i) uses: Shlwapi.lib
//

//
// #include
//
#include <Windows.h>
#include <WinUser.h>
#include <CommCtrl.h>
#include <memory>
#include <string>
#include <wincrypt.h>
#include <Shlwapi.h>
#include <vector>
#include "itsoftware.h"
#include "itsoftware-exceptions.h"
#include <time.h>

//
// using namespace
//
using namespace ItSoftware;
using namespace ItSoftware::Exceptions;

//
// namespace
//
namespace ItSoftware
{
	namespace Win
	{
		//
		// using
		//
		using std::unique_ptr;
		using std::vector;
		using std::wstring;

		//
		// struct: ItsTimer
		//
		struct ItsTimer
		{
		private:
			size_t m_start = 0;
			size_t m_end = 0;
			bool m_isRunning = false;
		public:
			void Start()
			{
				if (!this->IsRunning())
				{
					this->m_start = GetTickCount64();
					this->m_end = m_start;
					this->m_isRunning = true;
				}
			}

			void Stop()
			{
				this->m_end = GetTickCount64();
				this->m_isRunning = false;
			}

			bool IsRunning()
			{
				return this->m_isRunning;
			}

			size_t LapSeconds()
			{
				if (!this->IsRunning()) {
					return 0;
				}

				return (size_t)((GetTickCount64() - this->m_start) / CLOCKS_PER_SEC);
			}

			size_t LapMilliseconds()
			{
				if (!this->IsRunning()) {
					return 0;
				}

				return (size_t)((GetTickCount64() - this->m_start) / (size_t)(CLOCKS_PER_SEC / 1'000));
			}

			size_t LapMicroseconds()
			{
				if (!this->IsRunning()) {
					return 0;
				}

				return (size_t)((GetTickCount64() - this->m_start) / (size_t)(CLOCKS_PER_SEC / 1'000'000));
			}

			size_t LapTicks()
			{
				if (!this->IsRunning()) {
					return 0;
				}

				return (size_t)GetTickCount64() - this->m_start;
			}

			size_t GetSeconds()
			{
				if (this->IsRunning()) {
					return 0;
				}

				return (size_t)((this->m_end - this->m_start) / (size_t)CLOCKS_PER_SEC);
			}

			size_t GetMilliseconds()
			{
				if (this->IsRunning()) {
					return 0;
				}

				return (size_t)((this->m_end - this->m_start) / (size_t)(CLOCKS_PER_SEC / 1'000));
			}

			size_t GetMicroseconds()
			{
				if (this->IsRunning()) {
					return 0;
				}

				return (size_t)((this->m_end - this->m_start) / (size_t)(CLOCKS_PER_SEC / 1'000'000));
			}

			size_t GetTicks()
			{
				if (this->IsRunning()) {
					return 0;
				}

				return this->m_end - this->m_start;
			}
		};

		//
		// struct: ItsError
		//
		struct ItsError
		{
			//
			// Get GetLastError() Error Description
			//
			static wstring GetLastErrorDescription()
			{
				LPVOID lpMsgBuf;
				FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
					FORMAT_MESSAGE_FROM_SYSTEM |
					FORMAT_MESSAGE_IGNORE_INSERTS,
					NULL,
					GetLastError(),
					MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
					(wchar_t *)&lpMsgBuf,
					0,
					NULL);

				// Copy buffer contet to bstr...
				CComBSTR bstr;
				bstr += (wchar_t *)lpMsgBuf;

				// Free the buffer...
				LocalFree(lpMsgBuf);

				// Return value...
				return wstring(bstr.operator LPWSTR());;
			}// inline CComBSTR GetLastErrorDesc( )

			 //
			 // Get Error Description of Given Error Code
			 //
			static wstring GetErrorDescription(DWORD dwError)
			{
				LPVOID lpMsgBuf;
				FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
					FORMAT_MESSAGE_FROM_SYSTEM |
					FORMAT_MESSAGE_IGNORE_INSERTS,
					NULL,
					dwError,
					MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
					(wchar_t *)&lpMsgBuf,
					0,
					NULL);

				// Copy buffer contet to bstr...
				CComBSTR bstr;
				bstr += (wchar_t *)lpMsgBuf;

				// Free the buffer...
				LocalFree(lpMsgBuf);

				// Return value...
				return wstring(bstr.operator LPWSTR());;
			}// inline CComBSTR GetLastErrorDesc( )

			 //
			 // Get COM IErrorInfo Error Description
			 //
			static wstring GetCoLastErrorInfoDescription()
			{
				IErrorInfo *pErrInfo;
				HRESULT hr = GetErrorInfo(NULL, (IErrorInfo **)&pErrInfo);
				CComBSTR bstr;
				if (hr == S_OK && pErrInfo)
				{
					pErrInfo->GetDescription(&bstr);
					return wstring(bstr.operator LPWSTR());
				}
				return wstring(L"");
			}// inline CComBSTR GetLastErrorDesc( )
		};

		//
		// unique_handle
		//
		template<typename Traits>
		class unique_handle
		{
		private:
			typedef typename Traits::pointer pointer;

			pointer m_value;

		protected:
		public:

			explicit unique_handle(pointer value = Traits::invalid()) noexcept
				: m_value{ value }
			{

			}

			~unique_handle() noexcept
			{
				close();
			}

			auto close() noexcept -> void
			{
				if (*this)
				{
					Traits::close(m_value);
					m_value = Traits::invalid();
				}
			}

			pointer p()
			{
				return m_value;
			}

			pointer* GetAddressOf()
			{
				return &m_value;
			}

			//
			// bool
			//
			explicit operator bool() const noexcept
			{
				return m_value != Traits::invalid();
			}

			//
			// HANDLE
			//
			operator HANDLE() const noexcept
			{
				return static_cast<HANDLE>(m_value);
			}

			void operator=(HANDLE handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HWND
			//
			operator HWND() const noexcept
			{
				return static_cast<HWND>(m_value);
			}

			void operator=(HWND handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HBITMAP
			//
			operator HBITMAP() const noexcept
			{
				return static_cast<HBITMAP>(m_value);
			}

			void operator=(HBITMAP handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HGDIOBJ
			//
			//operator HGDIOBJ() const noexcept
			//{
			//	return static_cast<HGDIOBJ>(m_value);
			//}
			//
			//void operator=(HGDIOBJ handle)
			//{
			//	m_value = static_cast<Traits::pointer>(handle);
			//}

			//
			// HMENU
			//
			operator HMENU() const noexcept
			{
				return static_cast<HMENU>(m_value);
			}

			void operator=(HMENU handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HIMAGELIST
			//
			operator HIMAGELIST() const noexcept
			{
				return static_cast<HIMAGELIST>(m_value);
			}

			void operator=(HIMAGELIST handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HRGN
			//
			operator HRGN() const noexcept
			{
				return static_cast<HRGN>(m_value);
			}

			void operator=(HRGN handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HBRUSH
			//
			operator HBRUSH() const noexcept
			{
				return static_cast<HBRUSH>(m_value);
			}

			void operator=(HBRUSH handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HFONT
			//
			operator HFONT() const noexcept
			{
				return static_cast<HFONT>(m_value);
			}

			void operator=(HFONT handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// HICON
			//
			operator HICON() const noexcept
			{
				return static_cast<HICON>(m_value);
			}

			void operator=(HICON handle)
			{
				m_value = static_cast<Traits::pointer>(handle);
			}

			//
			// IsInvalid
			//
			bool IsInvalid()
			{
				return (m_value == Traits::invalid());
			}

			//
			// IsValid
			//
			bool IsValid()
			{
				return (m_value != Traits::invalid());
			}
		};

		//
		// HANDLE traits
		//
		struct handle_handle_traits
		{
			typedef HANDLE pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				CloseHandle(value);
			}
		};

		//
		// HWND traits
		//
		struct hwnd_handle_traits
		{
			typedef HWND pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					DestroyWindow(value);
				}
			}
		};

		//
		// HBITMAP traits
		//
		struct hbitmap_handle_traits
		{
			typedef HBITMAP pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					DeleteObject(value);
				}
			}
		};

		//
		// HMENU traits
		//
		struct hmenu_handle_traits
		{
			typedef HMENU pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					::DestroyMenu(value);
				}
			}
		};

		//
		// HIMAGELIST traits
		//
		struct himagelist_handle_traits
		{
			typedef HIMAGELIST pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					::ImageList_Destroy(value);
				}
			}
		};

		//
		// HRGN traits
		//
		struct hrgn_handle_traits
		{
			typedef HRGN pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					::DeleteObject(value);
				}
			}
		};

		//
		// HBRUSH traits
		//
		struct hbrush_handle_traits
		{
			typedef HBRUSH pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					::DeleteObject(value);
				}
			}
		};

		//
		// HFONT traits
		//
		struct hfont_handle_traits
		{
			typedef HFONT pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					::DeleteObject(value);
				}
			}
		};

		//
		// HICON traits
		//
		struct hicon_handle_traits
		{
			typedef HICON pointer;
			static auto invalid() noexcept -> pointer
			{
				return nullptr;
			}

			static auto close(pointer value) noexcept -> void
			{
				if (value != invalid())
				{
					::DestroyIcon(value);
				}
			}
		};

		//
		// unique_xxx_handle typedefs
		//
		typedef unique_handle<handle_handle_traits> unique_handle_handle;
		typedef unique_handle<hwnd_handle_traits> unique_hwnd_handle;
		typedef unique_handle<hbitmap_handle_traits> unique_hbitmap_handle;
		typedef unique_handle<hmenu_handle_traits> unique_hmenu_handle;
		typedef unique_handle<himagelist_handle_traits> unique_himagelist_handle;
		typedef unique_handle<hrgn_handle_traits> unique_hrgn_handle;
		typedef unique_handle<hbrush_handle_traits> unique_hbrush_handle;
		typedef unique_handle<hfont_handle_traits> unique_hfont_handle;
		typedef unique_handle<hicon_handle_traits> unique_hicon_handle;

		//
		// class: ItsEvent
		//
		class ItsEvent
		{
		private:
			unique_handle_handle m_handle;
		protected:
		public:
			ItsEvent()
			{
				this->m_handle = CreateEvent(NULL, TRUE, FALSE, NULL);
			}

			ItsEvent(bool signaled)
			{
				this->m_handle = CreateEvent(NULL, TRUE, (signaled) ? TRUE : FALSE, NULL);
			}

			ItsEvent(ItsEvent const &) = delete;
			auto operator=(ItsEvent const &)->ItsEvent & = delete;

			operator HANDLE() noexcept
			{
				return this->m_handle.p();
			}

			void Clear()
			{
				ResetEvent(this->m_handle.p());
			}

			void Signal()
			{
				SetEvent(this->m_handle.p());
			}

			bool Wait(DWORD milliseconds = INFINITE)
			{
				auto result = WaitForSingleObject(this->m_handle.p(), milliseconds);
				return (result == WAIT_OBJECT_0 || result == WAIT_TIMEOUT);
			}
		};

		//
		// enum: ItsFileOpen
		//
		enum class ItsFileOpenCreation
		{
			CreateAlways,
			CreateNew,
			OpenAlways,
			OpenExisting,
			TruncateExisting
		};

		//
		// enum: ItsFilePosition
		//
		enum class ItsFilePosition
		{
			FileBegin,
			FileCurrent,
			FileEnd
		};

		enum class ItsFileTextType
		{
			Ansi,
			UTF8,
			Unicode
		};

		class ItsFileBase
		{
		private:
		protected:
			unique_handle_handle m_handle;

			//
			// Method: GetFileSize
			//
			bool GetFileSize(size_t* size)
			{
				LARGE_INTEGER s{ 0 };
				bool status = (bool)GetFileSizeEx(this->m_handle.p(), &s);
				*size = s.QuadPart;
				return status;
			}
		public:

			//
			// Method: Read
			//
			bool Read(void* pBuffer, DWORD dwBytesToRead, DWORD* dwBytesRead)
			{
				if (this->m_handle.IsInvalid())
				{
					return false;
				}

				return (bool)ReadFile(this->m_handle.p(), pBuffer, dwBytesToRead, dwBytesRead, NULL);
			}

			//
			// Method: Write
			//
			bool Write(void* pBuffer, DWORD dwBytesToWrite, DWORD* dwBytesWritten)
			{
				if (this->m_handle.IsInvalid())
				{
					return false;
				}

				return (bool)WriteFile(this->m_handle.p(), pBuffer, dwBytesToWrite, dwBytesWritten, NULL);
			}

			//
			// Method: Close
			//
			bool Close()
			{
				if (this->m_handle.IsInvalid())
				{
					return false;
				}

				this->m_handle.close();

				return true;
			}

			//
			// Method: GetFilePosition
			//
			bool GetFilePosition(size_t* position)
			{
				if (this->m_handle.IsInvalid())
				{
					return false;
				}

				LARGE_INTEGER zero{ 0 };
				LARGE_INTEGER pos{ 0 };
				bool status = (bool)SetFilePointerEx(this->m_handle.p(), zero, &pos, FILE_CURRENT);
				*position = pos.QuadPart;
				return status;
			}

			//
			// Method: SetFilePosition
			//
			bool SetFilePosition(size_t distanceToMove, ItsFilePosition position)
			{
				if (this->m_handle.IsInvalid())
				{
					return false;
				}

				DWORD dwPosition = 0;
				if (position == ItsFilePosition::FileBegin)
				{
					dwPosition = FILE_BEGIN;
				}
				else if (position == ItsFilePosition::FileCurrent)
				{
					dwPosition = FILE_CURRENT;
				}
				else if (position == ItsFilePosition::FileEnd)
				{
					dwPosition = FILE_END;
				}

				LARGE_INTEGER distance{ 0 };
				distance.QuadPart = distanceToMove;
				return (bool)SetFilePointerEx(this->m_handle.p(), distance, NULL, dwPosition);
			}

			operator HANDLE() const
			{
				return this->m_handle.operator HANDLE();
			}

			bool IsValid()
			{
				return this->m_handle.IsValid();
			}

			bool IsInvalid()
			{
				return this->m_handle.IsInvalid();
			}

			bool SetFileSize(size_t pos)
			{
				if (!this->IsValid())
				{
					return false;
				}
				this->SetFilePosition(pos, ItsFilePosition::FileBegin);
				return (bool)SetEndOfFile(this->m_handle.p());
			}
		};

		//
		// class: ItsTextFile
		//
		class ItsTextFile : public ItsFileBase
		{
		private:
			ItsFileTextType m_textType = ItsFileTextType::Ansi;

			//
			// Method: OpenOrCreate
			//
			bool OpenOrCreate(wstring filename, wstring accessReadWrite, wstring shareReadWrite, ItsFileOpenCreation flagCreation)
			{
				if (this->m_handle.IsValid())
				{
					return false;
				}

				DWORD dwAccess = 0;
				if (accessReadWrite.size() != 0)
				{
					if (accessReadWrite.find(L"r") != wstring::npos) {
						dwAccess |= GENERIC_READ;
					}
					if (accessReadWrite.find(L"w") != wstring::npos) {
						dwAccess |= GENERIC_WRITE;
					}
				}

				DWORD dwShare = 0;
				if (shareReadWrite.size() != 0)
				{
					if (shareReadWrite.find(L"r") != wstring::npos) {
						dwShare |= FILE_SHARE_READ;
					}
					if (shareReadWrite.find(L"w") != wstring::npos) {
						dwShare |= FILE_SHARE_WRITE;
					}
					if (shareReadWrite.find(L"d") != wstring::npos) {
						dwShare |= FILE_SHARE_DELETE;
					}
				}

				DWORD dwFlags = 0;
				if (flagCreation == ItsFileOpenCreation::CreateAlways) {
					dwFlags = CREATE_ALWAYS;
				}
				else if (flagCreation == ItsFileOpenCreation::CreateNew) {
					dwFlags = CREATE_NEW;
				}
				else if (flagCreation == ItsFileOpenCreation::OpenAlways) {
					dwFlags = OPEN_ALWAYS;
				}
				else if (flagCreation == ItsFileOpenCreation::OpenExisting) {
					dwFlags = OPEN_EXISTING;
				}
				else if (flagCreation == ItsFileOpenCreation::TruncateExisting) {
					dwFlags = TRUNCATE_EXISTING;
				}

				this->m_handle = CreateFile(filename.c_str(), dwAccess, dwShare, NULL, dwFlags, FILE_ATTRIBUTE_NORMAL, NULL);

				return this->m_handle.IsValid();
			}

		protected:
		public:

			static const wstring LineDelimiterWindows;
			static const wstring LineDelimiterUnix;
			static const vector<BYTE> ByteOrderMarkUnicode;
			static const vector<BYTE> ByteOrderMarkUTF8;

			//
			// Constructor
			//
			ItsTextFile()
			{
			}

			//
			// Method: OpenOrCreateText
			//
			bool OpenOrCreateText(wstring filename, wstring accessReadWrite, wstring shareReadWrite, ItsFileOpenCreation flagCreation, ItsFileTextType type)
			{
				if (this->m_handle.IsValid())
				{
					return false;
				}

				if (!this->OpenOrCreate(filename, accessReadWrite, shareReadWrite, flagCreation))
				{
					return false;
				}

				size_t size{ 0 };
				this->GetFileSize(&size);
				if (size > 2)
				{
					this->SetFilePosition(0, ItsFilePosition::FileBegin);

					unique_ptr<BYTE[]> bom = make_unique<BYTE[]>(3);

					DWORD dwRead{ 0 };
					this->Read(bom.get(), 3, &dwRead);

					if (dwRead > 2) {
						if (bom.get()[0] == 0xFF &&
							bom.get()[1] == 0xFE)
						{
							this->SetFilePosition(2, ItsFilePosition::FileBegin);

							this->m_textType = ItsFileTextType::Unicode;
							return true;
						}
						else if (bom.get()[0] == 0xEF &&
							bom.get()[1] == 0xBB &&
							bom.get()[2] == 0xBF)
						{
							this->SetFilePosition(3, ItsFilePosition::FileBegin);

							this->m_textType = ItsFileTextType::UTF8;
							return true;
						}
					}
				}

				this->m_textType = type;

				if (type == ItsFileTextType::Ansi)
				{
					return true;
				}

				vector<BYTE> bom;
				if (type == ItsFileTextType::Unicode) {
					bom.push_back(0xFF);
					bom.push_back(0xFE);
				}
				else if (type == ItsFileTextType::UTF8)
				{
					bom.push_back(0xEF);
					bom.push_back(0xBB);
					bom.push_back(0xBF);
				}

				this->SetFilePosition(0, ItsFilePosition::FileBegin);

				DWORD dwWritten{ 0 };
				return this->Write(bom.data(), (DWORD)bom.size(), &dwWritten);
			}

			//
			// Method: ReadTextAllLines
			//
			bool ReadTextAllLines(vector<wstring>& lines, wstring lineDelimiter = L"\r\n")
			{
				if (!this->m_handle.IsValid())
				{
					return false;
				}

				size_t size{ 0 };
				if (!this->GetFileSize(&size))
				{
					return false;
				}

				if (size == 0) {
					lines.clear();
					return true;
				}

				if (this->m_textType == ItsFileTextType::Ansi)
				{
					this->SetFilePosition(0, ItsFilePosition::FileBegin);

					unique_ptr<BYTE[]> text = make_unique<BYTE[]>(size + 1);

					const DWORD readBuffSize = 8192;
					size_t read_pos{ 0 };

					DWORD dwRead{ 0 };

					this->Read((BYTE*)text.get(), readBuffSize, &dwRead);

					read_pos += dwRead;

					while (dwRead > 0 && (read_pos + 0) < size)
					{
						this->Read((BYTE*)&(text.get()[read_pos]), readBuffSize, &dwRead);

						read_pos += dwRead;
					}

					text.get()[size] = '\0';

					string all((char*)text.get());
					wstring out_all(all.begin(), all.end());
					lines = ItsString::Split(out_all, lineDelimiter);
				}
				else if (this->m_textType == ItsFileTextType::Unicode)
				{
					this->SetFilePosition(2, ItsFilePosition::FileBegin);

					unique_ptr<BYTE[]> text = make_unique<BYTE[]>(size);

					const DWORD readBuffSize = 8192;
					size_t read_pos{ 0 };

					DWORD dwRead{ 0 };

					this->Read((BYTE*)text.get(), readBuffSize, &dwRead);

					read_pos += dwRead;

					while (dwRead > 0 && (read_pos + 2) < size)
					{
						this->Read((BYTE*)&(text.get()[read_pos]), readBuffSize, &dwRead);

						read_pos += dwRead;
					}

					text.get()[size - 2] = '\0';
					text.get()[size - 1] = '\0';

					wstring all((wchar_t*)text.get());
					lines = ItsString::Split(all, lineDelimiter);
				}
				else if (this->m_textType == ItsFileTextType::UTF8)
				{
					this->SetFilePosition(3, ItsFilePosition::FileBegin);

					unique_ptr<BYTE[]> text = make_unique<BYTE[]>(size);

					const DWORD readBuffSize = 8192;
					size_t read_pos{ 0 };

					DWORD dwRead{ 0 };

					this->Read((BYTE*)text.get(), readBuffSize, &dwRead);

					read_pos += dwRead;

					while (dwRead > 0 && (read_pos + 3) < size)
					{
						this->Read((BYTE*)&(text.get()[read_pos]), readBuffSize, &dwRead);

						read_pos += dwRead;
					}

					text.get()[size - 3] = '\0';
					text.get()[size - 2] = '\0';
					text.get()[size - 1] = '\0';

					wstring all = ItSoftware::Encoding::UTF8::ToString((char*)text.get());
					lines = ItsString::Split(all, lineDelimiter);
				}
				else {
					return false;
				}

				return true;
			}

			//
			// Method: ReadTextAll
			//
			bool ReadTextAll(wstring& out)
			{
				if (!this->m_handle.IsValid())
				{
					return false;
				}

				size_t size{ 0 };
				if (!this->GetFileSize(&size))
				{
					return false;
				}

				if (size == 0) {
					return true;
				}

				if (this->m_textType == ItsFileTextType::Ansi)
				{
					this->SetFilePosition(0, ItsFilePosition::FileBegin);

					unique_ptr<BYTE[]> text = make_unique<BYTE[]>(size + 1);

					const DWORD readBuffSize = 8192;
					size_t read_pos{ 0 };

					DWORD dwRead{ 0 };

					this->Read((BYTE*)text.get(), readBuffSize, &dwRead);

					read_pos += dwRead;

					while (dwRead > 0 && (read_pos + 0) < size)
					{
						this->Read((BYTE*)&(text.get()[read_pos]), readBuffSize, &dwRead);

						read_pos += dwRead;
					}

					text.get()[size] = '\0';

					string all((char*)text.get());
					wstring out_all(all.begin(), all.end());
					out = out_all;
				}
				else if (this->m_textType == ItsFileTextType::Unicode)
				{
					this->SetFilePosition(2, ItsFilePosition::FileBegin);

					unique_ptr<BYTE[]> text = make_unique<BYTE[]>(size);

					const DWORD readBuffSize = 8192;
					size_t read_pos{ 0 };

					DWORD dwRead{ 0 };

					this->Read((BYTE*)text.get(), readBuffSize, &dwRead);

					read_pos += dwRead;

					while (dwRead > 0 && (read_pos + 2) < size)
					{
						this->Read((BYTE*)&(text.get()[read_pos]), readBuffSize, &dwRead);

						read_pos += dwRead;
					}

					text.get()[size - 2] = '\0';
					text.get()[size - 1] = '\0';

					wstring all((wchar_t*)text.get());
					out = all;
				}
				else if (this->m_textType == ItsFileTextType::UTF8)
				{
					this->SetFilePosition(3, ItsFilePosition::FileBegin);

					unique_ptr<BYTE[]> text = make_unique<BYTE[]>(size);

					const DWORD readBuffSize = 8192;
					size_t read_pos{ 0 };

					DWORD dwRead{ 0 };

					this->Read((BYTE*)text.get(), readBuffSize, &dwRead);

					read_pos += dwRead;

					while (dwRead > 0 && (read_pos + 3) < size)
					{
						this->Read((BYTE*)&(text.get()[read_pos]), readBuffSize, &dwRead);

						read_pos += dwRead;
					}

					text.get()[size - 3] = '\0';
					text.get()[size - 2] = '\0';
					text.get()[size - 1] = '\0';

					wstring all = ItSoftware::Encoding::UTF8::ToString((char*)text.get());
					out = all;
				}
				else {
					return false;
				}

				return true;
			}

			//
			// Method: WriteText
			//
			bool WriteText(string text)
			{
				wstring txt(text.begin(), text.end());
				return this->WriteText(txt);
			}

			//
			// Method: WriteText
			//
			bool WriteText(wstring text)
			{
				if (this->m_handle.IsInvalid())
				{
					return false;
				}

				if (this->m_textType == ItsFileTextType::Ansi)
				{
					string txt(text.begin(), text.end());
					DWORD dwWritten{ 0 };
					return this->Write((BYTE*)txt.data(), (DWORD)txt.size(), &dwWritten);
				}
				else if (this->m_textType == ItsFileTextType::Unicode) {
					DWORD dwWritten{ 0 };
					return this->Write((BYTE*)text.data(), (DWORD)(text.size() * sizeof(wchar_t)), &dwWritten);
				}
				else if (this->m_textType == ItsFileTextType::UTF8) {
					long cbLength{ 0 };
					BYTE* pData = ItSoftware::Encoding::UTF8::ToBytes(text, &cbLength);
					DWORD dwWritten{ 0 };
					return this->Write(pData, cbLength, &dwWritten);
				}

				return false;
			}

			//
			// Method: WriteTextLine
			//
			bool WriteTextLine(string text, string lineDelimiter = "\r\n")
			{
				wstring txt(text.begin(), text.end());
				wstring dl(lineDelimiter.begin(), lineDelimiter.end());
				return this->WriteTextLine(txt, dl);
			}

			//
			// Method: WriteTextLine
			//
			bool WriteTextLine(wstring text, wstring lineDelimiter = L"\r\n")
			{
				if (this->m_handle.IsInvalid())
				{
					return false;
				}

				if (!this->WriteText(text))
				{
					return false;
				}

				if (!this->WriteText(lineDelimiter))
				{
					return false;
				}

				return true;
			}
		};



		//
		// class: ItsFile
		//
		class ItsFile : public ItsFileBase
		{
		private:
		protected:
		public:

			//
			// Constructor
			//
			ItsFile()
			{
			}

			//
			// Method: OpenOrCreate
			//
			bool OpenOrCreate(wstring filename, wstring accessReadWrite, wstring shareReadWrite, ItsFileOpenCreation flagCreation)
			{
				if (this->m_handle.IsValid())
				{
					return false;
				}

				DWORD dwAccess = 0;
				if (accessReadWrite.size() != 0)
				{
					if (accessReadWrite.find(L"r") != wstring::npos) {
						dwAccess |= GENERIC_READ;
					}
					if (accessReadWrite.find(L"w") != wstring::npos) {
						dwAccess |= GENERIC_WRITE;
					}
				}

				DWORD dwShare = 0;
				if (shareReadWrite.size() != 0)
				{
					if (shareReadWrite.find(L"r") != wstring::npos) {
						dwShare |= FILE_SHARE_READ;
					}
					if (shareReadWrite.find(L"w") != wstring::npos) {
						dwShare |= FILE_SHARE_WRITE;
					}
					if (shareReadWrite.find(L"d") != wstring::npos) {
						dwShare |= FILE_SHARE_DELETE;
					}
				}

				DWORD dwFlags = 0;
				if (flagCreation == ItsFileOpenCreation::CreateAlways) {
					dwFlags = CREATE_ALWAYS;
				}
				else if (flagCreation == ItsFileOpenCreation::CreateNew) {
					dwFlags = CREATE_NEW;
				}
				else if (flagCreation == ItsFileOpenCreation::OpenAlways) {
					dwFlags = OPEN_ALWAYS;
				}
				else if (flagCreation == ItsFileOpenCreation::OpenExisting) {
					dwFlags = OPEN_EXISTING;
				}
				else if (flagCreation == ItsFileOpenCreation::TruncateExisting) {
					dwFlags = TRUNCATE_EXISTING;
				}

				this->m_handle = CreateFile(filename.c_str(), dwAccess, dwShare, NULL, dwFlags, FILE_ATTRIBUTE_NORMAL, NULL);

				return this->m_handle.IsValid();
			}

			//
			// Method: GetFileSize
			//
			static bool GetFileSize(wstring filename, size_t* size)
			{
				ItsFile file;
				if (!file.OpenOrCreate(filename, L"r", L"r", ItsFileOpenCreation::OpenExisting))
				{
					return false;
				}

				LARGE_INTEGER tmp{ 0 };
				bool result = (bool)GetFileSizeEx(file.m_handle.p(), &tmp);
				*size = tmp.QuadPart;
				return result;
			}

			//
			// Method: Delete
			//
			static bool Delete(wstring filename)
			{
				return (bool)DeleteFile(filename.c_str());
			}

			//
			// Method: Copy
			//
			static bool Copy(wstring existingFilename, wstring newFilename, bool failIfExists)
			{
				return (bool)CopyFile(existingFilename.c_str(), newFilename.c_str(), (failIfExists) ? TRUE : FALSE);
			}

			//
			// Method: Move
			//
			static bool Move(wstring existingFilename, wstring newFilename, bool failIfExists)
			{
				return (bool)MoveFileEx(existingFilename.c_str(), newFilename.c_str(), MOVEFILE_COPY_ALLOWED | ((failIfExists) ? MOVEFILE_REPLACE_EXISTING : 0));
			}

			//
			// Method: Exists
			//
			static bool Exists(wstring filename)
			{
				return (bool)PathFileExists(filename.c_str());
			}

			//
			// Method: SetFileSize
			//
			static bool SetSize(wstring filename, size_t pos)
			{
				ItsFile file;
				if (!file.OpenOrCreate(filename, L"rw", L"", ItsFileOpenCreation::OpenExisting)) {
					return false;
				}
				return file.SetFileSize(pos);
			}
		};

		//
		// struct: ItsWin
		//
		struct ItsWin
		{
		private:
			//
			// Variable: hCBTHook
			//
			static HHOOK hCBTHook;

			//
			// Function: CBTMsgBox
			//
			static LRESULT CALLBACK CBTMsgBox(int code, WPARAM wParam, LPARAM lParam)
			{
				if (HCBT_ACTIVATE != code || -1 == code)
				{
					// only when HCBT_ACTIVATE 
					return CallNextHookEx(ItsWin::hCBTHook, code, wParam, lParam);
				}

				HWND hWnd = (HWND)wParam;
				UnhookWindowsHookEx(ItsWin::hCBTHook);
				CenterToItsParentWindow(hWnd);

				return CallNextHookEx(ItsWin::hCBTHook, code, wParam, lParam);
			}

		public:

			//
			//
			//
			static wstring ToRGBString(COLORREF color)
			{
				wstringstream ss;

				ss << (unsigned char)((color & 0x0000FF));
				ss << L",";
				ss << (unsigned char)((color & 0x00FF00) >> 8);
				ss << L",";
				ss << (unsigned char)((color & 0xFF0000) >> 16);

				return ss.str();
			}

			//
			// Function: CenterToParentWindow
			//
			static void CenterToItsParentWindow(HWND hDlg)
			{
				RECT rcDlg;
				RECT rcOwner;

				//center
				HWND hWndOwner = ::GetWindow(hDlg, GW_OWNER);
				GetWindowRect(hDlg, &rcDlg);
				GetWindowRect(hWndOwner, &rcOwner);

				int iLeft = rcOwner.left + ((rcOwner.right - rcOwner.left)
					- (rcDlg.right - rcDlg.left)) / 2;

				int iTop = rcOwner.top + ((rcOwner.bottom - rcOwner.top)
					- (rcDlg.bottom - rcDlg.top)) / 2;

				SetWindowPos(hDlg, hWndOwner, iLeft, iTop, 0, 0, SWP_NOSIZE | SWP_FRAMECHANGED);
			}

			//
			// ## ShowInformationMessageXXX #################################
			//

			//
			// Function: ShowInformationMessage
			//
			static void ShowInformationMessage(wstring message)
			{
				ShowInformationMessage(NULL, L"Information!", message);
			}

			//
			// Function: ShowInformationMessage
			//
			static void ShowInformationMessage(wstring title, wstring message)
			{
				ShowInformationMessage(NULL, title, message);
			}

			//
			// Function: ShowInformationMessage
			//
			static void ShowInformationMessage(HWND hWnd, wstring message)
			{
				ShowInformationMessage(hWnd, L"Information!", message);
			}

			//
			// Function: ShowInformationMessage
			//
			static void ShowInformationMessage(HWND hWnd, wstring title, wstring message)
			{
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONINFORMATION);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONINFORMATION);
				}
			}

			//
			// Function: ShowInformationMessageYesNo
			//
			static int ShowInformationMessageYesNo(wstring message)
			{
				return ShowInformationMessageYesNo(NULL, L"Information!", message);
			}

			//
			// Function: ShowInformationMessageYesNo
			//
			static int ShowInformationMessageYesNo(wstring title, wstring message)
			{
				return ShowInformationMessageYesNo(NULL, title, message);
			}

			//
			// Function: ShowInformationMessageYesNo
			//
			static int ShowInformationMessageYesNo(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONINFORMATION | MB_YESNO);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONINFORMATION | MB_YESNO);
				}

				return result;
			}

			//
			// Function: ShowInformationMessageYesNoCancel
			//
			static int ShowInformationMessageYesNoCancel(wstring message)
			{
				return ShowInformationMessageYesNoCancel(NULL, L"Information!", message);
			}

			//
			// Function: ShowInformationMessageYesNoCancel
			//
			static int ShowInformationMessageYesNoCancel(wstring title, wstring message)
			{
				return ShowInformationMessageYesNoCancel(NULL, title, message);
			}

			//
			// Function: ShowInformationMessageYesNoCancel
			//
			static int ShowInformationMessageYesNoCancel(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONINFORMATION | MB_YESNOCANCEL);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONINFORMATION | MB_YESNOCANCEL);
				}

				return result;
			}

			//
			// Function: ShowInformationMessageAbortRetryIgnore
			//
			static int ShowInformationMessageAbortRetryIgnore(wstring message)
			{
				return ShowInformationMessageAbortRetryIgnore(NULL, L"Information!", message);
			}

			//
			// Function: ShowInformationMessageAbortRetryIgnore
			//
			static int ShowInformationMessageAbortRetryIgnore(wstring title, wstring message)
			{
				return ShowInformationMessageAbortRetryIgnore(NULL, title, message);
			}

			//
			// Function: ShowWarningMessageAbortRetryIgnore
			//
			static int ShowInformationMessageAbortRetryIgnore(HWND hWnd, wstring message)
			{
				return ShowInformationMessageAbortRetryIgnore(hWnd, L"Information!", message);
			}

			//
			// Function: ShowInformationMessageAbortRetryIgnore
			//
			static int ShowInformationMessageAbortRetryIgnore(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONINFORMATION | MB_ABORTRETRYIGNORE);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONINFORMATION | MB_ABORTRETRYIGNORE);
				}

				return result;
			}

			//
			// ## ShowWarningMessageXXX ##################################
			//

			//
			// ShowWarningMessage
			//
			static void ShowWarningMessage(wstring message)
			{
				ShowWarningMessage(NULL, L"Warning!", message);
			}

			//
			// ShowWarningMessage
			//
			static void ShowWarningMessage(wstring title, wstring message)
			{
				ShowWarningMessage(NULL, title, message);
			}

			//
			// ShowWarningMessage
			//
			static void ShowWarningMessage(HWND hWnd, wstring message)
			{
				ShowWarningMessage(hWnd, L"Warning!", message);
			}

			//
			// Function: ShowWarningMessage
			//
			static void ShowWarningMessage(HWND hWnd, wstring title, wstring message)
			{
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONWARNING);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONWARNING);
				}
			}

			//
			// Function: ShowWarningMessageYesNo
			//
			static int ShowWarningMessageYesNo(wstring message)
			{
				return ShowWarningMessageYesNo(NULL, L"Warning!", message);
			}

			//
			// Function: ShowWarningMessageYesNo
			//
			static int ShowWarningMessageYesNo(wstring title, wstring message)
			{
				return ShowWarningMessageYesNo(NULL, title, message);
			}

			//
			// Function: ShowWarningMessageYesNo
			//
			static int ShowWarningMessageYesNo(HWND hWnd, wstring message)
			{
				return ShowWarningMessageYesNo(hWnd, L"Warning!", message);
			}

			//
			// Function: ShowWarningMessageYesNo
			//
			static int ShowWarningMessageYesNo(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONWARNING | MB_YESNO);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONWARNING | MB_YESNO);
				}

				return result;
			}

			//
			// Function: ShowWarningMessageYesNoCancel
			//
			static int ShowWarningMessageYesNoCancel(wstring message)
			{
				return ShowWarningMessageYesNoCancel(NULL, L"Warning!", message);
			}

			//
			// Function: ShowWarningMessageYesNoCancel
			//
			static int ShowWarningMessageYesNoCancel(wstring title, wstring message)
			{
				return ShowWarningMessageYesNoCancel(NULL, title, message);
			}

			//
			// Function: ShowWarningMessageYesNoCancel
			//
			static int ShowWarningMessageYesNoCancel(HWND hWnd, wstring message)
			{
				return ShowWarningMessageYesNoCancel(hWnd, L"Warning!", message);
			}

			//
			// Function: ShowWarningMessageYesNoCancel
			//
			static int ShowWarningMessageYesNoCancel(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONWARNING | MB_YESNO);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONWARNING | MB_YESNO);
				}

				return result;
			}

			//
			// Function: ShowWarningMessageAbortRetryIgnore
			//
			static int ShowWarningMessageAbortRetryIgnore(wstring message)
			{
				return ShowWarningMessageAbortRetryIgnore(NULL, L"Warning!", message);
			}

			//
			// Function: ShowWarningMessageAbortRetryIgnore
			//
			static int ShowWarningMessageAbortRetryIgnore(wstring title, wstring message)
			{
				return ShowWarningMessageAbortRetryIgnore(NULL, title, message);
			}

			//
			// Function: ShowWarningMessageAbortRetryIgnore
			//
			static int ShowWarningMessageAbortRetryIgnore(HWND hWnd, wstring message)
			{
				return ShowWarningMessageAbortRetryIgnore(hWnd, L"Warning!", message);
			}

			//
			// Function: ShowWarningMessageAbortRetryIgnore
			//
			static int ShowWarningMessageAbortRetryIgnore(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONWARNING | MB_ABORTRETRYIGNORE);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONWARNING | MB_ABORTRETRYIGNORE);
				}

				return result;
			}

			//
			// ## ShowHandMessageXXX #################################
			//

			//
			// Function: ShowHandMessage
			//
			static void ShowHandMessage(wstring message)
			{
				ShowHandMessage(NULL, L"Hand!", message);
			}

			//
			// Function: ShowHandMessage
			//
			static void ShowHandMessage(wstring title, wstring message)
			{
				ShowHandMessage(NULL, title, message);
			}

			//
			// Function: ShowHandMessage
			//
			static void ShowHandMessage(HWND hWnd, wstring message)
			{
				ShowHandMessage(hWnd, L"Hand!", message);
			}

			//
			// Function: ShowHandMessage
			//
			static void ShowHandMessage(HWND hWnd, wstring title, wstring message)
			{
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONHAND);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONHAND);
				}
			}

			//
			// Function: ShowHandMessageYesNo
			//
			static int ShowHandMessageYesNo(wstring message)
			{
				return ShowHandMessageYesNo(NULL, L"Hand!", message);
			}

			//
			// Function: ShowHandMessageYesNo
			//
			static int ShowHandMessageYesNo(wstring title, wstring message)
			{
				return ShowHandMessageYesNo(NULL, title, message);
			}

			//
			// Function: ShowHandMessageYesNo
			//
			static int ShowHandMessageYesNo(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONHAND | MB_YESNO);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONHAND | MB_YESNO);
				}

				return result;
			}

			//
			// Function: ShowHandMessageYesNoCancel
			//
			static int ShowHandMessageYesNoCancel(wstring message)
			{
				return ShowHandMessageYesNoCancel(NULL, L"Hand!", message);
			}

			//
			// Function: ShowHandMessageYesNoCancel
			//
			static int ShowHandMessageYesNoCancel(wstring title, wstring message)
			{
				return ShowHandMessageYesNoCancel(NULL, title, message);
			}

			//
			// Function: ShowHandMessageYesNoCancel
			//
			static int ShowHandMessageYesNoCancel(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONHAND | MB_YESNOCANCEL);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONHAND | MB_YESNOCANCEL);
				}

				return result;
			}

			//
			// Function: ShowHandMessageAbortRetryIgnore
			//
			static int ShowHandMessageAbortRetryIgnore(wstring message)
			{
				return ShowHandMessageAbortRetryIgnore(NULL, L"Hand!", message);
			}

			//
			// Function: ShowHandMessageAbortRetryIgnore
			//
			static int ShowHandMessageAbortRetryIgnore(wstring title, wstring message)
			{
				return ShowHandMessageAbortRetryIgnore(NULL, title, message);
			}

			//
			// Function: ShowHandMessageAbortRetryIgnore
			//
			static int ShowHandMessageAbortRetryIgnore(HWND hWnd, wstring message)
			{
				return ShowHandMessageAbortRetryIgnore(hWnd, L"Hand!", message);
			}

			//
			// Function: ShowHandMessageAbortRetryIgnore
			//
			static int ShowHandMessageAbortRetryIgnore(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONHAND | MB_ABORTRETRYIGNORE);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONHAND | MB_ABORTRETRYIGNORE);
				}

				return result;
			}

			//
			// ## ShowErrorMessageXXX ################################
			//

			//
			// Function: ShowErrorMessage
			//
			static void ShowErrorMessage(wstring message)
			{
				ShowErrorMessage(NULL, L"Error!", message);
			}

			//
			// Function: ShowErrorMessage
			//
			static void ShowErrorMessage(wstring title, wstring message)
			{
				ShowErrorMessage(NULL, title, message);
			}

			//
			// Function: ShowErrorMessage
			//
			static void ShowErrorMessage(HWND hWnd, wstring message)
			{
				ShowErrorMessage(hWnd, L"Error!", message);
			}

			//
			// Function: ShowErrorMessage
			//
			static void ShowErrorMessage(HWND hWnd, wstring title, wstring message)
			{
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONERROR);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					MessageBox(hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONERROR);
				}
			}

			//
			// Function: ShowErrorMessageYesNo
			//
			static int ShowErrorMessageYesNo(wstring message)
			{
				return ShowErrorMessageYesNo(NULL, L"Error!", message);
			}

			//
			// Function: ShowErrorMessageYesNo
			//
			static int ShowErrorMessageYesNo(wstring title, wstring message)
			{
				return ShowErrorMessageYesNo(NULL, title, message);
			}

			//
			// Function: ShowErrorMessageYesNo
			//
			static int ShowErrorMessageYesNo(HWND hWnd, wstring message)
			{
				return ShowErrorMessageYesNo(hWnd, L"Error!", message);
			}

			//
			// Function: ShowErrorMessageYesNo
			//
			static int ShowErrorMessageYesNo(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONERROR | MB_YESNO);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONERROR | MB_YESNO);
				}

				return result;
			}

			//
			// Function: ShowErrorMessageYesNoCancel
			//
			static int ShowErrorMessageYesNoCancel(wstring message)
			{
				return ShowErrorMessageYesNoCancel(NULL, L"Error!", message);
			}

			//
			// Function: ShowErrorMessageYesNoCancel
			//
			static int ShowErrorMessageYesNoCancel(wstring title, wstring message)
			{
				return ShowErrorMessageYesNoCancel(NULL, title, message);
			}

			//
			// Function: ShowErrorMessageYesNoCancel
			//
			static int ShowErrorMessageYesNoCancel(HWND hWnd, wstring message)
			{
				return ShowErrorMessageYesNoCancel(hWnd, L"Error!", message);
			}

			//
			// Function: ShowErrorMessageYesNoCancel
			//
			static int ShowErrorMessageYesNoCancel(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONERROR | MB_YESNOCANCEL);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONERROR | MB_YESNOCANCEL);
				}

				return result;
			}

			//
			// Function: ShowErrorMessageAbortRetryIgnore
			//
			static int ShowErrorMessageAbortRetryIgnore(wstring message)
			{
				return ShowErrorMessageAbortRetryIgnore(NULL, L"Error!", message);
			}

			//
			// Function: ShowErrorMessageAbortRetryIgnore
			//
			static int ShowErrorMessageAbortRetryIgnore(wstring title, wstring message)
			{
				return ShowErrorMessageAbortRetryIgnore(NULL, title, message);
			}

			//
			// Function: ShowErrorMessageAbortRetryIgnore
			//
			static int ShowErrorMessageAbortRetryIgnore(HWND hWnd, wstring message)
			{
				return ShowErrorMessageAbortRetryIgnore(hWnd, L"Error!", message);
			}

			//
			// Function: ShowErrorMessageAbortRetryIgnore
			//
			static int ShowErrorMessageAbortRetryIgnore(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONERROR | MB_ABORTRETRYIGNORE);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONERROR | MB_ABORTRETRYIGNORE);
				}

				return result;
			}

			//
			// ## ShowQuestionMessageXXX
			//

			//
			// Function: ShowQuestionMessageYesNo
			//
			static int ShowQuestionMessageYesNo(wstring message)
			{
				return ShowQuestionMessageYesNo(NULL, L"Question!", message);
			}

			//
			// Function: ShowQuestionMessageYesNo
			//
			static int ShowQuestionMessageYesNo(wstring title, wstring message)
			{
				return ShowQuestionMessageYesNo(NULL, title, message);
			}

			//
			// Function: ShowQuestionMessageYesNo
			//
			static int ShowQuestionMessageYesNo(HWND hWnd, wstring message)
			{
				return ShowQuestionMessageYesNo(hWnd, L"Question!", message);
			}

			//
			// Function: ShowQuestionMessageYesNo
			//
			static int ShowQuestionMessageYesNo(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONQUESTION | MB_YESNO);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONQUESTION | MB_YESNO);
				}

				return result;
			}

			//
			// Function: ShowQuestionMessageYesNoCancel
			//
			static int ShowQuestionMessageYesNoCancel(wstring message)
			{
				return ShowQuestionMessageYesNoCancel(NULL, L"Question!", message);
			}

			//
			// Function: ShowQuestionMessageYesNoCancel
			//
			static int ShowQuestionMessageYesNoCancel(wstring title, wstring message)
			{
				return ShowQuestionMessageYesNoCancel(NULL, title, message);
			}

			//
			// Function: ShowQuestionMessageYesNoCancel
			//
			static int ShowQuestionMessageYesNoCancel(HWND hWnd, wstring message)
			{
				return ShowQuestionMessageYesNoCancel(hWnd, L"Question!", message);
			}

			//
			// Function: ShowQuestionMessageYesNoCancel
			//
			static int ShowQuestionMessageYesNoCancel(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONQUESTION | MB_YESNOCANCEL);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONQUESTION | MB_YESNOCANCEL);
				}

				return result;
			}

			//
			// Function: ShowQuestionMessageAbortRetryIgnore
			//
			static int ShowQuestionMessageAbortRetryIgnore(wstring message)
			{
				return ShowQuestionMessageAbortRetryIgnore(NULL, L"Question!", message);
			}

			//
			// Function: ShowQuestionMessageAbortRetryIgnore
			//
			static int ShowQuestionMessageAbortRetryIgnore(wstring title, wstring message)
			{
				return ShowQuestionMessageAbortRetryIgnore(NULL, title, message);
			}

			//
			// Function: ShowQuestionMessageAbortRetryIgnore
			//
			static int ShowQuestionMessageAbortRetryIgnore(HWND hWnd, wstring message)
			{
				return ShowQuestionMessageAbortRetryIgnore(hWnd, L"Question!", message);
			}

			//
			// Function: ShowQuestionMessageAbortRetryIgnore
			//
			static int ShowQuestionMessageAbortRetryIgnore(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONQUESTION | MB_ABORTRETRYIGNORE);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONQUESTION | MB_ABORTRETRYIGNORE);
				}

				return result;
			}

			//
			// ## ShowStopMessageXXX
			//

			//
			// Function: ShowStopMessageYesNo
			//
			static int ShowStopMessageYesNo(wstring message)
			{
				return ShowStopMessageYesNo(NULL, L"Stop!", message);
			}

			//
			// Function: ShowStopMessageYesNo
			//
			static int ShowStopMessageYesNo(wstring title, wstring message)
			{
				return ShowStopMessageYesNo(NULL, title, message);
			}

			//
			// Function: ShowStopMessageYesNo
			//
			static int ShowStopMessageYesNo(HWND hWnd, wstring message)
			{
				return ShowStopMessageYesNo(hWnd, L"Stop!", message);
			}

			//
			// Function: ShowStopMessageYesNo
			//
			static int ShowStopMessageYesNo(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONSTOP | MB_YESNO);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONSTOP | MB_YESNO);
				}

				return result;
			}

			//
			// Function: ShowStopMessageYesNoCancel
			//
			static int ShowStopMessageYesNoCancel(wstring message)
			{
				return ShowStopMessageYesNoCancel(NULL, L"Stop!", message);
			}

			//
			// Function: ShowStopMessageYesNoCancel
			//
			static int ShowStopMessageYesNoCancel(wstring title, wstring message)
			{
				return ShowStopMessageYesNoCancel(NULL, title, message);
			}

			//
			// Function: ShowStopMessageYesNoCancel
			//
			static int ShowStopMessageYesNoCancel(HWND hWnd, wstring message)
			{
				return ShowStopMessageYesNoCancel(hWnd, L"Stop!", message);
			}

			//
			// Function: ShowStopMessageYesNoCancel
			//
			static int ShowStopMessageYesNoCancel(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONSTOP | MB_YESNOCANCEL);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONSTOP | MB_YESNOCANCEL);
				}

				return result;
			}

			//
			// Function: ShowStopMessageAbortRetryIgnore
			//
			static int ShowStopMessageAbortRetryIgnore(wstring message)
			{
				return ShowStopMessageAbortRetryIgnore(NULL, L"Stop!", message);
			}

			//
			// Function: ShowStopMessageAbortRetryIgnore
			//
			static int ShowStopMessageAbortRetryIgnore(wstring title, wstring message)
			{
				return ShowStopMessageAbortRetryIgnore(NULL, title, message);
			}

			//
			// Function: ShowStopMessageAbortRetryIgnore
			//
			static int ShowStopMessageAbortRetryIgnore(HWND hWnd, wstring message)
			{
				return ShowStopMessageAbortRetryIgnore(hWnd, L"Stop!", message);
			}

			//
			// Function: ShowStopMessageYesNoCancel
			//
			static int ShowStopMessageAbortRetryIgnore(HWND hWnd, wstring title, wstring message)
			{
				int result = 0;
				if (hWnd != NULL)
				{
					ItsWin::hCBTHook = SetWindowsHookEx(WH_CBT, CBTMsgBox, NULL, ::GetCurrentThreadId());
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONSTOP | MB_ABORTRETRYIGNORE);
					UnhookWindowsHookEx(ItsWin::hCBTHook);
				}
				else
				{
					result = MessageBox(hWnd, message.c_str(), title.c_str(), MB_ICONSTOP | MB_ABORTRETRYIGNORE);
				}

				return result;
			}

			//
			// ReportSuccessEventLogEntry ###########################################
			//

			//
			// Function: ReportSuccessEventLogEntry
			//
			static bool ReportSuccessEventLogEntry(const wstring sourcename, const wstring text)
			{
				if (sourcename.size() == 0)
				{
					return false;
				}

				return ReportSuccessEventLogEntry(sourcename, 0, 0, 1, text.c_str());
			}

			//
			// Function: ReportSuccessEventLogEntry
			//
			static bool ReportSuccessEventLogEntry(wstring sourcename, short category, long eventid, short numberofstrings, const wchar_t* description)
			{
				if (sourcename.size() == 0 || (numberofstrings > 0 && wcslen(description) == 0))
				{
					return false;
				}

				HANDLE hEventLog = ::RegisterEventSource(NULL, sourcename.c_str());
				if (hEventLog == NULL)
				{
					return false;
				}

				BOOL bStatus = ::ReportEvent(hEventLog, EVENTLOG_SUCCESS, category, eventid, NULL, numberofstrings, 0, (LPCWSTR*)&description, NULL);
				if (!bStatus)
				{
					::DeregisterEventSource(hEventLog);
					return false;
				}

				::DeregisterEventSource(hEventLog);

				return true;
			}

			//
			// ReportErrorEventLogEntry ############################################
			//

			//
			// Function: ReportErrorEventLogEntry
			//
			static bool ReportErrorEventLogEntry(const wstring sourcename, const wstring text)
			{
				if (sourcename.size() == 0)
				{
					return false;
				}

				return ReportErrorEventLogEntry(sourcename, 0, 0, 1, text.c_str());
			}

			//
			// Function: ReportErrorEventLogEntry
			//
			static bool ReportErrorEventLogEntry(const wstring sourcename, short category, long eventid, short numberofstrings, const wchar_t* description)
			{
				if (sourcename.size() == 0 || (numberofstrings > 0 && wcslen(description) == 0))
				{
					return false;
				}

				HANDLE hEventLog = ::RegisterEventSource(NULL, sourcename.c_str());
				if (hEventLog == NULL)
				{
					return false;
				}

				BOOL bStatus = ::ReportEvent(hEventLog, EVENTLOG_ERROR_TYPE, category, eventid, NULL, numberofstrings, 0, (LPCWSTR*)&description, NULL);
				if (!bStatus)
				{
					::DeregisterEventSource(hEventLog);
					return false;
				}

				::DeregisterEventSource(hEventLog);

				return true;
			}

			//
			// ReportWarningEventLogEntry #########################################################
			//

			//
			// Function: ReportWarningEventLogEntry
			//
			static bool ReportWarningEventLogEntry(const wstring sourcename, const wstring text)
			{
				if (sourcename.size() == 0)
				{
					return false;
				}

				return ReportWarningEventLogEntry(sourcename, 0, 0, 1, text.c_str());
			}

			//
			// Function: ReportWarningEventLogEntry
			//
			static bool ReportWarningEventLogEntry(const wstring sourcename, short category, long eventid, short numberofstrings, const wchar_t* description)
			{
				if (sourcename.size() == 0 || (numberofstrings > 0 && wcslen(description) == 0))
				{
					return false;
				}

				HANDLE hEventLog = ::RegisterEventSource(NULL, sourcename.c_str());
				if (hEventLog == NULL)
				{
					return false;
				}

				BOOL bStatus = ::ReportEvent(hEventLog, EVENTLOG_WARNING_TYPE, category, eventid, NULL, numberofstrings, 0, (LPCWSTR*)&description, NULL);
				if (!bStatus)
				{
					::DeregisterEventSource(hEventLog);
					return false;
				}

				::DeregisterEventSource(hEventLog);

				return true;
			}

			//
			// ReportInformationEventLogEntry #########################################################
			//

			//
			// Function: ReportInformationEventLogEntry
			//
			static bool ReportInformationEventLogEntry(const wstring sourcename, const wstring text)
			{
				if (sourcename.size() == 0)
				{
					return false;
				}

				return ReportInformationEventLogEntry(sourcename, 0, 0, 1, text.c_str());
			}

			//
			// Function: ReportInformationEventLogEntry
			//
			static bool ReportInformationEventLogEntry(const wstring sourcename, short category, long eventid, short numberofstrings, const wchar_t* description)
			{
				if (sourcename.size() == 0 || (numberofstrings > 0 && wcslen(description) == 0))
				{
					return false;
				}

				HANDLE hEventLog = ::RegisterEventSource(NULL, sourcename.c_str());
				if (hEventLog == NULL)
				{
					return false;
				}

				BOOL bStatus = ::ReportEvent(hEventLog, EVENTLOG_INFORMATION_TYPE, category, eventid, NULL, numberofstrings, 0, (LPCWSTR*)&description, NULL);
				if (!bStatus)
				{
					::DeregisterEventSource(hEventLog);
					return false;
				}

				::DeregisterEventSource(hEventLog);

				return true;
			}

			//
			// ReportAuditSuccessEventLogEntry #########################################################
			//

			//
			// Function: ReportAuditSuccessEventLogEntry
			//
			static bool ReportAuditSuccessEventLogEntry(const wstring sourcename, const wstring text)
			{
				if (sourcename.size() == 0)
				{
					return false;
				}

				return ReportAuditSuccessEventLogEntry(sourcename, 0, 0, 1, text.c_str());
			}

			//
			// Function: ReportAuditSuccessEventLogEntry
			//
			static bool ReportAuditSuccessEventLogEntry(const wstring sourcename, short category, long eventid, short numberofstrings, const wchar_t* description)
			{
				if (sourcename.size() == 0 || (numberofstrings > 0 && wcslen(description) == 0))
				{
					return false;
				}

				HANDLE hEventLog = ::RegisterEventSource(NULL, sourcename.c_str());
				if (hEventLog == NULL)
				{
					return false;
				}

				BOOL bStatus = ::ReportEvent(hEventLog, EVENTLOG_AUDIT_SUCCESS, category, eventid, NULL, numberofstrings, 0, (LPCWSTR*)&description, NULL);
				if (!bStatus)
				{
					::DeregisterEventSource(hEventLog);
					return false;
				}

				::DeregisterEventSource(hEventLog);

				return true;
			}

			//
			// ReportAuditFailureEventLogEntry #########################################################
			//

			//
			// Function: ReportAuditFailureEventLogEntry
			//
			static bool ReportAuditFailureEventLogEntry(const wstring sourcename, const wstring text)
			{
				if (sourcename.size() == 0)
				{
					return false;
				}

				return ReportAuditFailureEventLogEntry(sourcename, 0, 0, 1, text.c_str());
			}

			//
			// Function: ReportAuditFailureEventLogEntry
			//
			static bool ReportAuditFailureEventLogEntry(const wstring sourcename, short category, long eventid, short numberofstrings, const wchar_t* description)
			{
				if (sourcename.size() == 0 || (numberofstrings > 0 && wcslen(description) == 0))
				{
					return false;
				}

				HANDLE hEventLog = ::RegisterEventSource(NULL, sourcename.c_str());
				if (hEventLog == NULL)
				{
					return false;
				}

				BOOL bStatus = ::ReportEvent(hEventLog, EVENTLOG_AUDIT_FAILURE, category, eventid, NULL, numberofstrings, 0, (LPCWSTR*)&description, NULL);
				if (!bStatus)
				{
					::DeregisterEventSource(hEventLog);
					return false;
				}

				::DeregisterEventSource(hEventLog);

				return true;
			}

			//
			// Function: FromBase64String
			//
			static vector<BYTE> FromBase64String(wstring str)
			{
				DWORD cbData(0);
				DWORD dwFlags(0);

				if (str.size() == 0)
				{
					return vector<BYTE>();
				}

				if (!CryptStringToBinary(str.c_str(), (DWORD)str.size(), CRYPT_STRING_BASE64, NULL, &cbData, NULL, &dwFlags))
				{
					throw ItsException(L"CryptStringToBinary failed");
				}

				BYTE* pReturnValue = new BYTE[cbData];
				if (!CryptStringToBinary(str.c_str(), (DWORD)str.size(), CRYPT_STRING_BASE64, pReturnValue, &cbData, NULL, &dwFlags))
				{
					throw ItsException(L"CryptStringToBinary failed");
				}

				auto retVal = vector<BYTE>(cbData);
				for (DWORD i = 0; i < cbData; i++) {
					retVal[i] = pReturnValue[i];
				}

				delete[] pReturnValue;

				return retVal;
			}

			//
			// Function: ToBase64String
			//
			static wstring ToBase64String(vector<BYTE> data)
			{
				DWORD dwLength(static_cast<DWORD>(data.size()));
				DWORD cchString(0);
				CryptBinaryToString(data.data(), (DWORD)data.size(), CRYPT_STRING_BASE64, NULL, &cchString);
				unique_ptr<wchar_t[]> buffer = make_unique<wchar_t[]>(cchString + 1);
				BOOL bResult = CryptBinaryToString(data.data(), (DWORD)data.size(), CRYPT_STRING_BASE64, buffer.get(), &cchString);
				buffer[cchString] = L'\0';

				return wstring(buffer.get());
			}

			//
			// Function: FromBase64StringToString
			//
			static wstring FromBase64StringToString(wstring str)
			{
				DWORD cbData(0);
				DWORD dwFlags(0);

				if (str.size() == 0)
				{
					return false;
				}

				if (!CryptStringToBinary(str.c_str(), (DWORD)str.size(), CRYPT_STRING_BASE64, NULL, &cbData, NULL, &dwFlags))
				{
					throw ItsException(L"CryptStringToBinary failed");
				}

				BYTE* pReturnValue = new BYTE[cbData + sizeof(wchar_t)];
				if (!CryptStringToBinary(str.c_str(), (DWORD)str.size(), CRYPT_STRING_BASE64, pReturnValue, &cbData, NULL, &dwFlags))
				{
					throw ItsException(L"CryptStringToBinary failed");
				}

				pReturnValue[cbData] = 0;
				pReturnValue[cbData + 1] = 0;

				auto retVal = wstring(reinterpret_cast<wchar_t*>(pReturnValue));

				delete[] pReturnValue;

				return retVal;
			}

			//
			// Function: ToBase64StringFromString
			//
			static wstring ToBase64StringFromString(wstring str)
			{
				if (str.size() == 0) {
					return L"";
				}

				BYTE* pData = reinterpret_cast<BYTE*>(const_cast<wchar_t*>(str.c_str()));
				size_t dwSize = str.size() * sizeof(wchar_t);

				vector<BYTE> buffer;
				for (size_t i = 0; i < dwSize; i++) {
					buffer.push_back(pData[i]);
				}

				return ToBase64String(buffer);
			}
		};

		enum EREGCLASS
		{
			EHKEY_CLASSES_ROOT = 0,
			EHKEY_CURRENT_USER = 1,
			EHKEY_LOCAL_MACHINE = 2,
			EHKEY_USERS = 3,
			EHKEY_PERFORMANCE_DATA = 4,
			EHKEY_CURRENT_CONFIG = 5,
			EHKEY_DYN_DATA = 6
		};

		//
		// class: ItsRegistry
		//
		class ItsRegistry
		{
		private:
			//
			// Function: CheckHKey
			//
			// (i)	Checks to see wheather or not a valid REGCLASS value has been
			//		passed. If so the function returnes true, else it returnes false.
			//
			static bool CheckHKey(EREGCLASS a_regclass, HKEY &a_hKey)
			{
				switch (a_regclass)
				{
				case EREGCLASS(0):
					a_hKey = HKEY_CLASSES_ROOT;
					break;
				case EREGCLASS(1):
					a_hKey = HKEY_CURRENT_USER;
					break;
				case EREGCLASS(2):
					a_hKey = HKEY_LOCAL_MACHINE;
					break;
				case EREGCLASS(3):
					a_hKey = HKEY_USERS;
					break;
				case EREGCLASS(4):
					a_hKey = HKEY_PERFORMANCE_DATA;
					break;
				case EREGCLASS(5):
					a_hKey = HKEY_CURRENT_CONFIG;
					break;
				case EREGCLASS(6):
					a_hKey = HKEY_DYN_DATA;
					break;
				default:
					return false;
					break;
				};

				return true;
			}// bool CSettings::CheckHKey( long a_lhKey, HKEY &a_hKey )
		public:
			//
			// Function: EnumerateKeys
			//
			static bool EnumerateKeys(EREGCLASS eregclass, const wstring path, vector<wstring>* result)
			{
				HKEY hKey;
				if (!CheckHKey(eregclass, hKey) || path.size() == 0 || result == nullptr)
				{
					return false;
				}

				HKEY hOpenKey;
				LONG lResult(0);
				lResult = RegOpenKeyEx(hKey, path.c_str(), 0, KEY_ENUMERATE_SUB_KEYS, &hOpenKey);
				if (lResult != ERROR_SUCCESS)
				{
					return false;
				}

				result->clear();

				long lIndex(0);
				ULONG cbAvailable(0);
				wchar_t *pwcsName = new wchar_t[512];

				FILETIME ft;
				do
				{
					cbAvailable = 512;
					lResult = RegEnumKeyEx(hOpenKey, lIndex++, pwcsName, &cbAvailable, 0, NULL, 0, &ft);

					if (lResult == ERROR_SUCCESS)
					{
						result->push_back(pwcsName);
					}// if ( lResult == ERROR_SUCCESS ) {			

				} while (lResult == ERROR_SUCCESS);

				delete[] pwcsName;

				if (lResult != ERROR_NO_MORE_ITEMS)
				{
					RegCloseKey(hOpenKey);
					return false;
				}

				RegCloseKey(hOpenKey);

				return true;
			}

			//
			// Function: EnumerateValues
			//
			static bool EnumerateValues(EREGCLASS eregclass, wstring path, const wstring key, vector<wstring>* result)
			{
				HKEY hKey;
				if (!CheckHKey(eregclass, hKey) || (path.size() == 0 && key.size() == 0) || result == nullptr)
				{
					return false;
				}

				if (path.size() > 0 && path[path.size() - 1] != L'\\')
				{
					path += L'\\';
				}
				path += key;

				HKEY hOpenKey;
				LONG lResult(0);
				lResult = RegOpenKeyEx(hKey, path.c_str(), 0, KEY_QUERY_VALUE, &hOpenKey);
				if (lResult != ERROR_SUCCESS)
				{
					return false;
				}

				result->clear();

				long lIndex(0);
				ULONG cbAvailable(0);
				DWORD dwType(0);
				wchar_t *pwcsName = new wchar_t[512];

				do
				{
					cbAvailable = 512;
					lResult = RegEnumValue(hOpenKey, lIndex++, pwcsName, &cbAvailable, NULL, NULL, NULL, NULL);
					if (lResult == ERROR_SUCCESS)
					{
						result->push_back(pwcsName);
					}// if ( lResult == ERROR_SUCCESS ) {			
				} while (lResult == ERROR_SUCCESS);

				delete[] pwcsName;

				if (lResult != ERROR_NO_MORE_ITEMS)
				{
					return false;
				}

				RegCloseKey(hOpenKey);

				return true;
			}

			//
			// Function: ReadValue
			//
			static bool ReadValue(EREGCLASS eregclass, wstring path, wstring key, wstring value, wstring *data)
			{
				HKEY hKey;
				if (!CheckHKey(eregclass, hKey) || (path.size() == 0 && key.size() == 0) || data == nullptr)
				{
					return false;
				}

				if (path.size() > 0 && path[path.size() - 1] != L'\\')
				{
					path += L'\\';
				}
				path += key;

				HKEY hOpenKey;
				LONG lResult(0);
				lResult = RegOpenKeyEx(hKey, path.c_str(), 0, KEY_READ, &hOpenKey);
				if (lResult != ERROR_SUCCESS)
				{
					return false;
				}

				DWORD dwType(0);
				ULONG lSize(0);
				lResult = RegQueryValueEx(hOpenKey, value.c_str(), 0, &dwType, NULL, &lSize);
				if (lResult != ERROR_SUCCESS)
				{
					RegCloseKey(hOpenKey);
					return false;
				}

				if (dwType != REG_SZ)
				{
					RegCloseKey(hOpenKey);
					return false;
				}

				wchar_t *pszBuffer = new wchar_t[lSize / sizeof(wchar_t)];
				lResult = RegQueryValueEx(hOpenKey, value.c_str(), 0, &dwType, (BYTE *)pszBuffer, &lSize);
				if (lResult != ERROR_SUCCESS)
				{
					RegCloseKey(hOpenKey);
					delete[] pszBuffer;
					return false;
				}

				RegCloseKey(hOpenKey);

				*data = wstring(pszBuffer);

				return true;
			}

			//
			// Function: DeleteKey
			//
			static bool DeleteKey(EREGCLASS eregclass, wstring path, wstring key)
			{
				HKEY hKey;
				if (!CheckHKey(eregclass, hKey) || (path.size() == 0 && key.size() == 0))
				{
					return false;
				}

				if (path.size() > 0 && path[path.size() - 1] != L'\\')
				{
					path += L'\\';
				}
				path += key;

				DWORD dwResult(0);
				dwResult = ::SHDeleteKey(hKey, path.c_str());
				if (dwResult != ERROR_SUCCESS) {
					if (GetLastError() != 0) {
						return false;
					}
					else {
						return false;
					}
				}

				return true;
			}

			//
			// Function: CreateKey
			//
			static bool CreateKey(EREGCLASS eregclass, wstring path, wstring key, wstring default_data)
			{
				HKEY hKey;
				if (!CheckHKey(eregclass, hKey) || (path.size() == 0 && key.size() == 0))
				{
					return false;
				}

				if (path.length() > 0) {
					wchar_t wch = path.at(0);
					while (wch == L'\\') {
						path.erase(0);
						if (path.length() > 0) {
							wch = path.at(0);
						}
						else {
							break;
						}
					}
				}

				if (path.length() > 0) {
					if (path.at(path.length() - 1) != L'\\') {
						path += L'\\';
					}
				}

				path += key;

				// Create new key
				HKEY hNewKey;
				LONG lResult = RegCreateKeyEx(hKey, path.c_str(), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hNewKey, NULL);
				if (lResult != ERROR_SUCCESS) {
					return false;
				}

				// Set default value.	
				if (default_data.size() > 0) {
					lResult = RegSetValueEx(hNewKey, NULL, 0, REG_SZ, (CONST BYTE*)default_data.data(), (static_cast<DWORD>(default_data.size()) * sizeof(wchar_t)) + 2);
					if (lResult != ERROR_SUCCESS) {
						RegCloseKey(hNewKey);	// Close New Key				
						return false;
					}
				}

				RegCloseKey(hNewKey);	// Close New Key

				return true;
			}

			//
			// Function: AddValue
			//
			static bool AddValue(EREGCLASS eregclass, wstring path, wstring key, wstring value, wstring data)
			{
				HKEY hKey;
				if (!CheckHKey(eregclass, hKey) || (path.size() == 0 && key.size() == 0) || value.size() == 0)
				{
					return false;
				}

				if (path.size() > 0 && path[path.size() - 1] != L'\\')
				{
					path += L'\\';
				}
				path += key;

				HKEY hOpenKey;
				LONG lResult(0);
				lResult = RegOpenKeyEx(hKey, path.c_str(), 0, KEY_ALL_ACCESS, &hOpenKey);
				if (lResult != ERROR_SUCCESS) {
					return false;
				}

				lResult = RegSetValueEx(hOpenKey, value.c_str(), 0, REG_SZ, reinterpret_cast<const BYTE *>(data.data()), (static_cast<DWORD>(data.size()) * sizeof(wchar_t)) + 2);
				if (lResult != ERROR_SUCCESS) {
					RegCloseKey(hOpenKey);
					return false;
				}

				RegCloseKey(hOpenKey);

				return true;
			}

			//
			// Function: DeleteValue
			//
			static bool DeleteValue(EREGCLASS eregclass, wstring path, wstring key, wstring value)
			{
				HKEY hKey;
				if (!CheckHKey(eregclass, hKey) || (path.size() == 0 && key.size() == 0) || value.size() == 0)
				{
					return false;
				}

				if (path.size() > 0 && path[path.size() - 1] != L'\\')
				{
					path += L'\\';
				}
				path += key;

				HKEY hOpenKey;
				LONG lResult(0);
				lResult = RegOpenKeyEx(hKey, path.c_str(), 0, KEY_ALL_ACCESS, &hOpenKey);
				if (lResult != ERROR_SUCCESS) {
					return false;
				}
				lResult = RegDeleteValue(hOpenKey, value.c_str());
				if (lResult != ERROR_SUCCESS) {
					RegCloseKey(hOpenKey);
					return false;
				}

				RegCloseKey(hOpenKey);

				return true;
			}
		};// ItsRegistry
	}// namespace Win
}// namespace ItSoftware