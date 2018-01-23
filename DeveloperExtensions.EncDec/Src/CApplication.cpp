//
// #include
//
#include <Windows.h>
#include <MsXml6.h>
#include <thread>
#include "../resource.h"
#include "CApplication.h"
#include "EAction.h"
#include <memory>
#include <vector>
#include <string>
#include "../../Include/itsoftware.h"
#include "../../Include/itsoftware-com.h"
#include "../../Include/itsoftware-win.h"
#include <sstream>
#include <Shlwapi.h>
#include <shellapi.h>

//
// extern
//
extern vector<wstring> g_errorlog;

//
// namespace
//
namespace DeveloperExtensions
{
	//
	// using
	//
	using std::shared_ptr;
	using std::vector;
	using std::wstring;
	using std::wstringstream;
	using std::unique_ptr;
	using std::thread;

	//
	// using namespace
	//
	using namespace ItSoftware;
	using namespace ItSoftware::COM;
	using namespace ItSoftware::Win;

	const int g_private_key_size = 64;
	static const uint8_t g_private_key[] =
	{ 0xc6, 0x5a, 0x23 ,0xd4, 0xbc, 0x76, 0x46, 0x24, 0xa9, 0xa7, 0xb0, 0x34, 0x4e, 0x5a, 0xe9, 0xf7,
		0xb0, 0x65, 0xa8, 0x10, 0x5c, 0x93, 0x46, 0x03, 0xba, 0x96, 0x0f, 0x88, 0xda, 0x9b, 0x64, 0x2b,
		0x28, 0x8a, 0x50, 0x2d, 0xd6, 0xe5, 0x4e, 0x98, 0xb8, 0x2f, 0x8e, 0x29, 0x91, 0xc6, 0xb9, 0x5a,
		0x61, 0x41, 0xca, 0xca, 0x26, 0xa1, 0x41, 0xd7, 0x80, 0xdc, 0x06, 0x8a, 0x23, 0xed, 0x2f, 0x06 };

	//
	// Global data
	//
	CApplication* g_pApplication;
	EAction g_action = EAction::Encrypt;
	vector<wstring> g_files;
	volatile bool g_quit = false;
	volatile bool g_endPB = false;
	bool g_by_file_association = false;
	bool g_delete_source_files = false;
	int g_wposx = -1;
	int g_wposy = -1;

	//
	// Method: wWinMain
	//
	// (i) Command line arguments are: action:<encrypt|decrypt> file:"<path+filename>"
	//
	INT WINAPI CApplication::wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, INT nCmdShow)
	{
		ComRuntime runtime{ ComApartment::SingleThreaded };
		
		g_pApplication = this;

		int argc = 0;
		LPWSTR* argv = CommandLineToArgvW(pCmdLine, &argc);
		

		if (argc == 1) 
		{
			//
			// We have a file argument. File association presumed through console or file explorer.
			//
			wstring argument1(argv[0]);
			wstring filename(argument1);
			if (!PathFileExists(filename.c_str()))
			{
				g_errorlog.push_back(L"Argument <file> does not exist!");
				return 1;
			}
			g_files.push_back(filename);

			if (filename.rfind(L".devenc") != wstring::npos)
			{
				//
				// We have an encrypted file. Lets decrypt it.
				//
				g_action = EAction::Decrypt;
			}
			else if (filename.rfind(L".devdec") != wstring::npos)
			{
				//
				// We have a decrypted file. Lets encrypt it.
				//
				g_action = EAction::Encrypt;
			}
			else 
			{
				g_errorlog.push_back(L"Argument <file> is not a known filetype!");
				return 1;
			}

			g_by_file_association = true;
		}
		else if ( argc >= 2 )
		{
			//
			// First argument. <action>.
			//
			wstring argument1(argv[0]);

			if (argument1 == L"action:encrypt")
			{
				g_action = EAction::Encrypt;
			}
			else if (argument1 == L"action:decrypt")
			{
				g_action = EAction::Decrypt;
			}
			else
			{
				g_errorlog.push_back(L"Argument <action> is invalid!");
				return 1;
			}
			
			//
			// Second argument. "<filename>"
			//			
			wstring argument2(argv[1]);
			wstring filename(argument2.begin() + 5, argument2.end());
			if (!PathFileExists(filename.c_str()))
			{
				g_errorlog.push_back(L"Argument <file> does not exist!");
				return 1;
			}

			if (!this->LoadFileList(filename))
			{
				g_errorlog.push_back(L"Argument <file> cannot load!");
				return 1;
			}

			if (argc >= 4) 
			{
				//
				// Thrid argument. <wposx>
				//				
				wstring argument3(argv[2]);
				wstring wposx(argument3.begin() + 6, argument3.end());
				g_wposx = ItsConvert::ToInt(wposx);

				//
				// Fourth argument. <wposy>
				//
				wstring argument4(argv[3]);
				wstring wposy(argument4.begin() + 6, argument4.end());
				g_wposy = ItsConvert::ToInt(wposy);
			}			
		}

		DialogBox(hInstance, MAKEINTRESOURCE(IDD_ENCDEC), GetActiveWindow(), RunEncryptDialogProc);
		
		if (g_errorlog.size() > 0) 
		{
			return 1;
		}

		return 0;
	}

	wstring CApplication::RenderFileOfCompletedText(size_t num_completed, size_t num_files)
	{
		wstringstream ss;
		ss << ItsConvert::ToStringFormatted(num_completed) << L" of " << ItsConvert::ToStringFormatted(num_files) << L" file(s) completed";		

		wstring str = ss.str();
		
		return str;
	}

	//
	// Method: RunEncryptDialogProc
	//
	INT_PTR CALLBACK CApplication::RunEncryptDialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		static unique_ptr<thread> t1;
		static unique_hicon_handle hIconSmall;
		static unique_hicon_handle hIconBig;
		static bool bStarted = false;
		switch (uMsg)
		{
		case WM_INITDIALOG:
		{			
			if (g_by_file_association) 
			{
				HWND hwndCB = GetDlgItem(hDlg, IDC_DLG_ENCDEC_REPLACE_EXISTING_FILE);
				SendMessage(hwndCB, BM_SETCHECK, BST_CHECKED, 0);
				EnableWindow(hwndCB, FALSE);
			}

			hIconSmall = (HICON)LoadImageW(GetModuleHandleW(NULL),
				MAKEINTRESOURCEW(IDI_ICON_MAIN),
				IMAGE_ICON,
				32,
				32,
				0);

			hIconBig = (HICON)LoadImageW(GetModuleHandleW(NULL),
				MAKEINTRESOURCEW(IDI_ICON_MAIN),
				IMAGE_ICON,				
				256,
				256,
				0);

			if (hIconSmall && hIconBig)
			{
				SendMessage(hDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIconSmall.operator HICON());
				SendMessage(hDlg, WM_SETICON, ICON_BIG, (LPARAM)hIconBig.operator HICON());
			}

			if (g_action == EAction::Encrypt) 
			{
				SendMessage(hDlg, WM_SETTEXT, 0, (LPARAM)L"Developer Extensions - Encryption");
			}
			else if (g_action == EAction::Decrypt) 
			{
				SendMessage(hDlg, WM_SETTEXT, 0, (LPARAM)L"Developer Extensions - Decryption");
			}
			
			SendMessage(GetDlgItem(hDlg, IDC_DLG_ENCDEC_FILE_PROGRESS), WM_SETTEXT, 0, (LPARAM)g_pApplication->RenderFileOfCompletedText(0, g_files.size()).c_str());
			SendMessage(GetDlgItem(hDlg, IDC_DLG_ENCDEC_KEY), EM_SETLIMITTEXT, 64, 0);
			if (g_files.size() > 1)
			{
				SetDlgItemTextW(hDlg, IDC_DLG_ENCDEC_FILENAME, L"multiple files");
			}
			else
			{
				wstring filename = g_pApplication->ExtractFilename(g_files.at(0));
				SetDlgItemTextW(hDlg, IDC_DLG_ENCDEC_FILENAME, filename.c_str());
			}
			
			//SetFocus(GetDlgItem(hDlg, IDC_DLG_ENCDEC_KEY));

			if (g_wposx != -1 && g_wposy != -1)
			{
				RECT rect{ 0 };
				GetWindowRect(hDlg, &rect);

				int offsetx = static_cast<int>((rect.right - rect.left) / 2);
				int offsety = static_cast<int>((rect.bottom - rect.top) / 2);

				MoveWindow(hDlg, g_wposx - offsetx, g_wposy - offsety, rect.right - rect.left, rect.bottom - rect.top, TRUE);
			}
			else 
			{
				RECT rectDlg{ 0 };
				GetWindowRect(hDlg, &rectDlg);

				RECT rectShell{ 0 };
				GetWindowRect(GetShellWindow(), &rectShell);

				int shellWidth = static_cast<int>((rectShell.right - rectShell.left)/2);
				int shellHeight = static_cast<int>((rectShell.bottom - rectShell.top)/2);

				int offsetx = static_cast<int>((rectDlg.right - rectDlg.left) / 2);
				int offsety = static_cast<int>((rectDlg.bottom - rectDlg.top) / 2);

				MoveWindow(hDlg, shellWidth - offsetx, shellHeight - offsety, rectDlg.right - rectDlg.left, rectDlg.bottom - rectDlg.top, TRUE);
			}			
			return TRUE;
		}
		case WM_COMMAND:
			switch (LOWORD(wParam))
			{
			case IDOK:
			{
				wchar_t key[200];
				GetDlgItemText(hDlg, IDC_DLG_ENCDEC_KEY, key, 200);
				
				if (wcslen(key) == 0)
				{									
					ItsWin::ShowInformationMessage(hDlg, L"Developer Extensions", L"Key must contain between 1 and 64 characters." );
					return FALSE;
				}

				LRESULT lResult = SendMessage(GetDlgItem(hDlg, IDC_DLG_ENCDEC_DELETE_SOURCE_FILES), BM_GETCHECK, 0, 0);
				if (lResult == BST_CHECKED)
				{
					g_delete_source_files = true;
				}
				
				EnableWindow(GetDlgItem(hDlg, IDC_DLG_ENCDEC_REPLACE_EXISTING_FILE), FALSE);
				EnableWindow(GetDlgItem(hDlg, IDC_DLG_ENCDEC_DELETE_SOURCE_FILES), FALSE);
				EnableWindow(GetDlgItem(hDlg, IDOK), FALSE);
				EnableWindow(GetDlgItem(hDlg, IDCANCEL), FALSE);
				
				t1 = make_unique<thread>(ThreadEncryptDecrypt, hDlg, key);

				bStarted = true;

				return TRUE;
			}
			case IDCANCEL:
			{
				EnableWindow(GetDlgItem(hDlg, IDCANCEL), FALSE);
				g_quit = true;
				if ( t1 != nullptr && t1->joinable())
				{
					t1->join();

					wstringstream ss;
					if (g_action == EAction::Encrypt) 
					{
						ss << L"Encryption is complete." << endl;
					}
					else if (g_action == EAction::Decrypt) 
					{
						ss << L"Decryption is complete." << endl;
					}

					if (g_errorlog.size() == 0) 
					{
						ss << L"No errors during operations.";
					}
					else 
					{
						ss << g_errorlog.size() << L" errors during operations.";
					}

					wstring msg = ss.str();
					ItsWin::ShowInformationMessage(hDlg, L"Developer Extensions", msg.c_str());
				}																
				
				EndDialog(hDlg, 0);
				return TRUE;
			}
			case IDC_DLG_ENCDEC_KEY_CHECK:
			{
				LRESULT lr = SendMessage(GetDlgItem(hDlg,IDC_DLG_ENCDEC_KEY_CHECK), BM_GETCHECK, 0, 0);

				HWND hwndKey = GetDlgItem(hDlg, IDC_DLG_ENCDEC_KEY); 

				LONG_PTR style = GetWindowLongPtr(hwndKey, GWL_STYLE);				

				if (lr == BST_CHECKED)
				{
					style -= ES_PASSWORD;
					SetWindowLongPtr(hwndKey, GWL_STYLE, style);
					SendMessage(hwndKey, EM_SETPASSWORDCHAR, 0, 0);
				}
				else
				{
					style |= ES_PASSWORD;
					SetWindowLongPtr(hwndKey, GWL_STYLE, style);
					SendMessage(hwndKey, EM_SETPASSWORDCHAR, (WPARAM)L'*', 0);
				}
				
				SetFocus(hwndKey);

				return TRUE;
			}
			default:
				break;
			}
			break;

		default:
			break;
		}

		return FALSE;
	}// RunEncryptDialogProc

	void CApplication::ThreadEncryptDecrypt(HWND hDlg, wstring crypt_key)
	{
		ComRuntime runtime(ComApartment::SingleThreaded);

		LRESULT lr = SendMessage(GetDlgItem(hDlg, IDC_DLG_ENCDEC_REPLACE_EXISTING_FILE), BM_GETCHECK, 0, 0);
		bool replace_existing_file = false;
		if (lr == BST_CHECKED) {
			replace_existing_file = true;
		}

		size_t length = g_files.size();
		size_t l = 0;

		uint8_t progress = 0;		

		thread t2{ ProcThreadEncryptProgress, hDlg, &progress };
		while (!g_quit && l < length)
		{			
			wstring filename = g_pApplication->ExtractFilename(g_files.at(l));
			SetDlgItemTextW(hDlg, IDC_DLG_ENCDEC_FILENAME, filename.c_str());		

			//SendMessage(GetDlgItem(hDlg, IDC_DLG_ENCDEC_FILE_PROGRESS), WM_SETTEXT, 0, (LPARAM)L"0%");			

			if (g_action == EAction::Encrypt)
			{				
				wstring filename = g_files.at(l);
				filename += L".devenc.tmp";
				DoEncryptFile(g_files.at(l),filename, g_private_key, g_private_key_size, reinterpret_cast<uint8_t*>(const_cast<wchar_t*>(crypt_key.data())), static_cast<int>(crypt_key.size()) * sizeof(wchar_t), &progress);

				if (g_quit) 
				{
					if (!DeleteFile(filename.c_str())) 
					{
						g_errorlog.push_back(L"Error deleting temporary encryption file!");
					}
				}
				else 
				{
					if (replace_existing_file) 
					{
						if (g_by_file_association)
						{
							wstring fname(g_files.at(l).begin(), g_files.at(l).end() - 7);
							if (!MoveFileEx(filename.c_str(), fname.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to source file!");
							}
						}
						else 
						{
							if (!MoveFileEx(filename.c_str(), g_files.at(l).c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to source file!");
							}
						}
					}
					else 
					{
						if (g_by_file_association)
						{
							wstring fname(filename.begin(), filename.end() - 7);
							if (!MoveFileEx(filename.c_str(), fname.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to devenc file!");
							}
						}
						else
						{
							wstring new_filename = g_files.at(l);
							new_filename += L".devenc";
							if (!MoveFileEx(filename.c_str(), new_filename.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to devenc file!");
							}
						}
					}
				}				
			}
			else if (g_action == EAction::Decrypt)
			{
				wstring filename = g_files.at(l);
				filename += L".devdec.tmp";
				DoDecryptFile(g_files.at(l),filename, g_private_key, g_private_key_size, reinterpret_cast<uint8_t*>(const_cast<wchar_t*>(crypt_key.data())), static_cast<int>(crypt_key.size()) * sizeof(wchar_t), &progress);

				if (g_quit) 
				{
					if (!DeleteFile(filename.c_str())) 
					{
						g_errorlog.push_back(L"Error deleting temporary decryption file!");
					}
				}
				else 
				{
					if (replace_existing_file)
					{
						if (g_by_file_association)
						{
							wstring fname(g_files.at(l).begin(), g_files.at(l).end() - 7);
							if (!MoveFileEx(filename.c_str(), fname.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to source file!");
							}
						}
						else
						{
							if (!MoveFileEx(filename.c_str(), g_files.at(l).c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to source file!");
							}
						}
					}
					else 
					{
						if (g_by_file_association)
						{
							wstring fname(filename.begin(), filename.end() - 7);
							if (!MoveFileEx(filename.c_str(), fname.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to devenc file!");
							}
						}
						else
						{
							wstring new_filename = g_files.at(l);
							new_filename += L".devdec";
							if (!MoveFileEx(filename.c_str(), new_filename.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
							{
								g_errorlog.push_back(L"Error moving and replacing temporary to devdec file!");
							}
						}
					}
				}
			}

			if (g_delete_source_files) {
				if (!DeleteFile(g_files.at(l).c_str()))
				{
					g_errorlog.push_back(L"Error deleting source file!");
				}
			}


			if (!g_quit) {
				SetDlgItemTextW(hDlg, IDC_DLG_ENCDEC_FILE_PROGRESS, g_pApplication->RenderFileOfCompletedText(l + 1, length).c_str());
			}

			l++;
		}

		g_endPB = true;
		t2.join();
		
		if (!g_quit) {
			PostMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
		}
	}

	void CApplication::ProcThreadEncryptProgress(HWND hDlg, uint8_t* progress)
	{		
		EnableWindow(GetDlgItem(hDlg, IDCANCEL), TRUE);
		HWND hwndPB = GetDlgItem(hDlg, IDC_DLG_ENCDEC_PROGRESS);
		HWND hwndPS = GetDlgItem(hDlg, IDC_DLG_ENCDEC_PROGRESS_STATIC);
		SendMessage(hwndPB, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
		SendMessage(hwndPB, PBM_SETSTEP, (WPARAM)1, 0);

		HWND hwndFileP = GetDlgItem(hDlg, IDC_DLG_ENCDEC_FILE_PROGRESS);
		
		int prev = 0;
		do
		{
			if (*progress != prev)
			{
				SendMessage(hwndPB, PBM_SETPOS, *progress, 0);
				prev = *progress;
				wstring progress_text = L"Progress ";
				progress_text += ItsConvert::ToString(*progress);
				progress_text += L"%";
				SendMessage(hwndPS, WM_SETTEXT, 0, (LPARAM)progress_text.c_str());
				Sleep(100);
			}				
		} while (!g_endPB && !g_quit);

		if (!g_quit) {
			SendMessage(hwndPB, PBM_SETPOS, 100, 0);
		}
	}

	//
	//
	//
	bool CApplication::LoadFileList(wstring filename)
	{
		CComPtr<IXMLDOMDocument2> pIXMLDOMDocument;
		HRESULT hr = pIXMLDOMDocument.CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX::CLSCTX_INPROC_SERVER);
		if (FAILED(hr))
		{
			g_errorlog.push_back(L"Could not create DOM Document!");
			g_errorlog.push_back(ItsError::GetCoLastErrorInfoDescription());
			return false;
		}

		VARIANT_BOOL vbStatus;
		hr = pIXMLDOMDocument->load(CComVariant(filename.c_str()), &vbStatus);
		if (FAILED(hr) || vbStatus == VARIANT_FALSE)
		{
			g_errorlog.push_back(L"Could not load file list!");
			g_errorlog.push_back(ItsError::GetCoLastErrorInfoDescription());
			return false;
		}

		CComPtr<IXMLDOMNodeList> pIXDNL;
		hr = pIXMLDOMDocument->selectNodes(CComBSTR(L"/files/file"), &pIXDNL);

		long length(0);
		hr = pIXDNL->get_length(&length);

		if (length == 0) {
			return false;
		}

		for (long l = 0; l < length; l++)
		{
			CComPtr<IXMLDOMNode> listitem;
			hr = pIXDNL->get_item(l, &listitem);

			CComBSTR text;
			hr = listitem->get_text(&text);

			g_files.push_back(text.operator LPWSTR());
		}

		return true;
	}

	wstring CApplication::ExtractFilename(wstring full_path)
	{
		size_t lLength = full_path.size();
		long lLastHit(-1);
		for (long l = 0; l < lLength; l++) {
			if (full_path[l] == L'\\' || full_path[l] == L'/') {
				lLastHit = l;
			}
		}

		wstringstream stream;
		if (lLastHit != -1) {
			for (long l = lLastHit + 1; l < lLength; l++) {
				stream << full_path[l];
			}
		}
		else {
			stream << full_path;
		}

		wstring retVal = stream.str();

		return retVal;
	}


	//
	// Function: DoEncryptFile
	//
	bool CApplication::DoEncryptFile(wstring filename, wstring outFilename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress)
	{
		unique_handle_handle hFile(CreateFile(filename.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL));
		unique_handle_handle hFileOut(CreateFile(outFilename.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));

		LARGE_INTEGER fileSize{ 0 };
		GetFileSizeEx(hFile, &fileSize);

		if (fileSize.QuadPart == 0) {
			return true;
		}

		LARGE_INTEGER readSize{ 0 };
		LARGE_INTEGER writtenSize{ 0 };
		const int bufferSize = 8 * 1024 * 1024;
		unique_ptr<uint8_t[]> buffer = make_unique<uint8_t[]>(bufferSize);

		*progress = 0;

		while (readSize.QuadPart + bufferSize < fileSize.QuadPart)
		{
			DWORD dwRead{ 0 };
			SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
			if (!ReadFile(hFile, buffer.get(), bufferSize, &dwRead, NULL)) {
				g_errorlog.push_back(L"Error reading file '" + filename + L"'!");
				g_errorlog.push_back(ItsError::GetLastErrorDescription());
				return false;
			}
			readSize.QuadPart += dwRead;

			EncryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

			DWORD dwWritten{ 0 };
			SetFilePointer(hFileOut, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);
			if (!WriteFile(hFileOut, buffer.get(), dwRead, &dwWritten, NULL))
			{
				g_errorlog.push_back(L"Error writing file '" + filename + L"'!");
				g_errorlog.push_back(ItsError::GetLastErrorDescription());
				return false;
			}
			writtenSize.QuadPart += dwWritten;

			*progress = (uint8_t)(writtenSize.QuadPart * 100 / fileSize.QuadPart);

			if (g_quit) {
				return false;
			}
		}

		DWORD dwRead{ 0 };
		SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
		if (!ReadFile(hFile, buffer.get(), static_cast<DWORD>(fileSize.QuadPart - readSize.QuadPart), &dwRead, NULL)) {
			g_errorlog.push_back(L"Error reading file '" + filename + L"'!");
			g_errorlog.push_back(ItsError::GetLastErrorDescription());
			return false;
		}
		readSize.QuadPart += dwRead;

		EncryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

		DWORD dwWritten{ 0 };
		SetFilePointer(hFileOut, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);
		if (!WriteFile(hFileOut, buffer.get(), dwRead, &dwWritten, NULL))
		{
			g_errorlog.push_back(L"Error writing file '" + filename + L"'!");
			g_errorlog.push_back(ItsError::GetLastErrorDescription());
			return false;
		}
		writtenSize.QuadPart += dwWritten;

		*progress = 100;

		return true;
	}


	//
	// Function: DoDecryptFile
	//
	bool CApplication::DoDecryptFile(wstring filename, wstring outFilename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress)
	{
		unique_handle_handle hFile(CreateFile(filename.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL));
		unique_handle_handle hFileOut(CreateFile(outFilename.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));

		LARGE_INTEGER fileSize{ 0 };
		GetFileSizeEx(hFile, &fileSize);

		if (fileSize.QuadPart == 0) {
			return true;
		}

		LARGE_INTEGER readSize{ 0 };
		LARGE_INTEGER writtenSize{ 0 };
		const int bufferSize = 8 * 1024 * 1024;
		unique_ptr<uint8_t[]> buffer = make_unique<uint8_t[]>(bufferSize);

		*progress = 0;

		while (readSize.QuadPart + bufferSize < fileSize.QuadPart)
		{
			DWORD dwRead{ 0 };
			SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
			if (!ReadFile(hFile, buffer.get(), bufferSize, &dwRead, NULL)) {
				g_errorlog.push_back(L"Error reading file '" + filename + L"'!");
				g_errorlog.push_back(ItsError::GetLastErrorDescription());
				return false;
			}
			readSize.QuadPart += dwRead;

			DecryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

			DWORD dwWritten{ 0 };
			SetFilePointer(hFileOut, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);
			if (!WriteFile(hFileOut, buffer.get(), dwRead, &dwWritten, NULL))
			{
				g_errorlog.push_back(L"Error writing file '" + filename + L"'!");
				g_errorlog.push_back(ItsError::GetLastErrorDescription());
				return false;
			}
			writtenSize.QuadPart += dwWritten;

			*progress = (uint8_t)(writtenSize.QuadPart * 100 / fileSize.QuadPart);

			if (g_quit) {
				return false;
			}
		}

		DWORD dwRead{ 0 };
		SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
		if (!ReadFile(hFile, buffer.get(), static_cast<DWORD>(fileSize.QuadPart - readSize.QuadPart), &dwRead, NULL)) {
			g_errorlog.push_back(L"Error reading file '" + filename + L"'!");
			g_errorlog.push_back(ItsError::GetLastErrorDescription());
			return false;
		}
		readSize.QuadPart += dwRead;

		DecryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

		DWORD dwWritten{ 0 };
		SetFilePointer(hFileOut, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);
		if (!WriteFile(hFileOut, buffer.get(), dwRead, &dwWritten, NULL))
		{
			g_errorlog.push_back(L"Error writing file '" + filename + L"'!");
			g_errorlog.push_back(ItsError::GetLastErrorDescription());
			return false;
		}
		writtenSize.QuadPart += dwWritten;

		*progress = 100;

		return true;
	}

	//
	// Function:: EncryptData
	//
	void CApplication::EncryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size)
	{
		bool bEven = true;
		for (unsigned int i = 0; i < size; i++)
		{
			for (int j = 0; j < private_key_size; j++)
			{
				data[i] ^= private_key[j];
				data[i] += i;
				data[i] = ROL(data[i], 3);
			}

			for (int k = 0; k < crypt_key_size; k++)
			{
				data[i] ^= crypt_key[k];
			}

			data[i] = ROL(data[i], 3);

			data[i] ^= 0x29;

			if (bEven)
			{
				const int key_size = 16;
				static const uint8_t trail_key[key_size] =
				{ 0xab, 0x6b, 0x0f, 0x82, 0x45, 0xa2, 0x41, 0x71, 0x93, 0xe4, 0x36, 0x88, 0x5a, 0x31, 0x2b, 0x1a };

				for (int p = 0; p < key_size; p++)
				{
					data[i] ^= trail_key[p];
				}

				data[i] = ROL(data[i], 3);

				bEven = false;
			}
			else
			{
				bEven = true;
			}
		}
	}

	//
	// Function:: DecryptData
	//
	void CApplication::DecryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size)
	{
		bool bEven = true;

		for (unsigned int i = 0; i < size; i++)
		{
			if (bEven)
			{
				data[i] = ROR(data[i], 3);

				const int key_size = 16;
				static const uint8_t trail_key[key_size] =
				{ 0xab, 0x6b, 0x0f, 0x82, 0x45, 0xa2, 0x41, 0x71, 0x93, 0xe4, 0x36, 0x88, 0x5a, 0x31, 0x2b, 0x1a };

				for (int p = key_size - 1; p >= 0; p--)
				{
					data[i] ^= trail_key[p];
				}

				bEven = false;
			}
			else
			{
				bEven = true;
			}

			data[i] ^= 0x29;

			data[i] = ROR(data[i], 3);

			for (int k = crypt_key_size - 1; k >= 0; k--)
			{
				data[i] ^= crypt_key[k];
			}

			for (int j = private_key_size - 1; j >= 0; j--)
			{
				data[i] = ROR(data[i], 3);
				data[i] -= i;
				data[i] ^= private_key[j];
			}
		}
	}
}// namespace