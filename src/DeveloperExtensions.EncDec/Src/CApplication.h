#pragma once
//
// #include
//
#include <Windows.h>
#include <string>

//
// namespace
//
namespace DeveloperExtensions
{
	//
	// using
	//
	using std::wstring;

	//
	// class: CApplication
	//
	class CApplication
	{
	private:
		bool LoadFileList(wstring filename);
		wstring ExtractFilename(wstring full_path);
		wstring RenderFileOfCompletedText(size_t num_completed, size_t num_files);
		static void ThreadEncryptDecrypt(HWND hDlg, wstring crypt_key);
		static INT_PTR CALLBACK RunEncryptDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
		static void ProcThreadEncryptProgress(HWND hwndDlg, uint8_t* progress);
		static void DecryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size);
		static void EncryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size);
		static bool DoDecryptFile(wstring filename, wstring outFilename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress);
		static bool DoEncryptFile(wstring filename, wstring outFilename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress);
	protected:
	public:
		INT WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, INT nCmdShow);
		
		
	};
}