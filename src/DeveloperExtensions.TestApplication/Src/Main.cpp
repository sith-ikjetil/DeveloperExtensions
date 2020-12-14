//
// #include
//
#include <iostream>
#include <memory>
#include <string>
#include <MsXml6.h>
#include <Windows.h>
#include <atlbase.h>
#include <atlcom.h>
#include "../../Include/itsoftware.h"
#include "../../Include/itsoftware-com.h"
#include "../../Include/itsoftware-win.h"
#include "../../Include/itsoftware-cli.h"

//
// using
//
using std::unique_ptr;
using std::make_unique;
using std::wstring;
using std::wcout;
using std::endl;

//
// using namespace
//
using namespace ItSoftware;
using namespace ItSoftware::COM;
using namespace ItSoftware::CLI;
using namespace System;
using namespace System::Threading;

//
// #define
//
//#define EOP_BUFFER_SIZE 4096
#define CONSOLE_INPUT_BUFFER_SIZE 4096

//
// Function Prototypes
//
void PrintHeader();
bool PrintMenu();
void ExecuteEncryption();
void ExecuteDecryption();
String^ GetInputFilename();
String^ GetInputCryptKey();
void ProcessEndOfProgram();
bool DoEncryptFile(wstring filename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress);
bool DoDecryptFile(wstring filename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress);
void EncryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size);
void DecryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size);
wstring SaveTargetsForEncDec(wstring filename, vector<wstring>& files);
void ExecuteSaveTargetsForEncDec();

public ref class ThreadX
{
	uint8_t* progress;
public:
	ThreadX(uint8_t* progress)
	{
		this->progress = progress;
	}

	void ThreadEntryPoint()
	{
		Console::Write(L"Progress: [");

		int prev = 0;

		do
		{
			if (*progress > prev)
			{
				for (int i = (prev/10); i < (*progress / 10); i++) {
					Console::Write(L"#");
				}
				prev = *progress;
			}
			Sleep(250);
		} while (*progress < 100);

		Console::Write("#");

		if (prev == 0) 
		{
			Console::Write(L"#########");
		}

		Console::Write(L"]");
		Console::WriteLine();
	}
};

//
// Function: main
//
int wmain()
{
	//ComRuntime runtime(ComApartment::SingleThreaded);

	PrintHeader();

	while (!PrintMenu());

	ProcessEndOfProgram();	
	
	return 0;
}

//
// Function: PrintHeader
//
void PrintHeader()
{
	Console::WriteLine(L"## Developer Extension - Test Application ##");
}

//
// Function
//
bool PrintMenu()
{
	Console::WriteLine();
	Console::WriteLine(L"**********");
	Console::WriteLine(L"M E N U");
	Console::WriteLine(L" 0) :: Exit");
	Console::WriteLine(L" 1) :: Encrypt File");
	Console::WriteLine(L" 2) :: Decrypt File");
	Console::WriteLine(L" 3) :: SaveTargetsForEncDec");
	Console::WriteLine();
	Console::Write(L"::>");

	String^ choice = Console::ReadLine()->Trim();

	if (choice == L"0")
	{
		return true;
	}

	if (choice == L"1")
	{
		ExecuteEncryption();
		return false;
	}

	if (choice == L"2")
	{
		ExecuteDecryption();
		return false;
	}

	if (choice == L"3")
	{
		ExecuteSaveTargetsForEncDec();
		return false;
	}

	return false;
}

//
// Function: ExecuteEncryption
//
void ExecuteEncryption()
{
	//
	// Retrieve Filename
	//
	String^ filename = GetInputFilename();
	if (!System::IO::File::Exists(filename))
	{
		Console::WriteLine("#[Error]# File not found!");
		return;
	}

	//
	// Retrieve Input Crypt Key
	//
	String^ crypt_key = GetInputCryptKey();
	if (crypt_key->Length <= 0) {
		Console::WriteLine("#[Error]# Encryption key invalid!");
		return;
	}

	//
	// Crypt Key 
	//
	cli::array<Byte>^ crypt_key_as_bytes = System::Text::Encoding::Unicode->GetBytes(crypt_key);
	int _crypt_key_size = crypt_key_as_bytes->Length;
	unique_ptr<uint8_t[]> _crypt_key = make_unique<uint8_t[]>(_crypt_key_size);
	for (int i = 0; i < _crypt_key_size; i++)
	{
		_crypt_key.get()[i] = crypt_key_as_bytes[i];		
		if (_crypt_key.get()[i] == 0x00) {
			_crypt_key.get()[i] = 0x6F;
		}
	}

	//
	// Private Key
	//
	const int private_key_size = 64;
	static const uint8_t private_key[private_key_size] =
	{ 0xc6, 0x5a, 0x23 ,0xd4, 0xbc, 0x76, 0x46, 0x24, 0xa9, 0xa7, 0xb0, 0x34, 0x4e, 0x5a, 0xe9, 0xf7,
	  0xb0, 0x65, 0xa8, 0x10, 0x5c, 0x93, 0x46, 0x03, 0xba, 0x96, 0x0f, 0x88, 0xda, 0x9b, 0x64, 0x2b,
	  0x28, 0x8a, 0x50, 0x2d, 0xd6, 0xe5, 0x4e, 0x98, 0xb8, 0x2f, 0x8e, 0x29, 0x91, 0xc6, 0xb9, 0x5a,
	  0x61, 0x41, 0xca, 0xca, 0x26, 0xa1, 0x41, 0xd7, 0x80, 0xdc, 0x06, 0x8a, 0x23, 0xed, 0x2f, 0x06 };

	//
	// Prepare Arguments
	//
	wstring _filename(ItsCli::ToString(filename));

	//
	// Call Encryption Method
	//
	uint8_t progress = 0;
	ThreadX^ o1 = gcnew ThreadX(&progress);
	Thread^ t1 = gcnew Thread(gcnew ThreadStart(o1, &ThreadX::ThreadEntryPoint));
	t1->Start();
	if (!DoEncryptFile(_filename, private_key, private_key_size, _crypt_key.get(), _crypt_key_size, &progress))
	{
		Console::WriteLine(L"#[Error]# Encryption Failed!");
		return;
	}
	t1->Join();
	Console::WriteLine(L"#[Success]# Encryption OK!");
}

//
// Function: ExecuteEncryption
//
void ExecuteDecryption()
{
	//
	// Retrieve Filename
	//
	String^ filename = GetInputFilename();
	if (!System::IO::File::Exists(filename))
	{
		Console::WriteLine("#[Error]# File not found!");
		return;
	}

	//
	// Retrieve Input Crypt Key
	//
	String^ crypt_key = GetInputCryptKey();
	if (crypt_key->Length <= 0) {
		Console::WriteLine("#[Error]# Encryption key invalid!");
		return;
	}

	//
	// Crypt Key 
	//
	cli::array<Byte>^ crypt_key_as_bytes = System::Text::Encoding::Unicode->GetBytes(crypt_key);
	int _crypt_key_size = crypt_key_as_bytes->Length;
	unique_ptr<uint8_t[]> _crypt_key = make_unique<uint8_t[]>(_crypt_key_size);
	for (int i = 0; i < _crypt_key_size; i++)
	{
		_crypt_key.get()[i] = crypt_key_as_bytes[i];
		if (_crypt_key.get()[i] == 0x00) {
			_crypt_key.get()[i] = 0x6F;
		}
	}

	//
	// Private Key
	//
	const int private_key_size = 64;
	static const uint8_t private_key[private_key_size] =
	{ 0xc6, 0x5a, 0x23 ,0xd4, 0xbc, 0x76, 0x46, 0x24, 0xa9, 0xa7, 0xb0, 0x34, 0x4e, 0x5a, 0xe9, 0xf7,
		0xb0, 0x65, 0xa8, 0x10, 0x5c, 0x93, 0x46, 0x03, 0xba, 0x96, 0x0f, 0x88, 0xda, 0x9b, 0x64, 0x2b,
		0x28, 0x8a, 0x50, 0x2d, 0xd6, 0xe5, 0x4e, 0x98, 0xb8, 0x2f, 0x8e, 0x29, 0x91, 0xc6, 0xb9, 0x5a,
		0x61, 0x41, 0xca, 0xca, 0x26, 0xa1, 0x41, 0xd7, 0x80, 0xdc, 0x06, 0x8a, 0x23, 0xed, 0x2f, 0x06 };

	//
	// Prepare Arguments
	//
	wstring _filename(ItsCli::ToString(filename));

	//
	// Call Encryption Method
	//
	uint8_t progress = 0;	
	
	ThreadX^ o1 = gcnew ThreadX(&progress);
	Thread^ t1 = gcnew Thread(gcnew ThreadStart(o1, &ThreadX::ThreadEntryPoint));	
	t1->Start();
	if (!DoDecryptFile(_filename, private_key, private_key_size, _crypt_key.get(), _crypt_key_size, &progress)) {
		Console::WriteLine(L"#[Error]# Decryption Failed!");
		return;
	}
	t1->Join();

	Console::WriteLine(L"#[Success]# Decryption OK!");
}

//
// Function: GetInputFilename
//
String^ GetInputFilename()
{
	Console::WriteLine();
	Console::Write(L"Input Filename: ");
	return Console::ReadLine();
}

//
// Function: GetInputCryptKey
//
String^ GetInputCryptKey()
{
	Console::WriteLine();
	Console::Write(L"Encryption key: ");
	return Console::ReadLine();
}

//
// Function: ProcessEndOfProgram
//
void ProcessEndOfProgram()
{
	Console::WriteLine();
	Console::WriteLine(L"Bye bye!");

	//unique_ptr<wchar_t[]> buffer = make_unique<wchar_t[]>(EOP_BUFFER_SIZE);
	//wcin.getline(buffer.get(), 255);
}

//
// Function: DoEncryptFile
//
bool DoEncryptFile(wstring filename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress)
{	
	ItSoftware::Win::unique_handle_handle hFile(CreateFile(filename.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL));

	LARGE_INTEGER fileSize{ 0 };
	GetFileSizeEx(hFile, &fileSize);

	if (fileSize.QuadPart == 0) {
		return true;
	}

	LARGE_INTEGER readSize{ 0 };
	LARGE_INTEGER writtenSize{ 0 };
	const int bufferSize = 64*1024*1024;
	unique_ptr<uint8_t[]> buffer = make_unique<uint8_t[]>(bufferSize);	

	*progress = 0;

	while (readSize.QuadPart + bufferSize < fileSize.QuadPart) 
	{
		DWORD dwRead{ 0 };
		SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
		if (!ReadFile(hFile, buffer.get(), bufferSize, &dwRead, NULL)) {
			Console::WriteLine(L"#[Error]# Error reading file!");
			return false;
		}
		readSize.QuadPart += dwRead;

		EncryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

		DWORD dwWritten{ 0 };
		SetFilePointer(hFile, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);		
		if (!WriteFile(hFile, buffer.get(), dwRead, &dwWritten, NULL))
		{
			Console::WriteLine(L"#[Error]# Error writing file!");
			return false;
		}
		writtenSize.QuadPart += dwWritten;		

		*progress = (uint8_t)(writtenSize.QuadPart * 100 / fileSize.QuadPart);
	}

	DWORD dwRead{ 0 };
	SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
	if (!ReadFile(hFile, buffer.get(), static_cast<DWORD>(fileSize.QuadPart-readSize.QuadPart), &dwRead, NULL)) {
		Console::WriteLine(L"#[Error]# Error reading file!");
		return false;
	}
	readSize.QuadPart += dwRead;

	EncryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

	DWORD dwWritten{ 0 };
	SetFilePointer(hFile, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);
	if (!WriteFile(hFile, buffer.get(), dwRead, &dwWritten, NULL))
	{
		Console::WriteLine(L"#[Error]# Error writing file!");
		return false;
	}
	writtenSize.QuadPart += dwWritten;

	*progress = 100;

	return true;
}


//
// Function: DoDecryptFile
//
bool DoDecryptFile(wstring filename, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size, uint8_t* progress)
{
	ItSoftware::Win::unique_handle_handle hFile(CreateFile(filename.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL));

	LARGE_INTEGER fileSize{ 0 };
	GetFileSizeEx(hFile, &fileSize);

	if (fileSize.QuadPart == 0) {
		return true;
	}

	LARGE_INTEGER readSize{ 0 };
	LARGE_INTEGER writtenSize{ 0 };
	const int bufferSize = 64*1024*1024;
	unique_ptr<uint8_t[]> buffer = make_unique<uint8_t[]>(bufferSize);

	*progress = 0;

	while (readSize.QuadPart + bufferSize < fileSize.QuadPart)
	{
		DWORD dwRead{ 0 };
		SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
		if (!ReadFile(hFile, buffer.get(), bufferSize, &dwRead, NULL)) {
			Console::WriteLine(L"#[Error]# Error reading file!");
			return false;
		}
		readSize.QuadPart += dwRead;

		DecryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

		DWORD dwWritten{ 0 };
		SetFilePointer(hFile, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);
		if (!WriteFile(hFile, buffer.get(), dwRead, &dwWritten, NULL))
		{
			Console::WriteLine(L"#[Error]# Error writing file!");
			return false;
		}
		writtenSize.QuadPart += dwWritten;

		*progress = (uint8_t)(writtenSize.QuadPart * 100 / fileSize.QuadPart);
	}

	DWORD dwRead{ 0 };
	SetFilePointer(hFile, readSize.LowPart, &readSize.HighPart, FILE_BEGIN);
	if (!ReadFile(hFile, buffer.get(), static_cast<DWORD>(fileSize.QuadPart - readSize.QuadPart), &dwRead, NULL)) {
		Console::WriteLine(L"#[Error]# Error reading file!");
		return false;
	}
	readSize.QuadPart += dwRead;

	DecryptData(buffer.get(), dwRead, private_key, private_key_size, crypt_key, crypt_key_size);

	DWORD dwWritten{ 0 };
	SetFilePointer(hFile, writtenSize.LowPart, &writtenSize.HighPart, FILE_BEGIN);
	if (!WriteFile(hFile, buffer.get(), dwRead, &dwWritten, NULL))
	{
		Console::WriteLine(L"#[Error]# Error writing file!");
		return false;
	}
	writtenSize.QuadPart += dwWritten;	

	*progress = 100;

	return true;
}

//
// Function:: EncryptData
//
void EncryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size)
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
void DecryptData(uint8_t* data, uint32_t size, const uint8_t* private_key, const int private_key_size, uint8_t* crypt_key, int crypt_key_size)
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

			for (int p = key_size-1; p >= 0; p--)
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

		for (int j = private_key_size-1; j >= 0; j--)
		{
			data[i] = ROR(data[i], 3);
			data[i] -= i;
			data[i] ^= private_key[j];
		}
	}	
}

void ExecuteSaveTargetsForEncDec()
{
	Console::WriteLine();
	Console::WriteLine(L"Filename: ");
	String^ filename = Console::ReadLine();

	wstring fname = ItsCli::ToString(filename);

	Console::WriteLine();
	Console::WriteLine(L"Enter Filenames. <File> = Enter to stop.");
	
	vector<wstring> files;
	for (int i = 0; i < INT_MAX; i++) 
	{
		Console::Write(L"File: ");
		String^ ifile = Console::ReadLine();
		if (ifile->Length == 0)
		{
			break;
		}
		files.push_back(ItsCli::ToString(ifile));
	}
	
	SaveTargetsForEncDec(fname, files);
}

wstring SaveTargetsForEncDec(wstring filename, vector<wstring>& files)
{	
	wchar_t path[MAX_PATH];
	GetTempPath(MAX_PATH, path);

	wchar_t file[MAX_PATH];
	GetTempFileName(path, NULL, 0, file);	

	CComPtr<IXMLDOMDocument2> pIXMLDOMDocument;
	HRESULT hr = pIXMLDOMDocument.CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX::CLSCTX_INPROC_SERVER);
	if (FAILED(hr))
	{
		return L"";
	}

	CComPtr<IXMLDOMProcessingInstruction> pi;
	hr = pIXMLDOMDocument->createProcessingInstruction(CComBSTR(L"xml"), CComBSTR(L"version='1.0' encoding='utf-8'"), &pi);
	CComPtr<IXMLDOMNode> pin;
	hr = pIXMLDOMDocument->appendChild(pi, &pin);

	CComPtr<IXMLDOMElement> pIXDE;
	hr = pIXMLDOMDocument->createElement(CComBSTR(L"files"), &pIXDE);
	CComPtr<IXMLDOMNode> pIXDN;
	hr = pIXMLDOMDocument->appendChild(pIXDE, &pIXDN);
	
	for (auto& item : files)
	{
		CComPtr<IXMLDOMElement> element;
		hr = pIXMLDOMDocument->createElement(CComBSTR(L"file"), &element);

		CComPtr<IXMLDOMNode> node;
		hr = pIXDE->appendChild(element, &node);

		hr = node->put_text(CComBSTR(item.c_str()));
	}

	hr = pIXMLDOMDocument->save(CComVariant(filename.c_str()));

	return filename;
}