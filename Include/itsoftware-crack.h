#pragma once
//
// (i) When including this file, remember to in your Main.cpp file to include itsoftware-win.cpp.
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
#include <vector>
#include "itsoftware.h"
#include "itsoftware-win.h"
#include "itsoftware-exceptions.h"
#include <time.h>

//
// using namespace
//
using namespace ItSoftware;
using namespace ItSoftware::Win;
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
		using std::make_unique;
		using std::vector;
		using std::wstring;

		//
		// struct: ItsCrack
		//
		struct ItsCrack
		{
			//
			// Function: CrackFile
			//
			static bool CrackFile(wstring filename, const vector<uint8_t>& searchFor, const vector<uint8_t>& replacedBy, wstring& errorinfo)
			{
				if (searchFor.size() != replacedBy.size()) {
					errorinfo = L"Size error. searchFor and replaced by must be of same size.";
					return false;
				}
				if (searchFor.size() == 0 || replacedBy.size() == 0) {
					errorinfo = L"Size error. searchFor and replaced by cannot be zero length.";
					return false;
				}
				if (!ItsFile::Exists(filename)) {
					errorinfo = L"File not exist error. File could not be found.";
					return false;
				}

				size_t size_of_file{ 0 };
				if (!ItsFile::GetFileSize(filename, &size_of_file)) {
					errorinfo = L"GetFileSize error.";
					return false;
				}

				if (size_of_file == 0)
				{
					errorinfo = L"Size error. File has zero size.";
					return false;
				}

				ItsFile file;
				if (!file.OpenOrCreate(filename, L"rw", L"rw", ItsFileOpenCreation::OpenExisting)) {
					errorinfo = L"File open error. Could not open file for rw.";
					return false;
				}

				size_t size_read{ 0 };
				const DWORD bufferSize = 8192;
				unique_ptr<uint8_t[]> buffer = make_unique<uint8_t[]>(size_of_file);
				DWORD dwRead{ 0 };
				file.Read(buffer.get(), bufferSize, &dwRead);
				size_read += dwRead;
				while (size_of_file > size_read) {
					file.Read(&buffer.get()[size_read], bufferSize, &dwRead);
					size_read += dwRead;
				}

				bool found = true;
				size_t index{ 0 };
				for (size_t i{ 0 }; i < size_of_file; i++)
				{
					found = true;
					size_t j = 0;
					for (; j < searchFor.size(); j++) {
						if (buffer.get()[i + j] != searchFor.at(j))
						{
							found = false;
						}
					}

					if (found) {
						index = i;
						break;
					}
				}

				if (!found) {
					errorinfo = L"Not found error. searchFor not found in file.";
					return false;
				}

				if (!file.SetFilePosition(index, ItsFilePosition::FileBegin)) {
					errorinfo = L"File Position error. Could not set file position for replacing.";
					return false;
				}
				DWORD dwWritten{ 0 };
				if (!file.Write((BYTE*)replacedBy.data(), (DWORD)replacedBy.size(), &dwWritten)) {
					errorinfo = L"File write error. Could not write replaceBy to file.";
					return false;
				}
				if (dwWritten != (DWORD)replacedBy.size()) {
					errorinfo = L"File write error. Wrong number of bytes written. File might now be invalid.";
					return false;
				}

				return true;
			}

			//
			// Function: CrackFile
			//
			static bool CrackFile(wstring filename, const size_t startIndex, const vector<uint8_t>& patternToReplace, const vector<uint8_t>& replacedBy, wstring& errorinfo)
			{
				if (replacedBy.size() == 0) {
					errorinfo = L"Size error. replacedBy cannot be zero length.";
					return false;
				}
				if (!ItsFile::Exists(filename)) {
					errorinfo = L"File not exist error. File could not be found.";
					return false;
				}

				size_t size_of_file{ 0 };
				if (!ItsFile::GetFileSize(filename, &size_of_file)) {
					errorinfo = L"GetFileSize error. ";
					return false;
				}

				if (size_of_file == 0)
				{
					errorinfo = L"Size error. File has zero size.";
					return false;
				}

				if (startIndex + replacedBy.size() > size_of_file) {
					errorinfo = L"Size error. startIndex + replacedBy.size() must not be larger than file size.";
					return false;
				}

				ItsFile file;
				if (!file.OpenOrCreate(filename, L"rw", L"rw", ItsFileOpenCreation::OpenExisting)) {
					errorinfo = L"File open error. Could not open file for rw.";
					return false;
				}

				if (!file.SetFilePosition(startIndex, ItsFilePosition::FileBegin)) {
					errorinfo = L"File Position error. Could not set file position for replacing.";
					return false;
				}

				DWORD dwRead{ 0 };
				unique_ptr<BYTE[]> buffer = make_unique<BYTE[]>(patternToReplace.size());
				if (!file.Read(buffer.get(), (DWORD)patternToReplace.size(), &dwRead))
				{
					errorinfo = L"File read error. Could not read patternToReplace from file.";
					return false;
				}

				if (dwRead != (DWORD)patternToReplace.size())
				{
					errorinfo = L"File read error. Could not read the correct number of bytes.";
					return false;
				}

				DWORD i{ 0 };
				for (auto b : patternToReplace)
				{
					if (b != buffer.get()[i++]) {
						errorinfo = L"patternToReplace not found in file at location.";
						return false;
					}
				}

				if (!file.SetFilePosition(startIndex, ItsFilePosition::FileBegin)) {
					errorinfo = L"File Position error. Could not set file position for replacing.";
					return false;
				}

				DWORD dwWritten{ 0 };
				if (!file.Write((BYTE*)replacedBy.data(), (DWORD)replacedBy.size(), &dwWritten)) {
					errorinfo = L"File write error. Could not write replaceBy to file.";
					return false;
				}
				if (dwWritten != (DWORD)replacedBy.size()) {
					errorinfo = L"File write error. Wrong number of bytes written. File might now be invalid.";
					return false;
				}

				return true;
			}
		};// struct ItsCrack
	}// namespace Win
}// namespace ItSoftware