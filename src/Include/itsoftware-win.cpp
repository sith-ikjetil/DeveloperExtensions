#include <Windows.h>
#include "itsoftware-win.h"

//
// using namespace
//
using namespace ItSoftware;
using namespace ItSoftware::Win;

namespace ItSoftware
{
	namespace Win
	{
		HHOOK ItsWin::hCBTHook = NULL;
		const wstring ItsTextFile::LineDelimiterWindows = L"\r\n";
		const wstring ItsTextFile::LineDelimiterUnix = L"\n";		
		const vector<BYTE> ItsTextFile::ByteOrderMarkUnicode = { 0xFF, 0xFE };
		const vector<BYTE> ItsTextFile::ByteOrderMarkUTF8 = { 0xEF, 0xBB, 0xBF };
	}
}