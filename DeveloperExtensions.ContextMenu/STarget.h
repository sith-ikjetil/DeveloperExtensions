////////////////////////////////////////////////////////////////////////////////////////
//
// File: STarget.h
//
#ifndef _STARGET_H_
#define _STARGET_H_
////////////////////////////////////////////////////////////////////////////////////////
//
// struct: CConfig
//
struct STarget
{
	bool IsComServer;
	bool IsManaged;
	bool IsDLL;
	bool IsTLB;
	bool IsEXE;
	bool IsNETMODULE;
	bool IsDirectory;
	wchar_t FileName[1024];	
	wchar_t ManagedVersion[32];
};// struct STarget

#endif