#pragma once
//
// #include
//
#include <msclr\marshal.h>
#include <msclr\marshal_cppstd.h>

//
// namespace
//
namespace ItSoftware
{
	namespace CLI
	{		
		using namespace System;

		class ItsCli
		{
		public:
			static wstring ToString(String^ str)
			{
				return msclr::interop::marshal_as<std::wstring>(str);
			}
		};
	}
}