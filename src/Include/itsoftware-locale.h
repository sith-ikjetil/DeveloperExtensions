#pragma once
#ifndef __ITSOFTWARE_LOCALE_H__
#define __ITSOFTWARE_LOCALE_H__

#include <crtdbg.h>
#include <stdio.h>
#include <string>
#include <sstream>
#include <iomanip>
#include <exception>
#include <windows.h>	
#include "atlcomcli.h"		
#include <string>
#include <codecvt>
#include <random>
#include <functional>
using namespace ATL;
using namespace std;

//
// Put spesific methods in namespace
//
namespace ItSoftware
{
	namespace Locale
	{
		enum class ItsLocaleLanguage
		{
			American,
			AmericanEnglish,
			EnglishAmerican,
			EnglishUS,
			EnglishUSA,
			en_US,
			Australian,
			EnglishAus,
			en_AU,
			Belgian,
			DutchBelgian,
			nl_BE,
			Canadian,
			EnglishCan,
			en_CA,
			Chinese,
			zh,
			ChineseHongKong,
			zh_HK,
			ChineseSimplified,
			zh_CN,
			ChineseSingapore,
			zh_SG,
			ChineseTraditional,
			zh_TW,
			EnglishBelize,
			en_BZ,
			EnglishCaribbean,
			en_029,
			EnglishIre,
			IrishEnglish,
			en_IE,
			EnglishJamaica,
			en_JM,
			EnglishNZ,
			en_NZ,
			EnglishSouthAfrica,
			en_ZA,
			EnglishTrinidadyTobago,
			en_TT,
			EnglishUK,
			en_GB,
			FrenchBelgian,
			fr_BE,
			FrenchCanadian,
			fr_CA,
			FrenchLuxembourg,
			fr_LU,
			FrenchSwiss,
			fr_CH,
			GermanAustrian,
			de_AT,
			GermanLichtenstein,
			de_LI,
			GermanLuxembourg,
			de_LU,
			GermanSwiss,
			de_CH,
			ItalianSwiss,
			it_CH,
			Norwegian,
			no,
			NorwegianBokmal,
			nb_NO,
			NorwegianNyNorsk,
			nn_NO,
			PortugueseBrazilian,
			pt_BR,
			SpanishArgentina,
			es_AR,
			SpanishBolivia,
			es_BO,
			SpanishChile,
			es_CL,
			SpanishColombia,
			es_CO,
			SpanishCostaRica,
			es_CR,
			SpanishDominicanRepublic,
			es_DO,
			SpanishEcuador,
			es_EC,
			SpanishElSalvador,
			es_SV,
			SpanishGuatemala,
			es_GT,
			SpanishHonduras,
			es_HN,
			SpanishMexican,
			es_MX,
			SpanishModern,
			es_ES,
			SpanishNicaragua,
			es_NI,
			SpanishPanama,
			es_PA,
			SpanishParaguay,
			es_PY,
			SpanishPeru,
			es_PE,
			SpanishPuertoRico,
			es_PR,
			SpanishUruguay,
			es_UY,
			SpanishVenezuela,
			es_VE,
			SwedishFinland,
			sv_FI
		};

		struct ItsLocale
		{
			static void Initialize( ItsLocaleLanguage lan )
			{
				locale::global( locale( ItsLocale::GetLocaleString( lan ).c_str( ) ) );
			}

			static void Initialize( locale loc )
			{
				locale::global( loc );
			}

			static locale GetLocale( ItsLocaleLanguage lan ) {
				return locale( ItsLocale::GetLocaleString( lan ) );
			}

			static string GetLocaleString( ItsLocaleLanguage lan )
			{
				string loc;
				switch ( lan ) {
					case ItsLocaleLanguage::American:
					case ItsLocaleLanguage::AmericanEnglish:
					case ItsLocaleLanguage::EnglishAmerican:
					case ItsLocaleLanguage::EnglishUS:
					case ItsLocaleLanguage::EnglishUSA:
					case ItsLocaleLanguage::en_US:
						loc = "en-US";
						break;
					case ItsLocaleLanguage::Australian:
					case ItsLocaleLanguage::EnglishAus:
					case ItsLocaleLanguage::en_AU:
						loc = "en-AU";
						break;
					case ItsLocaleLanguage::Belgian:
					case ItsLocaleLanguage::DutchBelgian:
					case ItsLocaleLanguage::nl_BE:
						loc = "nl-BE";
						break;
					case ItsLocaleLanguage::Canadian:
					case ItsLocaleLanguage::EnglishCan:
					case ItsLocaleLanguage::en_CA:
						loc = "en-CA";
						break;
					case ItsLocaleLanguage::Chinese:
					case ItsLocaleLanguage::zh:
						loc = "zh";
						break;
					case ItsLocaleLanguage::ChineseHongKong:
					case ItsLocaleLanguage::zh_HK:
						loc = "zh-HK";
						break;
					case ItsLocaleLanguage::ChineseSimplified:
					case ItsLocaleLanguage::zh_CN:
						loc = "zh-CN";
						break;
					case ItsLocaleLanguage::ChineseSingapore:
					case ItsLocaleLanguage::zh_SG:
						loc = "zh-SG";
						break;
					case ItsLocaleLanguage::ChineseTraditional:
					case ItsLocaleLanguage::zh_TW:
						loc = "zh-TW";
						break;
					case ItsLocaleLanguage::EnglishBelize:
					case ItsLocaleLanguage::en_BZ:
						loc = "en-BZ";
						break;
					case ItsLocaleLanguage::EnglishCaribbean:
					case ItsLocaleLanguage::en_029:
						loc = "en-029";
						break;
					case ItsLocaleLanguage::EnglishIre:
					case ItsLocaleLanguage::IrishEnglish:
					case ItsLocaleLanguage::en_IE:
						loc = "en-IE";
						break;
					case ItsLocaleLanguage::EnglishJamaica:
					case ItsLocaleLanguage::en_JM:
						loc = "en-JM";
						break;
					case ItsLocaleLanguage::EnglishNZ:
					case ItsLocaleLanguage::en_NZ:
						loc = "en-NZ";
						break;
					case ItsLocaleLanguage::EnglishSouthAfrica:
					case ItsLocaleLanguage::en_ZA:
						loc = "en-ZA";
						break;
					case ItsLocaleLanguage::EnglishTrinidadyTobago:
					case ItsLocaleLanguage::en_TT:
						loc = "en-TT";
						break;
					case ItsLocaleLanguage::EnglishUK:
					case ItsLocaleLanguage::en_GB:
						loc = "en-GB";
						break;
					case ItsLocaleLanguage::FrenchBelgian:
					case ItsLocaleLanguage::fr_BE:
						loc = "fr-BE";
						break;
					case ItsLocaleLanguage::FrenchCanadian:
					case ItsLocaleLanguage::fr_CA:
						loc = "fr-CA";
						break;
					case ItsLocaleLanguage::FrenchLuxembourg:
					case ItsLocaleLanguage::fr_LU:
						loc = "fr-LU";
						break;
					case ItsLocaleLanguage::FrenchSwiss:
					case ItsLocaleLanguage::fr_CH:
						loc = "fr-CH";
						break;
					case ItsLocaleLanguage::GermanAustrian:
					case ItsLocaleLanguage::de_AT:
						loc = "de-AT";
						break;
					case ItsLocaleLanguage::GermanLichtenstein:
					case ItsLocaleLanguage::de_LI:
						loc = "de-LI";
						break;
					case ItsLocaleLanguage::GermanLuxembourg:
					case ItsLocaleLanguage::de_LU:
						loc = "de-LU";
						break;
					case ItsLocaleLanguage::GermanSwiss:
					case ItsLocaleLanguage::de_CH:
						loc = "de-CH";
						break;
					case ItsLocaleLanguage::ItalianSwiss:
					case ItsLocaleLanguage::it_CH:
						loc = "it-CH";
						break;
					case ItsLocaleLanguage::Norwegian:
					case ItsLocaleLanguage::no:
						loc = "no";
						break;
					case ItsLocaleLanguage::NorwegianBokmal:
					case ItsLocaleLanguage::nb_NO:
						loc = "nb-NO";
						break;
					case ItsLocaleLanguage::NorwegianNyNorsk:
					case ItsLocaleLanguage::nn_NO:
						loc = "nn-NO";
						break;
					case ItsLocaleLanguage::PortugueseBrazilian:
					case ItsLocaleLanguage::pt_BR:
						loc = "pt-BR";
						break;
					case ItsLocaleLanguage::SpanishArgentina:
					case ItsLocaleLanguage::es_AR:
						loc = "es-AR";
						break;
					case ItsLocaleLanguage::SpanishBolivia:
					case ItsLocaleLanguage::es_BO:
						loc = "es-BO";
						break;
					case ItsLocaleLanguage::SpanishChile:
					case ItsLocaleLanguage::es_CL:
						loc = "es-CL";
						break;
					case ItsLocaleLanguage::SpanishColombia:
					case ItsLocaleLanguage::es_CO:
						loc = "es-CO";
						break;
					case ItsLocaleLanguage::SpanishCostaRica:
					case ItsLocaleLanguage::es_CR:
						loc = "es-CR";
						break;
					case ItsLocaleLanguage::SpanishDominicanRepublic:
					case ItsLocaleLanguage::es_DO:
						loc = "es-DO";
						break;
					case ItsLocaleLanguage::SpanishEcuador:
					case ItsLocaleLanguage::es_EC:
						loc = "es-EC";
						break;
					case ItsLocaleLanguage::SpanishElSalvador:
					case ItsLocaleLanguage::es_SV:
						loc = "es-SV";
						break;
					case ItsLocaleLanguage::SpanishGuatemala:
					case ItsLocaleLanguage::es_GT:
						loc = "es-GT";
						break;
					case ItsLocaleLanguage::SpanishHonduras:
					case ItsLocaleLanguage::es_HN:
						loc = "es-HN";
						break;
					case ItsLocaleLanguage::SpanishMexican:
					case ItsLocaleLanguage::es_MX:
						loc = "es-MX";
						break;
					case ItsLocaleLanguage::SpanishModern:
					case ItsLocaleLanguage::es_ES:
						loc = "es-ES";
						break;
					case ItsLocaleLanguage::SpanishNicaragua:
					case ItsLocaleLanguage::es_NI:
						loc = "es-NI";
						break;
					case ItsLocaleLanguage::SpanishPanama:
					case ItsLocaleLanguage::es_PA:
						loc = "es-PA";
						break;
					case ItsLocaleLanguage::SpanishParaguay:
					case ItsLocaleLanguage::es_PY:
						loc = "es-PY";
						break;
					case ItsLocaleLanguage::SpanishPeru:
					case ItsLocaleLanguage::es_PE:
						loc = "es-PE";
						break;
					case ItsLocaleLanguage::SpanishPuertoRico:
					case ItsLocaleLanguage::es_PR:
						loc = "es-PR";
						break;
					case ItsLocaleLanguage::SpanishUruguay:
					case ItsLocaleLanguage::es_UY:
						loc = "es-UY";
						break;
					case ItsLocaleLanguage::SpanishVenezuela:
					case ItsLocaleLanguage::es_VE:
						loc = "es-VE";
						break;
					case ItsLocaleLanguage::SwedishFinland:
					case ItsLocaleLanguage::sv_FI:
						loc = "sv-FI";
						break;
					default:
						break;
				}
				return loc;
			}
		};// ItsLocale
	}// namespace locale
}// namespace
#endif