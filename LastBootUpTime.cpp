#include <Windows.h>
#include <iostream>
#include <iomanip>
#include <wbemidl.h>
#include <comdef.h>
#include <comip.h>

#include "wil/result.h"

#import <C:\Windows\System32\wbem\WbemDisp.tlb>

#pragma comment(lib, "wbemuuid.lib")

_COM_SMARTPTR_TYPEDEF(IWbemLocator        , __uuidof(IWbemLocator));
_COM_SMARTPTR_TYPEDEF(IWbemServices       , __uuidof(IWbemServices));
_COM_SMARTPTR_TYPEDEF(IWbemClassObject    , __uuidof(IWbemClassObject));
_COM_SMARTPTR_TYPEDEF(IEnumWbemClassObject, __uuidof(IEnumWbemClassObject));

using namespace std;
using namespace WbemScripting;

ostream& operator<<(ostream& o, const SYSTEMTIME& t)
{
	return o << t.wDay << "." << setfill('0') << setw(2) << t.wMonth << "." << t.wYear << " " << t.wHour << ":" << setfill('0') << setw(2) << t.wMinute << ":" << setfill('0') << setw(2) << t.wSecond << "." << t.wMilliseconds;
}

ostream& operator<<(ostream& o, const FILETIME& ft)
{
	SYSTEMTIME t;
	THROW_IF_WIN32_BOOL_FALSE(FileTimeToSystemTime(&ft, &t));
	return o << t.wDay << "." << setfill('0') << setw(2) << t.wMonth << "." << t.wYear << " " << t.wHour << ":" << setfill('0') << setw(2) << t.wMinute << ":" << setfill('0') << setw(2) << t.wSecond << "." << t.wMilliseconds;
}

void GetLastBootTime()
{
	ULONG                   ulReturn;
	_variant_t              vtLastBootUpTime;
	DATE                    Date;
	FILETIME                LastBootUpTime;
	IWbemLocatorPtr         pLocator;
	IWbemServicesPtr        pServices;
	IEnumWbemClassObjectPtr pEnum;
	IWbemClassObjectPtr     pClassObject;
	ISWbemDateTimePtr       pDateTime;
	THROW_IF_FAILED(CoInitializeSecurity(nullptr, -1, nullptr, nullptr, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE, nullptr));
	THROW_IF_FAILED(pLocator.CreateInstance(__uuidof(WbemLocator), nullptr, CLSCTX_INPROC_SERVER));
	THROW_IF_FAILED(pLocator->ConnectServer(BSTR(L"ROOT\\CIMv2"), nullptr, nullptr, nullptr, 0, nullptr, nullptr, &pServices));
	THROW_IF_FAILED(CoSetProxyBlanket(pServices, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nullptr, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE));
	THROW_IF_FAILED(pServices->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_OperatingSystem"), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &pEnum));
	THROW_IF_FAILED(pEnum->Next(WBEM_INFINITE, 1, &pClassObject, &ulReturn));
	THROW_IF_FAILED(pClassObject->Get(L"LastBootUpTime", 0, &vtLastBootUpTime, nullptr, nullptr));
//	wcout << L"Win32_OperatingSystem.LastBootUpTime = " << vtLastBootUpTime.bstrVal << L"\n";

	THROW_IF_FAILED(pDateTime.CreateInstance(__uuidof(WbemScripting::SWbemDateTime), nullptr, CLSCTX_INPROC_SERVER));
	pDateTime->Value = vtLastBootUpTime.bstrVal;
	(ULONGLONG&)LastBootUpTime = _wtoi64(pDateTime->GetFileTime(VARIANT_TRUE));
	cout << "Win32_OperatingSystem.LastBootUpTime = " << LastBootUpTime << "\n";
}

int main()
{
	SYSTEMTIME LocalTime;
	SYSTEMTIME SystemTime;
	FILETIME SystemTimeAsFileTime;
	FILETIME BootTimeAsFileTime;
	SYSTEMTIME BootTime;
	SYSTEMTIME LocalBootTime;
	SYSTEMTIME LocalBootTimeTz;
	SYSTEMTIME LocalBootTimeEx;
	TIME_ZONE_INFORMATION TimeZone;
	DYNAMIC_TIME_ZONE_INFORMATION DynamicTimeZone;
	GetLocalTime(&LocalTime);
	GetSystemTime(&SystemTime);
	GetSystemTimeAsFileTime(&SystemTimeAsFileTime);
	auto TickCount = GetTickCount64();
	auto TimeZoneId = GetTimeZoneInformation(&TimeZone);
	auto DynamicTimeZoneId = GetDynamicTimeZoneInformation(&DynamicTimeZone);
	(ULONGLONG&)BootTimeAsFileTime = (ULONGLONG&)SystemTimeAsFileTime - TickCount * 10000;
	THROW_IF_WIN32_BOOL_FALSE(FileTimeToSystemTime(&BootTimeAsFileTime, &BootTime));
	THROW_IF_WIN32_BOOL_FALSE(SystemTimeToTzSpecificLocalTime(nullptr, &BootTime, &LocalBootTime));
	THROW_IF_WIN32_BOOL_FALSE(SystemTimeToTzSpecificLocalTime(&TimeZone, &BootTime, &LocalBootTimeTz));
	THROW_IF_WIN32_BOOL_FALSE(SystemTimeToTzSpecificLocalTimeEx(&DynamicTimeZone, &BootTime, &LocalBootTimeEx));

	cout << "GetTickCount                         = " << TickCount / 86400'000        << "."
	                       << setfill('0') << setw(2) << TickCount /  3600'000 % 24   << ":"
	                       << setfill('0') << setw(2) << TickCount /    60'000 % 60   << ":"
	                       << setfill('0') << setw(2) << TickCount /      1000 % 60   << "."
	                       << setfill('0') << setw(3) << TickCount             % 1000 << "\n";
	cout << "GetLocalTime                         = " << LocalTime       << "\n";
	cout << "GetSystemTime                        = " << SystemTime      << "\n";
	cout << "FileTimeToSystemTime                 = " << BootTime        << "\n";
	cout << "SystemTimeToTzSpecificLocalTime      = " << LocalBootTime   << "\n";
	cout << "SystemTimeToTzSpecificLocalTime      = " << LocalBootTimeTz << "\n";
	cout << "SystemTimeToTzSpecificLocalTimeEx    = " << LocalBootTimeEx << "\n";

	THROW_IF_FAILED(CoInitializeEx(nullptr, COINIT_MULTITHREADED));
	GetLastBootTime();
	CoUninitialize();
}
