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

struct FileTime : FILETIME
{
	FileTime() noexcept = default;
	FileTime(ULONGLONG t) noexcept : FILETIME{ DWORD(t), DWORD(t >> 32) } {}
	FileTime(const SYSTEMTIME& t) noexcept { THROW_IF_WIN32_BOOL_FALSE(SystemTimeToFileTime(&t, this)); }
	operator ULONGLONG() const noexcept { return *(const ULONGLONG*)this; }
	DWORD ToUnix() const noexcept { return *(const ULONGLONG*)this / 10'000'000 - 11644473600; }
	FileTime ToUtc() const
	{
		FileTime Local;
		THROW_IF_WIN32_BOOL_FALSE(LocalFileTimeToFileTime(this, &Local));
		return Local;
	}
	FileTime ToLocal() const
	{
		FileTime Local;
		THROW_IF_WIN32_BOOL_FALSE(FileTimeToLocalFileTime(this, &Local));
		return Local;
	}
	SYSTEMTIME ToSystemTime() const
	{
		SYSTEMTIME st;
		THROW_IF_WIN32_BOOL_FALSE(FileTimeToSystemTime(this, &st));
		return st;
	}
	static FileTime FromUnix(DWORD t) noexcept { return ((ULONGLONG)t + 11644473600) * 10'000'000; }
	static FileTime GetSystemTime()
	{
		FileTime t;
		GetSystemTimeAsFileTime(&t);
		return t;
	}
	static FileTime FromWbem(BSTR bstrDateTime)
	{
		WbemScripting::ISWbemDateTimePtr pDateTime;
		THROW_IF_FAILED(pDateTime.CreateInstance(__uuidof(WbemScripting::SWbemDateTime), nullptr, CLSCTX_INPROC_SERVER));
		pDateTime->Value = bstrDateTime;
		return _wtoi64(pDateTime->GetFileTime(VARIANT_TRUE));
	}
};

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

_bstr_t GetLastBootUpTime()
{
	ULONG                   ulReturn;
	_variant_t              vtLastBootUpTime;
	DATE                    Date;
	FILETIME                LastBootUpTime;
	IWbemLocatorPtr         pLocator;
	IWbemServicesPtr        pServices;
	IEnumWbemClassObjectPtr pEnum;
	IWbemClassObjectPtr     pClassObject;
	THROW_IF_FAILED(CoInitializeSecurity(nullptr, -1, nullptr, nullptr, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE, nullptr));
	THROW_IF_FAILED(pLocator.CreateInstance(__uuidof(WbemLocator), nullptr, CLSCTX_INPROC_SERVER));
	THROW_IF_FAILED(pLocator->ConnectServer(BSTR(L"ROOT\\CIMv2"), nullptr, nullptr, nullptr, 0, nullptr, nullptr, &pServices));
	THROW_IF_FAILED(CoSetProxyBlanket(pServices, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nullptr, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE));
	THROW_IF_FAILED(pServices->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_OperatingSystem"), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &pEnum));
	THROW_IF_FAILED(pEnum->Next(WBEM_INFINITE, 1, &pClassObject, &ulReturn));
	THROW_IF_FAILED(pClassObject->Get(L"LastBootUpTime", 0, &vtLastBootUpTime, nullptr, nullptr));
	return vtLastBootUpTime.bstrVal;
}

DWORD GetLastEventTime(LPCWSTR Source, DWORD EventID)
{
	auto hEventLog = wil::unique_any_handle_null<decltype(CloseEventLog), CloseEventLog>{ OpenEventLog(nullptr, Source) };
	THROW_LAST_ERROR_IF_NULL(hEventLog);

	struct : EVENTLOGRECORD { char Data[0x10000]; } EventLogRecord;
	for (;;)
	{
		DWORD nBytesRead, nMinNumberOfBytesNeeded;
		if (!ReadEventLog(hEventLog.get(), EVENTLOG_SEQUENTIAL_READ | EVENTLOG_BACKWARDS_READ, 0, &EventLogRecord, sizeof EventLogRecord, &nBytesRead, &nMinNumberOfBytesNeeded))
		{
			auto LastError = GetLastError();
			if (LastError == ERROR_HANDLE_EOF) break;
			THROW_WIN32(LastError);
		}
		if (EventLogRecord.EventID == EventID)
			return EventLogRecord.TimeGenerated;
	}
	return FileTime::GetSystemTime().ToUnix();
}

DWORD GetFirstEventTime(LPCWSTR Source)
{
	auto hEventLog = wil::unique_any_handle_null<decltype(CloseEventLog), CloseEventLog>{ OpenEventLog(nullptr, Source) };
	THROW_LAST_ERROR_IF_NULL(hEventLog);

	DWORD LastTime;
	struct : EVENTLOGRECORD { char Data[0x10000]; } EventLogRecord;

	SYSTEMTIME stToday;
	GetLocalTime(&stToday);
	stToday.wHour = 0;
	stToday.wMinute = 0;
	stToday.wSecond = 0;
	stToday.wMilliseconds = 0;
	auto Today1970 = FileTime(stToday).ToUtc().ToUnix();

	EventLogRecord.TimeGenerated = FileTime::GetSystemTime().ToUnix();
	do {
		LastTime = EventLogRecord.TimeGenerated;
		DWORD nBytesRead, nMinNumberOfBytesNeeded;
		if (!ReadEventLog(hEventLog.get(), EVENTLOG_SEQUENTIAL_READ | EVENTLOG_BACKWARDS_READ, 0, &EventLogRecord, sizeof EventLogRecord, &nBytesRead, &nMinNumberOfBytesNeeded))
		{
			auto LastError = GetLastError();
			if (LastError == ERROR_HANDLE_EOF) break;
			THROW_WIN32(LastError);
		}
	} while (EventLogRecord.TimeGenerated >= Today1970);
	return LastTime;
}

int main()
{
	SYSTEMTIME LocalTime;
	SYSTEMTIME SystemTime;
	FileTime SystemTimeAsFileTime;
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
	FileTime BootTimeAsFileTime = SystemTimeAsFileTime - TickCount * 10000;
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
	{
		auto bstrLastBootUpTime = GetLastBootUpTime();
		auto ftLastBootUpTime = FileTime::FromWbem(bstrLastBootUpTime);
		cout << "Win32_OperatingSystem.LastBootUpTime = " << ftLastBootUpTime << "\n";
		wcout << L"Win32_OperatingSystem.LastBootUpTime = " << bstrLastBootUpTime << L"\n";
	}
	CoUninitialize();

	cout << "Application                          = " << FileTime::FromUnix(GetFirstEventTime(L"Application")).ToLocal() << "\n";
	cout << "System                               = " << FileTime::FromUnix(GetFirstEventTime(L"System")).ToLocal() << "\n";
	cout << "Event Log Started                    = " << FileTime::FromUnix(GetLastEventTime(L"System", 6005/*Event Log Started*/)).ToLocal() << "\n";
}
