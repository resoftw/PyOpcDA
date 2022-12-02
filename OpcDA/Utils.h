#pragma once
#include <windows.h>
#include <Psapi.h>
#include <iostream>
#include <fstream>
#include <stdio.h>
#include <stdlib.h>
#include <stdarg.h>

#define DEBUG_OFF


time_t  filetime_to_timet(const FILETIME& ft);
size_t GetProcessMemory();
std::string Var2Str(const VARIANT& _In_ v);
std::string VarTypeStr(ULONG vt);
wchar_t* convertMBSToWCS(char const* value);
std::string DBTime(time_t);
std::string strTime(time_t t=0);
std::string strTime(std::string const& fmt,time_t t = 0);
std::string hResultStr(HRESULT h);

