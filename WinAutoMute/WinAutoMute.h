#pragma once

#ifndef __AFXWIN_H__
#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


class CWinAutoMuteApp : public CWinApp
{
public:
	CWinAutoMuteApp();
	virtual ~CWinAutoMuteApp();

public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};


extern CWinAutoMuteApp theApp;
