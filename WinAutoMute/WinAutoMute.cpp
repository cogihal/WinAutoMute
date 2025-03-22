#include "pch.h"
#include "framework.h"
#include "WinAutoMute.h"
#include "WinAutoMuteDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


BEGIN_MESSAGE_MAP(CWinAutoMuteApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


CWinAutoMuteApp::CWinAutoMuteApp()
{
}


CWinAutoMuteApp::~CWinAutoMuteApp()
{
	delete m_pMainWnd;
}


CWinAutoMuteApp theApp;


BOOL CWinAutoMuteApp::InitInstance()
{
	CWinApp::InitInstance();

	CWinAutoMuteDlg* pDlg = new CWinAutoMuteDlg();
	m_pMainWnd = pDlg;
	pDlg->Create(IDD_WINAUTOMUTE_DIALOG);
	pDlg->ShowWindow(SW_HIDE);
	return TRUE;
}
