#pragma once

#include <mmdeviceapi.h>

#define		WM_TRAY_MESSAGE		WM_USER+1


class CWinAutoMuteDlg : public CDialogEx
{
public:
	CWinAutoMuteDlg(CWnd* pParent = nullptr);	// standard constructor

#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_WINAUTOMUTE_DIALOG };
#endif


protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


protected:
	HICON			m_hIcon;
	NOTIFYICONDATA	m_IconData;

	virtual BOOL OnInitDialog();
	LRESULT OnTrayNotification(WPARAM wParam, LPARAM lParam);
	DECLARE_MESSAGE_MAP()


public:
	afx_msg void OnTraySettings();
	afx_msg void OnProcessNow();
	afx_msg void OnTrayExit();
	afx_msg void OnEndSession(BOOL bEnding);
	virtual void OnOK();
	virtual void OnCancel();


protected:
	BOOL	m_bMute, m_bMuteOrg;
	BOOL	m_bVolume, m_bVolumeOrg;
	BOOL	m_bAllSpeakers, m_bAllSpeakersOrg;
	BOOL	m_bLogging;

	void LoadSettings();
	void SaveSettings();

	void ProcessAllSpeakers(BOOL bMute, BOOL bVolume);
	void ProcessCurrentSpeaker(BOOL bMute, BOOL bVolume);

	void Logging(IMMDevice* pDevice, BOOL bMute, BOOL bVolume);
	CString GetDeviceName(IMMDevice* pDevice);
};
