#include "pch.h"
#include "framework.h"
#include "WinAutoMute.h"
#include "WinAutoMuteDlg.h"
#include "afxdialogex.h"

#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <Functiondiscoverykeys_devpkey.h>

#include <locale.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CWinAutoMuteDlg::CWinAutoMuteDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_WINAUTOMUTE_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_bMute = m_bMuteOrg = TRUE;
	m_bVolume = m_bVolumeOrg = FALSE;
	m_bAllSpeakers = m_bAllSpeakersOrg = TRUE;
	m_bLogging = FALSE;
}


void CWinAutoMuteDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Check(pDX, IDC_MUTE, m_bMute);
	DDX_Check(pDX, IDC_VOLUME, m_bVolume);
	DDX_Check(pDX, IDC_ALL_SPEAKERS, m_bAllSpeakers);
}


BEGIN_MESSAGE_MAP(CWinAutoMuteDlg, CDialogEx)
	ON_MESSAGE(WM_TRAY_MESSAGE, OnTrayNotification)
	ON_COMMAND(ID_TRAY_SETTINGS, &CWinAutoMuteDlg::OnTraySettings)
	ON_COMMAND(ID_PROCESS_NOW, &CWinAutoMuteDlg::OnProcessNow)
	ON_COMMAND(ID_TRAY_EXIT, &CWinAutoMuteDlg::OnTrayExit)
	ON_WM_ENDSESSION()
END_MESSAGE_MAP()


BOOL CWinAutoMuteDlg::OnInitDialog()
{
	m_IconData.cbSize = sizeof(NOTIFYICONDATA);
	m_IconData.hWnd = this->m_hWnd;
	m_IconData.uID = IDR_MAINFRAME;
	m_IconData.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	m_IconData.uCallbackMessage = WM_TRAY_MESSAGE;
	m_IconData.hIcon = AfxGetApp()->LoadIcon(IDI_TRAY_ICON);
	_tcscpy_s(m_IconData.szTip, _T("WinAutoMute"));

	Shell_NotifyIcon(NIM_ADD, &m_IconData);

	ShowWindow(SW_HIDE);

	LoadSettings();
	UpdateData(FALSE);

	return TRUE;
}


LRESULT CWinAutoMuteDlg::OnTrayNotification(WPARAM wParam, LPARAM lParam)
{
	if (wParam != IDR_MAINFRAME)
		return 0;

	if (lParam == WM_LBUTTONDOWN || lParam == WM_RBUTTONDOWN) {
		CMenu menu;
		menu.LoadMenu(IDR_TRAY_MENU);
		CMenu* pSubMenu = menu.GetSubMenu(0);
		CPoint pos;
		GetCursorPos(&pos);
		SetForegroundWindow();
		pSubMenu->TrackPopupMenu(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON, pos.x, pos.y, this);
		PostMessage(WM_NULL, 0, 0);
	}

	return 0;
}


void CWinAutoMuteDlg::OnTraySettings()
{
	// Save original values before showing the dialog
	m_bMuteOrg = m_bMute;
	m_bVolumeOrg = m_bVolume;
	m_bAllSpeakersOrg = m_bAllSpeakers;
	ShowWindow(SW_SHOW);
}


void CWinAutoMuteDlg::OnTrayExit()
{
	Shell_NotifyIcon(NIM_DELETE, &m_IconData);
	PostQuitMessage(0);
}


void CWinAutoMuteDlg::OnOK()
{
	ShowWindow(SW_HIDE);
	UpdateData(TRUE);
	SaveSettings();
}


void CWinAutoMuteDlg::OnCancel()
{
	ShowWindow(SW_HIDE);
	// Restore original values
	m_bMute = m_bMuteOrg;
	m_bVolume = m_bVolumeOrg;
	m_bAllSpeakers = m_bAllSpeakersOrg;
	UpdateData(FALSE);
}


void CWinAutoMuteDlg::LoadSettings()
{
	// Get current path name
	TCHAR szPath[MAX_PATH];
	GetModuleFileName(NULL, szPath, MAX_PATH);
	PathRemoveFileSpec(szPath);

	// Generate ini file path name
	TCHAR szIniPath[MAX_PATH];
	_stprintf_s(szIniPath, _T("%s\\WinAutoMute.ini"), szPath);

	// Load settings from ini file
	m_bMute = GetPrivateProfileInt(_T("Settings"), _T("Mute"), 1, szIniPath);
	m_bVolume = GetPrivateProfileInt(_T("Settings"), _T("Volume"), 0, szIniPath);
	m_bAllSpeakers = GetPrivateProfileInt(_T("Settings"), _T("AllSpeakers"), 1, szIniPath);
	m_bLogging = GetPrivateProfileInt(_T("Settings"), _T("Logging"), 0, szIniPath);
}


void CWinAutoMuteDlg::SaveSettings()
{
	// Get current path name
	TCHAR szPath[MAX_PATH];
	GetModuleFileName(NULL, szPath, MAX_PATH);
	PathRemoveFileSpec(szPath);

	// Generate ini file path name
	TCHAR szIniPath[MAX_PATH];
	_stprintf_s(szIniPath, _T("%s\\WinAutoMute.ini"), szPath);

	// Save settings to ini file
	WritePrivateProfileString(_T("Settings"), _T("Mute"), m_bMute ? _T("1") : _T("0"), szIniPath);
	WritePrivateProfileString(_T("Settings"), _T("Volume"), m_bVolume ? _T("1") : _T("0"), szIniPath);
	WritePrivateProfileString(_T("Settings"), _T("AllSpeakers"), m_bAllSpeakers ? _T("1") : _T("0"), szIniPath);
}


void CWinAutoMuteDlg::OnEndSession(BOOL bEnding)
{
	CDialogEx::OnEndSession(bEnding);

	UpdateData(TRUE);

	if (m_bAllSpeakers) {
		ProcessAllSpeakers(m_bMute, m_bVolume);
	}
	else {
		ProcessCurrentSpeaker(m_bMute, m_bVolume);
	}
}


void CWinAutoMuteDlg::OnProcessNow()
{
	UpdateData(TRUE);
	if (m_bAllSpeakers) {
		ProcessAllSpeakers(m_bMute, m_bVolume);
	}
	else {
		ProcessCurrentSpeaker(m_bMute, m_bVolume);
	}
}


void CWinAutoMuteDlg::ProcessAllSpeakers(BOOL bMute, BOOL bVolume)
{
	CoInitialize(NULL);

	IMMDeviceEnumerator* pEnumerator = NULL;
	IMMDeviceCollection* pCollection = NULL;
	IMMDevice* pDevice = NULL;
	IAudioEndpointVolume* pAudioEndpointVolume = NULL;

	HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
	if (SUCCEEDED(hr)) {
		// Enumerate all speakers
		hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pCollection);
		if (SUCCEEDED(hr)) {
			UINT count;
			pCollection->GetCount(&count);
			for (UINT i = 0; i < count; i++) {
				hr = pCollection->Item(i, &pDevice);
				if (SUCCEEDED(hr)) {
					hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pAudioEndpointVolume);
					if (SUCCEEDED(hr)) {
						// Mute and/or set volume to Zero
						float fVolume;
						pAudioEndpointVolume->GetMasterVolumeLevelScalar(&fVolume);
						// Mute
						if (bMute) {
							pAudioEndpointVolume->SetMute(TRUE, NULL);
						}
						// Set volume to Zero.
						if (bVolume) {
							pAudioEndpointVolume->SetMasterVolumeLevelScalar(0.0f, NULL);
						}
					}

					// Logging
					Logging(pDevice, bMute, bVolume);

					pDevice->Release();
				}
			}
			pCollection->Release();
		}
		pEnumerator->Release();
	}
	CoUninitialize();
}


void CWinAutoMuteDlg::ProcessCurrentSpeaker(BOOL bMute, BOOL bVolume)
{
	CoInitialize(NULL);

	IMMDeviceEnumerator* pEnumerator = NULL;
	IMMDevice* pDevice = NULL;
	IAudioEndpointVolume* pAudioEndpointVolume = NULL;

	HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
	if (SUCCEEDED(hr)) {
		// Get current audio device
		hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
		if (SUCCEEDED(hr)) {
			hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pAudioEndpointVolume);
			if (SUCCEEDED(hr)) {
				// Mute and/or set volume to Zero
				float fVolume;
				pAudioEndpointVolume->GetMasterVolumeLevelScalar(&fVolume);
				// Mute
				if (bMute) {
					pAudioEndpointVolume->SetMute(TRUE, NULL);
				}
				// Set volume to Zero.
				if (bVolume) {
					pAudioEndpointVolume->SetMasterVolumeLevelScalar(0.0f, NULL);
				}
			}

			// Logging
			Logging(pDevice, bMute, bVolume);

			pDevice->Release();
		}
		pEnumerator->Release();
	}
	CoUninitialize();
}


void CWinAutoMuteDlg::Logging(IMMDevice* pDevice, BOOL bMute, BOOL bVolume)
{
	if (!m_bLogging) {
		return;
	}

	// Get device name
	CString strDeviceName = GetDeviceName(pDevice);

	// Get current time
	CTime time = CTime::GetCurrentTime();

	// Get current path name
	TCHAR szPath[MAX_PATH];
	GetModuleFileName(NULL, szPath, MAX_PATH);
	PathRemoveFileSpec(szPath);
	// Generate log file path name
	TCHAR szLogPath[MAX_PATH];
	_stprintf_s(szLogPath, TEXT("%s\\WinAutoMute.log"), szPath);

	// Set locale to system default
	setlocale(LC_ALL, "");

	// Open log file
	FILE* fp;
	_tfopen_s(&fp, szLogPath, TEXT("a"));
	if (fp) {
		_ftprintf(fp, TEXT("%04d/%02d/%02d %02d:%02d:%02d : %s - Mute(%s) Volume(%s)\n"),
			time.GetYear(), time.GetMonth(), time.GetDay(),
			time.GetHour(), time.GetMinute(), time.GetSecond(),
			strDeviceName,
			bMute ? TEXT("True") : TEXT("False"),
			bVolume ? TEXT("True") : TEXT("False"));
		fclose(fp);
	}
}


CString CWinAutoMuteDlg::GetDeviceName(IMMDevice* pDevice)
{
	IPropertyStore* pPropertyStore = NULL;
	CString strDeviceName = TEXT("");
	HRESULT hr = pDevice->OpenPropertyStore(STGM_READ, &pPropertyStore);
	if (SUCCEEDED(hr)) {
		PROPVARIANT pv;
		PropVariantInit(&pv);
		hr = pPropertyStore->GetValue(PKEY_Device_FriendlyName, &pv);
		if (SUCCEEDED(hr)) {
			strDeviceName = pv.pwszVal;
			PropVariantClear(&pv);
		}
		pPropertyStore->Release();
	}
	return strDeviceName;
}
