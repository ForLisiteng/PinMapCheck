#pragma once
#include "afxdialogex.h"


// CDlgGeneral 对话框

class CDlgGeneral : public CDialogEx
{
	DECLARE_DYNAMIC(CDlgGeneral)

public:
	CDlgGeneral(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CDlgGeneral();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_GENERAL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnCbnSelchangeCombo1();
	CComboBox m_cmbUser;
	afx_msg void OnCbnSelchangeCombo2();
	CComboBox m_project_code;
	CComboBox m_vgl;
	CComboBox m_phy;
	CComboBox m_power_mode;
	CComboBox m_flash_voltage;
	CComboBox m_mipi_lane;
	CComboBox m_pmic;
	CComboBox m_pcd2;
	CComboBox m_finger;
	CComboBox m_mtp;
};
