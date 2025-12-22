#pragma once
#include "afxdialogex.h"


// CDlComponent 对话框

class CDlgComponent : public CDialogEx
{
	DECLARE_DYNAMIC(CDlgComponent)

public:
	CDlgComponent(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CDlgComponent();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_COMPONENT_DL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CEdit m_ddic;
	CString ddicStr_ = _T("U2");
	CEdit m_cci;
	CString cciStr_ = _T("CN1");
	CEdit m_fci;
	CString fciStr_ = _T("F1");
	afx_msg void OnPaint();
};
