#pragma once
#include "afxdialogex.h"


// CDlgPinMapInfo 对话框

class CDlgPinMapInfo : public CDialogEx
{
	DECLARE_DYNAMIC(CDlgPinMapInfo)

public:
	CDlgPinMapInfo(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CDlgPinMapInfo();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PINMAPINFO };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
private:
	CBrush m_brWhite;
public:
	CStatic m_pic;
	CEdit m_sheet_num;
	CEdit m_fop;
	CString icSideTopStr; // m_ic_side_top
	CString icSideBottomStr; // m_ic_side_bottom
	CEdit m_pin;
	CEdit m_ic_side;
	CEdit m_glass_railing;
	CEdit m_ic_side_top;
	CEdit m_ic_side_bottom;
	CEdit m_olb;
	afx_msg void OnPaint();
};
