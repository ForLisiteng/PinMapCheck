#pragma once
#include "afxdialogex.h"


// CDlgHint 对话框

class CDlgHint : public CDialogEx
{
	DECLARE_DYNAMIC(CDlgHint)

public:
	CDlgHint(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CDlgHint();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_HINT };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
};
