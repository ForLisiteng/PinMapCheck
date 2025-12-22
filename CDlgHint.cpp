// CDlgHint.cpp: 实现文件
//

#include "pch.h"
#include "PinmapCheckTool.h"
#include "afxdialogex.h"
#include "CDlgHint.h"


// CDlgHint 对话框

IMPLEMENT_DYNAMIC(CDlgHint, CDialogEx)

CDlgHint::CDlgHint(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_HINT, pParent)
{

}

CDlgHint::~CDlgHint()
{
}

void CDlgHint::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CDlgHint, CDialogEx)
END_MESSAGE_MAP()


// CDlgHint 消息处理程序
