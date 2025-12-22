// CDlgPinMapInfo.cpp: 实现文件
//

#include "pch.h"
#include "PinmapCheckTool.h"
#include "afxdialogex.h"
#include "CDlgPinMapInfo.h"


// CDlgPinMapInfo 对话框

IMPLEMENT_DYNAMIC(CDlgPinMapInfo, CDialogEx)

CDlgPinMapInfo::CDlgPinMapInfo(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_PINMAPINFO, pParent)
{

}

CDlgPinMapInfo::~CDlgPinMapInfo()
{
}

void CDlgPinMapInfo::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    //DDX_Control(pDX, IDC_IMAGE, m_pic);
    DDX_Control(pDX, IDC_SHEET_NUM, m_sheet_num);
    DDX_Control(pDX, IDC_FOP_NO, m_fop);
    DDX_Control(pDX, IDC_PIN, m_pin);
    DDX_Control(pDX, IDC_IC_SIDE, m_ic_side);
    DDX_Control(pDX, IDC_GLASS_RAILING, m_glass_railing);
    DDX_Control(pDX, IDC_ICSIDE_TOP, m_ic_side_top);
    DDX_Control(pDX, IDC_ICSIDE_BOTTOM, m_ic_side_bottom);
    DDX_Control(pDX, IDC_OLB, m_olb);
}


BEGIN_MESSAGE_MAP(CDlgPinMapInfo, CDialogEx)
	ON_WM_CTLCOLOR()
    ON_WM_PAINT()
END_MESSAGE_MAP()


// CDlgPinMapInfo 消息处理程序

HBRUSH CDlgPinMapInfo::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

    // 处理对话框背景
    if (nCtlColor == CTLCOLOR_DLG)
    {
        pDC->SetBkColor(RGB(255, 255, 255)); // 设置背景色
        return m_brWhite; // 返回白色画刷
    }

    // 处理静态文本背景透明
    if (nCtlColor == CTLCOLOR_STATIC)
    {
        pDC->SetBkMode(TRANSPARENT);
        return (HBRUSH)GetStockObject(HOLLOW_BRUSH);
    }
	return hbr;
}

void CDlgPinMapInfo::OnPaint()
{
    CImage* m_sLogInPic;
    m_sLogInPic = new CImage;
    m_sLogInPic->Load(_T("./res/exwin.jpg"));
    CPaintDC dc(this); 
    CRect rect;
    GetDlgItem(IDC_PIC)->GetWindowRect(&rect);
    ScreenToClient(&rect);
    dc.SetStretchBltMode(HALFTONE);
    m_sLogInPic->Draw(dc.m_hDC, rect); 
    CDialogEx::OnPaint();
}
