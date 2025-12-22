// CDlgComponent.cpp: 实现文件
//

#include "pch.h"
#include "PinmapCheckTool.h"
#include "afxdialogex.h"
#include "CDlgComponent.h"


// CDlComponent 对话框

IMPLEMENT_DYNAMIC(CDlgComponent, CDialogEx)

CDlgComponent::CDlgComponent(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_COMPONENT_DL, pParent)
{

}

CDlgComponent::~CDlgComponent()
{
}

void CDlgComponent::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT1, m_ddic);
	DDX_Control(pDX, IDC_EDIT2, m_cci);
	DDX_Control(pDX, IDC_EDIT3, m_fci);
}


BEGIN_MESSAGE_MAP(CDlgComponent, CDialogEx)
	ON_WM_PAINT()
END_MESSAGE_MAP()


// CDlgComponent 消息处理程序

void CDlgComponent::OnPaint()
{
    CImage* m_sLogInPic;
    m_sLogInPic = new CImage;
    m_sLogInPic->Load(_T("./res/ddic.jpg"));
    CPaintDC dc(this);
    CRect rect;
    GetDlgItem(IDC_DDIC_PIC)->GetWindowRect(&rect);
    ScreenToClient(&rect);
    dc.SetStretchBltMode(HALFTONE);
    m_sLogInPic->Draw(dc.m_hDC, rect);

    CImage* m_sLogInPic1;
    m_sLogInPic1 = new CImage;
    m_sLogInPic1->Load(_T("./res/cci.jpg"));
    CRect rect1;
    GetDlgItem(IDC_CCI_PIC)->GetWindowRect(&rect1);
    ScreenToClient(&rect1);
    dc.SetStretchBltMode(HALFTONE);
    m_sLogInPic1->Draw(dc.m_hDC, rect1);

    CImage* m_sLogInPic2;
    m_sLogInPic2 = new CImage;
    m_sLogInPic2->Load(_T("./res/fci.jpg"));
    CRect rect2;
    GetDlgItem(IDC_FCI_PIC)->GetWindowRect(&rect2);
    ScreenToClient(&rect2);
    dc.SetStretchBltMode(HALFTONE);
    m_sLogInPic2->Draw(dc.m_hDC, rect2);

    CDialogEx::OnPaint();
}
