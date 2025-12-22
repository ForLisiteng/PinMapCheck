// CDlgGeneral.cpp: 实现文件
//

#include "pch.h"
#include "PinmapCheckTool.h"
#include "afxdialogex.h"
#include "CDlgGeneral.h"


// CDlgGeneral 对话框

IMPLEMENT_DYNAMIC(CDlgGeneral, CDialogEx)

CDlgGeneral::CDlgGeneral(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_GENERAL, pParent)
{

}

CDlgGeneral::~CDlgGeneral()
{
}

void CDlgGeneral::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_COMBO1, m_cmbUser);
	DDX_Control(pDX, IDC_COMBO2, m_project_code);
	DDX_Control(pDX, IDC_COMBO3, m_vgl);
	DDX_Control(pDX, IDC_COMBO5, m_phy);
	DDX_Control(pDX, IDC_COMBO11, m_power_mode);
	DDX_Control(pDX, IDC_COMBO4, m_flash_voltage);
	DDX_Control(pDX, IDC_COMBO6, m_mipi_lane);
	DDX_Control(pDX, IDC_COMBO10, m_pmic);
	DDX_Control(pDX, IDC_COMBO7, m_pcd2);
	DDX_Control(pDX, IDC_COMBO8, m_finger);
	DDX_Control(pDX, IDC_COMBO9, m_mtp);
}


BEGIN_MESSAGE_MAP(CDlgGeneral, CDialogEx)
	ON_CBN_SELCHANGE(IDC_COMBO1, &CDlgGeneral::OnCbnSelchangeCombo1)
	ON_CBN_SELCHANGE(IDC_COMBO2, &CDlgGeneral::OnCbnSelchangeCombo2)
END_MESSAGE_MAP()


// CDlgGeneral 消息处理程序

void CDlgGeneral::OnCbnSelchangeCombo1()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CDlgGeneral::OnCbnSelchangeCombo2()
{
	// TODO: 在此添加控件通知处理程序代码
}
