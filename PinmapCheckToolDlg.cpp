
// PinmapCheckToolDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "PinmapCheckTool.h"
#include "PinmapCheckToolDlg.h"
#include "afxdialogex.h"
#include "CDlgPinMapInfo.h"
#include <afxdisp.h>  // MFC OLE Automation 支持
#include <vector>
#include <utility> // pair 头文件

using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CPinmapCheckToolDlg 对话框



CPinmapCheckToolDlg::CPinmapCheckToolDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_PINMAPCHECKTOOL_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CPinmapCheckToolDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_BTN_SEL_PINMAP, m_btnSelPinmap);
	DDX_Control(pDX, IDC_BTN_LOAD_NETLIST, m_btnLoadNetlist);
	DDX_Control(pDX, IDC_BTN_LOAD_BOM, m_btnLoadBOM);
	DDX_Control(pDX, IDC_BTN_LOAD_CFG, m_btnLoadCfg);
	DDX_Control(pDX, IDC_BTN_SAVE_CFG, m_btnSaveCfg);
	DDX_Control(pDX, IDC_BTN_CHECK, m_btnCheck);
	DDX_Control(pDX, IDC_BTN_SAVE_LOG, m_btnSaveLog);
	DDX_Control(pDX, IDC_TAB_CTRL, m_tabCtrl);
	DDX_Control(pDX, IDC_STC_SHEET_NAME, m_stcSheetName);
	DDX_Control(pDX, IDC_STC_MESSAGE, m_stcMessage);
	DDX_Control(pDX, IDC_RE_INFO, m_reInfo);
	DDX_Control(pDX, IDC_TEXT_PERCENT, m_text_percent);
	DDX_Control(pDX, IDC_PINMAP_NAME, m_staticPinmapName);
	DDX_Control(pDX, IDC_LOADNETLIST_NAME, m_netList);
	DDX_Control(pDX, IDC_LOADBOM_NAME, m_bom);
	DDX_Control(pDX, IDC_RE_CHECK, m_re_check);
	DDX_Control(pDX, IDC_SETTING, m_setting);
	DDX_Control(pDX, IDC_INFORMATION, m_information);
	DDX_Control(pDX, IDC_PASS, m_pass);
	DDX_Control(pDX, IDC_PROGRESS_CHECK, m_progressCheck);
}

BEGIN_MESSAGE_MAP(CPinmapCheckToolDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BTN_SEL_PINMAP, &CPinmapCheckToolDlg::OnBnClickedSelPinmap)
	ON_BN_CLICKED(IDC_BTN_LOAD_NETLIST, &CPinmapCheckToolDlg::OnBnClickedLoadNetlist)
	ON_BN_CLICKED(IDC_BTN_LOAD_BOM, &CPinmapCheckToolDlg::OnBnClickedLoadBom)
	ON_BN_CLICKED(IDC_BTN_LOAD_CFG, &CPinmapCheckToolDlg::OnBnClickedLoadCfg)
	ON_BN_CLICKED(IDC_BTN_SAVE_CFG, &CPinmapCheckToolDlg::OnBnClickedSaveCfg)
	ON_BN_CLICKED(IDC_BTN_CHECK, &CPinmapCheckToolDlg::OnBnClickedCheck)
	ON_BN_CLICKED(IDC_BTN_SAVE_LOG, &CPinmapCheckToolDlg::OnBnClickedSaveLog)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB_CTRL, &CPinmapCheckToolDlg::OnTcnSelchangeTab)
	ON_WM_CTLCOLOR()
	ON_EN_CHANGE(IDC_EDT_SHEET_NAME2, &CPinmapCheckToolDlg::OnEnChangeEdtSheetName2)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST_PINMAP, &CPinmapCheckToolDlg::OnLvnItemchangedListPinmap)
	ON_EN_CHANGE(IDC_EDT_SHEET_NAME, &CPinmapCheckToolDlg::OnEnChangeEdtSheetName)
	ON_EN_CHANGE(IDC_EDT_MESSAGE, &CPinmapCheckToolDlg::OnEnChangeEdtMessage)
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
END_MESSAGE_MAP()


// CPinmapCheckToolDlg 消息处理程序

BOOL CPinmapCheckToolDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置图标
	SetIcon(m_hIcon, TRUE);
	SetIcon(m_hIcon, FALSE);

	// 初始化标题
	SetWindowText(_T("Pinmap Check Tool"));

	m_btnSelPinmap.SetWindowText(_T("Sel Pinmap"));
	m_btnLoadNetlist.SetWindowText(_T("Load netlist"));
	m_btnLoadBOM.SetWindowText(_T("Load BOM"));
	m_btnLoadCfg.SetWindowText(_T("Load Cfg"));
	m_btnSaveCfg.SetWindowText(_T("Save Cfg"));
	m_btnCheck.SetWindowText(_T("Check"));
	m_btnSaveLog.SetWindowText(_T("Save Log"));

	// 初始化标签页
	m_tabCtrl.InsertItem(0, _T("Pin Map Info"));
	m_tabCtrl.InsertItem(1, _T("General"));
	m_tabCtrl.InsertItem(2, _T("Component"));
	m_tabCtrl.InsertItem(3, _T("Hint"));

	// 初始化信息区域
	m_stcSheetName.SetWindowText(_T("Select Sheet Name:"));
	m_stcMessage.SetWindowText(_T("message % :"));

	// 初始化富文本框（信息显示区域）
	m_reInfo.SetReadOnly(TRUE);

	m_btnSelPinmap.EnableWindowsTheming(FALSE);

	CRect rc;
	m_tabCtrl.GetClientRect(&rc);
	rc.DeflateRect(1, 22, 3, 1);
	m_pinmapinfo_.Create(IDD_PINMAPINFO, &m_tabCtrl);
	m_pinmapinfo_.MoveWindow(&rc);
	m_pinmapinfo_.m_sheet_num.SetWindowText(sheetNumStr_);
	m_pinmapinfo_.m_fop.SetWindowText(fopStr_);
	m_pinmapinfo_.m_glass_railing.SetWindowText(glassRailingStr_);
	m_pinmapinfo_.m_pin.SetWindowText(pinColStr_);
	m_pinmapinfo_.m_ic_side.SetWindowText(icSideColStr_);
	m_pinmapinfo_.m_olb.SetWindowText(olbStr_);
	m_pinmapinfo_.m_ic_side_top.SetWindowTextW(topRowStr_);
	m_pinmapinfo_.m_ic_side_bottom.SetWindowTextW(bottomRowStr_);
	m_pinmapinfo_.ShowWindow(SW_SHOW);

	InitGeneralComboBox();

	m_component_.Create(IDD_COMPONENT_DL, &m_tabCtrl);
	m_component_.m_ddic.SetWindowText(m_component_.ddicStr_);
	m_component_.m_cci.SetWindowText(m_component_.cciStr_);
	m_component_.m_fci.SetWindowText(m_component_.fciStr_);
	m_component_.MoveWindow(rc);
	m_component_.ShowWindow(SW_HIDE);

	m_hint_.Create(IDD_HINT, &m_tabCtrl);
	m_hint_.MoveWindow(rc);
	m_hint_.ShowWindow(SW_HIDE);

	CRect settingRect;
	m_setting.GetWindowRect(&settingRect);
	ScreenToClient(&settingRect);
	m_information.GetWindowRect(&m_information_rect);
	ScreenToClient(&m_information_rect);
	m_nMinWindowWidth = m_information_rect.right;

	m_nMinWindowHeight = settingRect.bottom - settingRect.top;

	m_re_check.GetWindowRect(&m_re_check_rect);
	ScreenToClient(&m_re_check_rect);
	m_pass.GetWindowRect(&m_pass_rect);
	ScreenToClient(&m_pass_rect);
	m_reInfo.GetWindowRect(&m_reInfo_rect);
	ScreenToClient(&m_reInfo_rect);
	CRect rect;
	GetClientRect(&rect);
	m_orig_cx = rect.Width();
	m_orig_cy = rect.Height();
	m_is_init = true;

	m_totalHeight = 0;
	m_displayedHeight = 0;

	// 初始化进度条
	m_progressCheck.SetRange(0, 100);
	m_progressCheck.SetPos(0);
	m_progressCheck.ShowWindow(SW_HIDE); // 初始隐藏

	return TRUE;
}

void CPinmapCheckToolDlg::InitGeneralComboBox()
{
	m_general_.Create(IDD_GENERAL, &m_tabCtrl);
	m_general_.m_cmbUser.AddString(_T("Customer0"));
	m_general_.m_cmbUser.SetCurSel(0);
	m_general_.m_project_code.AddString(_T("EPD8818"));
	m_general_.m_project_code.AddString(_T("EPD8819"));
	m_general_.m_project_code.AddString(_T("EPD8820"));
	m_general_.m_project_code.AddString(_T("EPD8828"));
	m_general_.m_project_code.AddString(_T("EPD8829"));
	m_general_.m_project_code.AddString(_T("EPD8830"));
	m_general_.m_project_code.AddString(_T("EPD8831"));
	m_general_.m_project_code.AddString(_T("EPD8831A"));
	m_general_.m_project_code.AddString(_T("EPD8835"));
	m_general_.m_project_code.AddString(_T("EPD8836"));
	m_general_.m_project_code.SetCurSel(0);

	m_general_.m_vgl.AddString(_T("Internal"));
	m_general_.m_vgl.AddString(_T("External"));
	m_general_.m_vgl.SetCurSel(0);

	m_general_.m_phy.AddString(_T("DPHY"));
	m_general_.m_phy.AddString(_T("CPHY"));
	m_general_.m_phy.SetCurSel(0);

	m_general_.m_power_mode.AddString(_T("3 power mode"));
	m_general_.m_power_mode.AddString(_T("3 power mode+VCC_VDDI"));
	m_general_.m_power_mode.AddString(_T("4 power mode"));
	m_general_.m_power_mode.SetCurSel(0);

	m_general_.m_flash_voltage.AddString(_T("1.8V"));
	m_general_.m_flash_voltage.AddString(_T("1.2V"));
	m_general_.m_flash_voltage.SetCurSel(0);

	m_general_.m_mipi_lane.AddString(_T("1"));
	m_general_.m_mipi_lane.AddString(_T("2"));
	m_general_.m_mipi_lane.AddString(_T("3"));
	m_general_.m_mipi_lane.AddString(_T("4"));
	m_general_.m_mipi_lane.SetCurSel(0);

	m_general_.m_pmic.AddString(_T("No"));
	m_general_.m_pmic.AddString(_T("Yes"));
	m_general_.m_pmic.SetCurSel(0);

	m_general_.m_pcd2.AddString(_T("No"));
	m_general_.m_pcd2.AddString(_T("Yes"));
	m_general_.m_pcd2.SetCurSel(0);

	m_general_.m_finger.AddString(_T("No"));
	m_general_.m_finger.AddString(_T("Yes"));
	m_general_.m_finger.SetCurSel(0);

	m_general_.m_mtp.AddString(_T("Internal"));
	m_general_.m_mtp.AddString(_T("External"));
	m_general_.m_mtp.SetCurSel(0);

	CRect rc;
	m_tabCtrl.GetClientRect(&rc);
	rc.DeflateRect(1, 22, 3, 1);
	m_general_.MoveWindow(rc);
	m_general_.ShowWindow(SW_HIDE);
	return;
}

int CPinmapCheckToolDlg::ColNameToIndex(const CString& colName)
{
	if (colName.IsEmpty()) return -1;
	CString upperCol = colName;
	upperCol.MakeUpper();  // 转为大写字母
	int index = 0;
	for (int i = 0; i < upperCol.GetLength(); i++) {
		TCHAR c = upperCol[i];
		if (c < 'A' || c > 'Z') return -1;  // 非法字符
		index = index * 26 + (c - 'A' + 1);
	}
	return index;
}

void CPinmapCheckToolDlg::OnBnClickedSelPinmap()
{
	// 定义过滤器
	CString filter = _T("文本文件 (*.xlsx;*.xls)|*.xlsx;*.xls|");

	// 创建打开文件对话框
	CFileDialog dlg(TRUE, _T("txt"), NULL,
		OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
		filter, this);

	// 显示对话框并处理结果
	if (dlg.DoModal() == IDOK) {
		CString fileName = dlg.GetFileName();
		m_staticPinmapName.SetWindowText(fileName);
		m_strPinmapPath = dlg.GetPathName();  // 保存完整路径
		m_btnSelPinmap.SetFaceColor(RGB(255, 255, 0));
		m_btnSelPinmap_status = true;
		MessageBox(_T("Pinmap文件已选择"));
	}
}

void CPinmapCheckToolDlg::OnBnClickedLoadNetlist()
{
	// 定义过滤器
	CString filter = _T("文本文件 (*.txt)|*.txt|");

	// 创建打开文件对话框
	CFileDialog dlg(TRUE, _T("txt"), NULL,
		OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
		filter, this);

	// 显示对话框并处理结果
	if (dlg.DoModal() == IDOK) {
		CString fileName = dlg.GetFileName();
		m_netList.SetWindowText(fileName);
		m_netListPath = dlg.GetPathName();  // 保存完整路径
		m_btnLoadNetlist.SetFaceColor(RGB(255, 255, 0));
		m_btnLoadNetlist_status = true;
		MessageBox(_T("Finished"));
	}
}

void CPinmapCheckToolDlg::OnBnClickedLoadBom()
{
	// 定义过滤器
	CString filter = _T("文本文件 (*.xlsx;*.xls)|*.xlsx;*.xls|");

	// 创建打开文件对话框
	CFileDialog dlg(TRUE, _T("txt"), NULL,
		OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
		filter, this);

	// 显示对话框并处理结果
	if (dlg.DoModal() == IDOK) {
		CString fileName = dlg.GetFileName();
		m_bom.SetWindowText(fileName);
		m_bomPath = dlg.GetPathName();  // 保存完整路径
		m_btnLoadBOM.SetFaceColor(RGB(255, 255, 0));
		m_btnLoadBOM_status = true;
		MessageBox(_T("Finished"));
	}
}

void CPinmapCheckToolDlg::OnBnClickedLoadCfg()
{
	// 定义过滤器
	CString filter = _T("文本文件 (*.xlsx;*.xls)|*.xlsx;*.xls|");

	// 创建打开文件对话框
	CFileDialog dlg(TRUE, _T("txt"), NULL,
		OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
		filter, this);

	// 显示对话框并处理结果
	if (dlg.DoModal() == IDOK) {
		m_loadCfgPath = dlg.GetPathName(); // 获取完整路径
		m_btnLoadCfg_status = true;

		CApplication app;
        CWorkbooks books;
        CWorkbook book;
        CWorksheets sheets;
        CWorksheet sheet;
        CRange range;
		COleVariant vtMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

        if (!app.CreateDispatch(_T("Excel.Application"))) {
            AfxMessageBox(_T("无法启动Excel！"));
            return;
        }
        app.put_Visible(FALSE);

        books.AttachDispatch(app.get_Workbooks());
        if (books.m_lpDispatch == NULL) {
            AfxMessageBox(_T("无法获取工作簿集合！"));
            app.Quit();
			app.ReleaseDispatch();
            return;
        }

		// 打开指定Excel文件（传递默认参数）
		book.AttachDispatch(books.Open(m_loadCfgPath, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing));
        if (book.m_lpDispatch == NULL) {
            AfxMessageBox(_T("无法打开文件：") + m_loadCfgPath);
            books.ReleaseDispatch();
            app.Quit();
			app.ReleaseDispatch();
            return;
        }

        // 获取工作表集合
        sheets.AttachDispatch(book.get_Worksheets());
        if (sheets.m_lpDispatch == NULL) {
            AfxMessageBox(_T("无法获取工作表集合！"));
            book.Close(vtMissing, vtMissing, vtMissing);
            books.ReleaseDispatch();
            app.Quit();
			app.ReleaseDispatch();
            return;
        }

		long sheetCount = sheets.get_Count();
		if (sheetCount < 3) {
			AfxMessageBox(_T("文件中Sheet页数量不足3个，无法加载第3个Sheet页！"));
			// 释放Excel资源并退出
			book.Close(vtMissing, vtMissing, vtMissing);
			sheets.ReleaseDispatch();
			books.ReleaseDispatch();
			app.Quit();
			app.ReleaseDispatch();
			return;
		}

        // 获取第3个Sheet页（索引从1开始）
        sheet.AttachDispatch(sheets.get_Item(COleVariant((long)3)));
        if (sheet.m_lpDispatch == NULL) {
            AfxMessageBox(_T("无法获取第3个Sheet页！"));
            book.Close(vtMissing, vtMissing, vtMissing);
            sheets.ReleaseDispatch();
            books.ReleaseDispatch();
            app.Quit();
			app.ReleaseDispatch();
            return;
        }

		CStringArray arrStandardFields;
		arrStandardFields.Add(_T("SHEET"));
		arrStandardFields.Add(_T("FOP"));
		arrStandardFields.Add(_T("PANEL_SIDE"));
		arrStandardFields.Add(_T("IC_OUT_SIDE"));
		arrStandardFields.Add(_T("IC_SIDE"));
		arrStandardFields.Add(_T("OLB_SIDE"));
		arrStandardFields.Add(_T("IC_UP_BOUNDARY"));
		arrStandardFields.Add(_T("IC_DOWN_BOUNDARY"));

		// 2. 读取A列字段（A2-A9，共8行）
		CStringArray arrActualFields;
		bool bReadFieldError = false;
		for (int i = 2; i <= 9; i++) {
			CString cellAddr;
			cellAddr.Format(_T("A%d"), i);
			range.AttachDispatch(sheet.get_Range(COleVariant(cellAddr), COleVariant(cellAddr)));
			if (range.m_lpDispatch == NULL) { bReadFieldError = true; break; }

			COleVariant varField = range.get_Value(vtMissing);
			range.ReleaseDispatch();

			CString strField;
			if (varField.vt == VT_BSTR) {
				strField = CString(varField.bstrVal).Trim().MakeUpper(); // 去空格、转大写
			} else if (varField.vt == VT_EMPTY) {
				strField = _T(""); // 空字段
			} else {
				bReadFieldError = true; // 非字符串/空类型，判定错误
			}

			if (strField.IsEmpty() || bReadFieldError) {
				bReadFieldError = true;
				break;
			}
			arrActualFields.Add(strField);
		}

		// 3. 校验字段数量与内容
		if (bReadFieldError || arrActualFields.GetSize() != 8) {
			AfxMessageBox(_T("文件错误：字段数量异常或存在空字段！"));
			book.Close(vtMissing, vtMissing, vtMissing);
			sheets.ReleaseDispatch();
			books.ReleaseDispatch();
			app.Quit();
			app.ReleaseDispatch();
			return;
		}

		bool bFieldMatch = true;
		for (int i = 0; i < arrStandardFields.GetSize(); i++) {
			CString strStandard = arrStandardFields[i].MakeUpper();
			CString strActual = arrActualFields[i];
			if (strActual != strStandard) {
				bFieldMatch = false;
				break;
			}
		}

		if (!bFieldMatch) {
			AfxMessageBox(_T("文件错误：字段不匹配！"));
			book.Close(vtMissing, vtMissing, vtMissing);
			sheets.ReleaseDispatch();
			books.ReleaseDispatch();
			app.Quit();
			app.ReleaseDispatch();
			return;
		}

        // 封装单元格值获取函数（避免重复代码）
        auto GetCellValue = [&](int row, int col) -> CString {
			CString cellAddrA;
			cellAddrA.Format(_T("A%d"), row);
			CString cellAddrB;
			cellAddrB.Format(_T("B%d"), row);
            range.AttachDispatch(sheet.get_Range(
				COleVariant(cellAddrA),
				COleVariant(cellAddrB)
            ));
            if (range.m_lpDispatch == NULL) return _T("");

			COleVariant varCell = range.get_Item(COleVariant((long)1), COleVariant((long)2));
			if (varCell.vt != VT_DISPATCH || varCell.pdispVal == NULL) {
				return _T("");
			}
			CRange cell;
			cell.AttachDispatch(varCell.pdispVal);
			COleVariant var = cell.get_Value(vtMissing);
			cell.ReleaseDispatch();

			CString strVal;
			if (var.vt == VT_EMPTY || var.vt == VT_NULL)
				strVal = _T("");
			else if (var.vt == VT_BSTR)
				strVal = var.bstrVal;
			else if (var.vt == VT_R8)
				strVal.Format(_T("%.0f"), var.dblVal);  // 数字去小数
			else if (var.vt == VT_I4)
				strVal.Format(_T("%d"), var.lVal);
			else
				strVal = _T("(未知类型)");

			return strVal;
        };

        // 1. 读取SHEET字段（A2-B2）
        sheetNumStr_ = GetCellValue(2, 2);
        m_pinmapinfo_.m_sheet_num.SetWindowText(sheetNumStr_);

        // 2. 读取FOP字段（A3-B3）
        fopStr_ = GetCellValue(3, 2);
        m_pinmapinfo_.m_fop.SetWindowText(fopStr_);

        // 3. 读取PANEL_SIE字段（A4-B4）
        glassRailingStr_ = GetCellValue(4, 2);
        m_pinmapinfo_.m_glass_railing.SetWindowText(glassRailingStr_);

        // 4. 读取IC_OUT_SIE字段（A5-B5）
        pinColStr_ = GetCellValue(5, 2);
        m_pinmapinfo_.m_pin.SetWindowText(pinColStr_);

        // 5. 读取IC_SIDE字段（A6-B6）
        icSideColStr_ = GetCellValue(6, 2);
        m_pinmapinfo_.m_ic_side.SetWindowText(icSideColStr_);

        // 6. 读取OLB_SIDE AC字段（A7-B7）
        olbStr_ = GetCellValue(7, 2);
        m_pinmapinfo_.m_olb.SetWindowText(olbStr_);

        // 7. 读取IC_UP_BO127字段（A8-B8）
        topRowStr_ = GetCellValue(8, 2);
        m_pinmapinfo_.m_ic_side_top.SetWindowTextW(topRowStr_);

        // 8. 读取IC_DOWN846字段（A9-B9）
        bottomRowStr_ = GetCellValue(9, 2);
        m_pinmapinfo_.m_ic_side_bottom.SetWindowTextW(bottomRowStr_);

        // 释放Excel资源
        book.Close(vtMissing, vtMissing, vtMissing);
        book.ReleaseDispatch();
        sheets.ReleaseDispatch();
        books.ReleaseDispatch();
        app.Quit();
        app.ReleaseDispatch();

		m_btnLoadCfg.SetFaceColor(RGB(255, 255, 0));
		MessageBox(_T("Finished"));
	}
}

void CPinmapCheckToolDlg::OnBnClickedSaveCfg()
{
	// 定义文件过滤器
	CString filter = _T("文本文件 (*.xlsx;*.xls)|*.xlsx;*.xls|");

	// 创建保存文件对话框（第一个参数为FALSE表示保存模式）
	CFileDialog dlg(FALSE, _T("xlsx"), _T("Config.xlsx"),
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		filter, this);

	// 显示对话框
	if (dlg.DoModal() == IDOK) {
		CString filePath = dlg.GetPathName(); // 获取完整路径

		// 初始化Excel应用
		CApplication excelApp;
		CWorkbooks workbooks;
		CWorkbook workbook;
		CWorksheets worksheets;
		CWorksheet sheet;

		// 创建Excel应用实例
		if (!excelApp.CreateDispatch(_T("Excel.Application"))) {
			AfxMessageBox(_T("无法启动Excel！"));
			return;
		}
		excelApp.put_Visible(FALSE); // 隐藏窗口

		// 获取工作簿集合
		workbooks.AttachDispatch(excelApp.get_Workbooks());
		if (workbooks.m_lpDispatch == NULL) {
			AfxMessageBox(_T("无法获取工作簿集合！"));
			excelApp.Quit();
			excelApp.ReleaseDispatch();
			return;
		}

		// 创建新工作簿
		COleVariant vMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		workbook.AttachDispatch(workbooks.Add(vMissing));
		if (workbook.m_lpDispatch == NULL) {
			AfxMessageBox(_T("无法创建新工作簿！"));
			workbooks.ReleaseDispatch();
			excelApp.Quit();
			excelApp.ReleaseDispatch();
			return;
		}

		// 获取工作表集合
		worksheets.AttachDispatch(workbook.get_Worksheets());
		if (worksheets.m_lpDispatch == NULL) {
			AfxMessageBox(_T("无法获取工作表集合！"));
			workbook.Close(vMissing, vMissing, vMissing);
			worksheets.ReleaseDispatch();
			workbooks.ReleaseDispatch();
			excelApp.Quit();
			excelApp.ReleaseDispatch();
			return;
		}

		SavePinMapSheet(sheet, worksheets, workbook, workbooks, excelApp);
		sheet.ReleaseDispatch();
		SaveComponentSheet(sheet, worksheets, workbook, workbooks, excelApp);
		sheet.ReleaseDispatch();
		SaveGeneralSheet(sheet, worksheets, workbook, workbooks, excelApp);


		// 保存工作簿
		COleVariant vFilePath((LPCTSTR)filePath), vFormat((long)51); // 51代表xlsx格式
		workbook.SaveAs(vFilePath, vFormat, vMissing, vMissing, vMissing, vMissing, 1,
			vMissing, vMissing, vMissing, vMissing, vMissing, vMissing);

		// 关闭工作簿和Excel应用
		workbook.Close(vMissing, vMissing, vMissing);
		sheet.ReleaseDispatch();
		worksheets.ReleaseDispatch();
		workbook.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();

		AfxMessageBox(_T("配置文件保存成功！"));
	}
}

void CPinmapCheckToolDlg::SaveGeneralSheet(CWorksheet& sheet, CWorksheets& worksheets, CWorkbook& workbook, CWorkbooks& workbooks, CApplication& excelApp)
{
	COleVariant vMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	// 获取第一个工作表并命名为GENERAL
	sheet.AttachDispatch(worksheets.get_Item(COleVariant((long)1)));
	if (sheet.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法获取工作表！"));
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return;
	}
	sheet.put_Name(_T("GENERAL"));

	// 设置第一行标题
	CRange cell;
	// A1单元格设置为Key
	cell.AttachDispatch(sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("A1"))));
	if (cell.m_lpDispatch != NULL) {
		cell.put_Value2(COleVariant(_T("Key")));
		cell.ReleaseDispatch();
	}
	// B1单元格设置为Value
	cell.AttachDispatch(sheet.get_Range(COleVariant(_T("B1")), COleVariant(_T("B1"))));
	if (cell.m_lpDispatch != NULL) {
		cell.put_Value2(COleVariant(_T("Value")));
		cell.ReleaseDispatch();
	}

	// 定义要保存的控件和对应的Key名称
	struct ControlKeyPair {
		CComboBox* pComboBox;
		CString keyName;
	};

	ControlKeyPair controlPairs[] = {
		{ &m_general_.m_cmbUser, _T("User") },
		{ &m_general_.m_project_code, _T("Project Code") },
		{ &m_general_.m_vgl, _T("VGL外灌？") },
		{ &m_general_.m_phy, _T("PHY") },
		{ &m_general_.m_power_mode, _T("Power Mode") },
		{ &m_general_.m_flash_voltage, _T("Flash Voltage") },
		{ &m_general_.m_mipi_lane, _T("MIPI Lane") },
		{ &m_general_.m_pmic, _T("I2C control PMIC") },
		{ &m_general_.m_pcd2, _T("使用PCD2？") },
		{ &m_general_.m_finger, _T("是否使用指纹？") },
		{ &m_general_.m_mtp, _T("MTP power外灌？") }
	};

	int pairCount = sizeof(controlPairs) / sizeof(ControlKeyPair);

	// 写入各个控件的Key和Value
	for (int i = 0; i < pairCount; i++) {
		int row = i + 2; // 从第二行开始

		// 获取Key单元格
		CString keyCell; 
		keyCell.Format(_T("A%d"), row);
		cell.AttachDispatch(sheet.get_Range(COleVariant(keyCell), COleVariant(keyCell)));
		if (cell.m_lpDispatch != NULL) {
			cell.put_Value2(COleVariant(controlPairs[i].keyName));
			cell.ReleaseDispatch();
		}

		// 获取Value单元格
		CString valueCell;
		valueCell.Format(_T("B%d"), row);
		cell.AttachDispatch(sheet.get_Range(COleVariant(valueCell), COleVariant(valueCell)));
		if (cell.m_lpDispatch != NULL) {
			// 获取当前选择的文本
			CString value;
			int curSel = controlPairs[i].pComboBox->GetCurSel();
			if (curSel != CB_ERR) {
				controlPairs[i].pComboBox->GetLBText(curSel, value);
			}
			cell.put_Value2(COleVariant(value));
			cell.ReleaseDispatch();
		}
	}
}

void CPinmapCheckToolDlg::SaveComponentSheet(CWorksheet& sheet, CWorksheets& worksheets, CWorkbook& workbook, CWorkbooks& workbooks, CApplication& excelApp)
{
	COleVariant vMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	// 获取第二个工作表并命名为COMPONENT
	int sheetCount = worksheets.get_Count();
	if (sheetCount < 2) {
		// 如果工作表数量不足2个，创建一个新的工作表
		sheet.AttachDispatch(worksheets.Add(vMissing, vMissing, vMissing, vMissing));
		sheet.AttachDispatch(worksheets.get_Item(COleVariant((long)2)));
	} else {
		// 否则获取第二个工作表
		sheet.AttachDispatch(worksheets.get_Item(COleVariant((long)2)));
	}
	
	if (sheet.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法获取工作表！"));
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return;
	}
	sheet.put_Name(_T("COMPONENT"));

	// 设置第一行标题
	CRange cell;
	// A1单元格设置为Key
	cell.AttachDispatch(sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("A1"))));
	if (cell.m_lpDispatch != NULL) {
		cell.put_Value2(COleVariant(_T("Key")));
		cell.ReleaseDispatch();
	}
	// B1单元格设置为Value
	cell.AttachDispatch(sheet.get_Range(COleVariant(_T("B1")), COleVariant(_T("B1"))));
	if (cell.m_lpDispatch != NULL) {
		cell.put_Value2(COleVariant(_T("Value")));
		cell.ReleaseDispatch();
	}

	// 定义要保存的控件和对应的Key名称
	struct ComponentDataPair {
		CString keyName;
		CString value;
	};

	ComponentDataPair componentData[] = {
		{ _T("DDIC"), m_component_.ddicStr_ },
		{ _T("CCI"), m_component_.cciStr_ },
		{ _T("FCI"), m_component_.fciStr_ }
	};

	int dataCount = sizeof(componentData) / sizeof(ComponentDataPair);

	// 写入各个组件数据的Key和Value
	for (int i = 0; i < dataCount; i++) {
		int row = i + 2; // 从第二行开始

		// 获取Key单元格
		CString keyCell;
		keyCell.Format(_T("A%d"), row);
		cell.AttachDispatch(sheet.get_Range(COleVariant(keyCell), COleVariant(keyCell)));
		if (cell.m_lpDispatch != NULL) {
			cell.put_Value2(COleVariant(componentData[i].keyName));
			cell.ReleaseDispatch();
		}

		// 获取Value单元格
		CString valueCell;
		valueCell.Format(_T("B%d"), row);
		cell.AttachDispatch(sheet.get_Range(COleVariant(valueCell), COleVariant(valueCell)));
		if (cell.m_lpDispatch != NULL) {
			cell.put_Value2(COleVariant(componentData[i].value));
			cell.ReleaseDispatch();
		}
	}
}

void CPinmapCheckToolDlg::SavePinMapSheet(CWorksheet& sheet, CWorksheets& worksheets, CWorkbook& workbook, CWorkbooks& workbooks, CApplication& excelApp)
{
	COleVariant vMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	// 获取第三个工作表并命名为PIN_MAP
	int sheetCount = worksheets.get_Count();
	if (sheetCount < 3) {
		// 如果工作表数量不足3个，创建新的工作表
		for (int i = sheetCount; i < 3; i++) {
			sheet.AttachDispatch(worksheets.Add(vMissing, vMissing, vMissing, vMissing));
			if (sheet.m_lpDispatch != NULL) {
				sheet.ReleaseDispatch();
			}
		}
	}
	// 获取第三个工作表
	sheet.AttachDispatch(worksheets.get_Item(COleVariant((long)3)));
	
	if (sheet.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法获取工作表！"));
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return;
	}
	sheet.put_Name(_T("PIN_MAP"));

	// 设置第一行标题
	CRange cell;
	// A1单元格设置为Key
	cell.AttachDispatch(sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("A1"))));
	if (cell.m_lpDispatch != NULL) {
		cell.put_Value2(COleVariant(_T("Key")));
		cell.ReleaseDispatch();
	}
	// B1单元格设置为Value
	cell.AttachDispatch(sheet.get_Range(COleVariant(_T("B1")), COleVariant(_T("B1"))));
	if (cell.m_lpDispatch != NULL) {
		cell.put_Value2(COleVariant(_T("Value")));
		cell.ReleaseDispatch();
	}

	// 定义要保存的控件和对应的Key名称
	struct PinMapDataPair {
		CString keyName;
		CString value;
	};

	PinMapDataPair pinMapData[] = {
		{ _T("Sheet Number"), sheetNumStr_ },
		{ _T("Pin Column"), pinColStr_ },
		{ _T("IC Side Column"), icSideColStr_ },
		{ _T("Top Row"), topRowStr_ },
		{ _T("Bottom Row"), bottomRowStr_ },
		{ _T("FOP"), fopStr_ },
		{ _T("Glass Railing"), glassRailingStr_ },
		{ _T("OLB"), olbStr_ }
	};

	int dataCount = sizeof(pinMapData) / sizeof(PinMapDataPair);

	// 写入各个组件数据的Key和Value
	for (int i = 0; i < dataCount; i++) {
		int row = i + 2; // 从第二行开始

		// 获取Key单元格
		CString keyCell;
		keyCell.Format(_T("A%d"), row);
		cell.AttachDispatch(sheet.get_Range(COleVariant(keyCell), COleVariant(keyCell)));
		if (cell.m_lpDispatch != NULL) {
			cell.put_Value2(COleVariant(pinMapData[i].keyName));
			cell.ReleaseDispatch();
		}

		// 获取Value单元格
		CString valueCell;
		valueCell.Format(_T("B%d"), row);
		cell.AttachDispatch(sheet.get_Range(COleVariant(valueCell), COleVariant(valueCell)));
		if (cell.m_lpDispatch != NULL) {
			cell.put_Value2(COleVariant(pinMapData[i].value));
			cell.ReleaseDispatch();
		}
	}
}

BOOL CPinmapCheckToolDlg::ReadExcelColumn(const CString& strFilePath, int nSheetIndex, int nColIndex, vector<CString>& data)
{
	CApplication excelApp;    // Excel应用实例
	CWorkbooks workbooks;     // 工作簿集合
	CWorkbook workbook;       // 单个工作簿
	CWorksheets worksheets;    // 工作表集合
	CWorksheet sheet;         // 单个工作表
	CRange usedRange;         // 有效数据范围
	CRange rows;              // 行集合
	CRange cell;              // 单个单元格

	// 1. 初始化Excel应用
	if (!excelApp.CreateDispatch(_T("Excel.Application"))) {
		AfxMessageBox(_T("无法启动Excel！"));
		return FALSE;
	}
	excelApp.put_Visible(FALSE);  // 隐藏窗口

	// 2. 获取工作簿集合
	workbooks.AttachDispatch(excelApp.get_Workbooks());
	if (workbooks.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法获取工作簿集合！"));
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return FALSE;
	}

	// 3. 打开Excel文件
	COleVariant vFilePath(strFilePath), vReadOnly((short)TRUE), vMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	workbook.AttachDispatch(workbooks.Open(strFilePath, vMissing, vReadOnly,
		vMissing, vMissing, vMissing, vMissing, vMissing,
		vMissing, vMissing, vMissing, vMissing, vMissing, vMissing, vMissing));
	if (workbook.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法打开文件：") + strFilePath);
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return FALSE;
	}

	// 4. 获取工作表集合
	worksheets.AttachDispatch(workbook.get_Worksheets());
	if (worksheets.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法获取工作表集合！"));
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return FALSE;
	}

	// 5. 检查Sheet索引有效性
	int sheetCount = worksheets.get_Count();
	if (nSheetIndex + 1 > sheetCount) {
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return FALSE;
	}

	// 6. 获取指定Sheet
	sheet.AttachDispatch(worksheets.get_Item(COleVariant((long)(nSheetIndex + 1))));  // 1基索引
	if (sheet.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法获取指定Sheet！"));
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return FALSE;
	}

	// 7. 获取有效数据范围
	usedRange.AttachDispatch(sheet.get_UsedRange());
	if (usedRange.m_lpDispatch == NULL) {
		AfxMessageBox(_T("无法获取数据范围！"));
		sheet.ReleaseDispatch();
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return FALSE;
	}

	// 8. 获取总行数
	rows.AttachDispatch(usedRange.get_Rows());
	int rowCount = rows.get_Count();
	if (rowCount <= 0) {
		AfxMessageBox(_T("Sheet页无数据！"));
		rows.ReleaseDispatch();
		usedRange.ReleaseDispatch();
		sheet.ReleaseDispatch();
		workbook.Close(vMissing, vMissing, vMissing);
		worksheets.ReleaseDispatch();
		workbooks.ReleaseDispatch();
		excelApp.Quit();
		excelApp.ReleaseDispatch();
		return FALSE;
	}

	// 9. 读取指定列数据
	data.clear();
	for (int row = 1; row <= rowCount; row++) {
		COleVariant varCell;
		varCell = usedRange.get_Item(COleVariant((long)row), COleVariant((long)nColIndex));
		// 检查返回值是否为有效的IDispatch指针
		if (varCell.vt != VT_DISPATCH || varCell.pdispVal == NULL) {
			data.push_back(_T("(无效单元格)"));
			continue;
		}
		cell.AttachDispatch(varCell.pdispVal);
		if (cell.m_lpDispatch == NULL) {
			data.push_back(_T("(无效单元格)"));
			continue;
		}

		COleVariant vMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  // 定义默认参数
		COleVariant vValue = cell.get_Value(vMissing);
		CString strVal;
		if (vValue.vt == VT_EMPTY || vValue.vt == VT_NULL)
			strVal = _T("");
		else if (vValue.vt == VT_BSTR)
			strVal = vValue.bstrVal;
		else if (vValue.vt == VT_R8)
			strVal.Format(_T("%.0f"), vValue.dblVal);  // 数字去小数
		else if (vValue.vt == VT_I4)
			strVal.Format(_T("%d"), vValue.lVal);
		else
			strVal = _T("(未知类型)");

		data.push_back(strVal);
		cell.ReleaseDispatch();  // 释放单元格对象
	}

	// 10. 清理资源
	rows.ReleaseDispatch();
	usedRange.ReleaseDispatch();
	sheet.ReleaseDispatch();
	workbook.Close(vMissing, vMissing, vMissing);
	workbook.ReleaseDispatch();
	worksheets.ReleaseDispatch();
	workbooks.ReleaseDispatch();
	excelApp.Quit();
	excelApp.ReleaseDispatch();

	return TRUE;
}

// 解析Flash引脚信息的辅助函数
map<int, CString> ParseFlashAndConnectorPins(const CString& flashLine)
{
	map<int, CString> flashPinMap;
	
	// 提取括号内的内容（去除首尾括号）
	int start = flashLine.Find(_T("(")) + 1;
	int end = flashLine.ReverseFind(_T(')'));
	if (start > 0 && end > start) {
		CString pinsStr = flashLine.Mid(start, end - start);

		CStringArray pinItems;
		int pos = 0;
		CString token;

		while ((token = pinsStr.Tokenize(_T(","), pos)) != _T("")) {
			token.Trim(); // 去除前后空格

			if (token.Left(1) == _T("(") && token.Right(1) == _T(")")) {
				token = token.Mid(1, token.GetLength() - 2); // 从第1个字符开始，截取长度-2的子串
			}

			if (!token.IsEmpty()) {
				int colonPos = token.Find(_T(":"));
				if (colonPos > 0) {
					CString numStr = token.Left(colonPos);
					CString nameStr = token.Mid(colonPos + 1);
					
					int pinNum = _ttoi(numStr);
					flashPinMap[pinNum] = nameStr;
				}
			}
		}
	}
	
	return flashPinMap;
}

void CPinmapCheckToolDlg::ProcessNetlistFile(vector<pair<int, CString>> validPairList)
{
	if (m_netListPath.IsEmpty()) {
		return;
	}
	CStdioFile netFile;
	if (!netFile.Open(m_netListPath, CFile::modeRead | CFile::typeText)) {
		AfxMessageBox(_T("网表文件读取失败：") + m_netListPath);
		return;
	}

	CString line;
	CString	ddicPrefix = m_component_.ddicStr_ + _T("(");
	CString flashPrefix = m_component_.fciStr_ + _T("(");
	CString connectorPrefix = m_component_.cciStr_ + _T("(");
	FopDataList netPinList;
	
	// 逐行读取网表，筛选DDIC ID和Flash ID并解析引脚
	while (netFile.ReadString(line)) {
		line.Trim();
		
		// 解析DDIC引脚
		if (line.Left(ddicPrefix.GetLength()) == ddicPrefix) {
			// 提取括号内的内容（去除首尾括号）
			int start = line.Find(_T("(")) + 1;
			int end = line.ReverseFind(_T(')'));
			if (start > 0 && end > start) {
				CString pinsStr = line.Mid(start, end - start);

				CStringArray pinItems;
				int pos = 0;
				CString token;

				while ((token = pinsStr.Tokenize(_T(","), pos)) != _T("")) {
					token.Trim(); // 去除前后空格，得到 "(1:1dummy)"

					if (token.Left(1) == _T("(") && token.Right(1) == _T(")")) {
						token = token.Mid(1, token.GetLength() - 2); // 从第1个字符开始，截取长度-2的子串
					}

					if (!token.IsEmpty()) {
						pinItems.Add(token); // 现在 token 是 "1:1dummy"
					}
				}

				// 解析每个引脚条目（如"490:490VGH" → 数值490，名称VGH）
				for (int i = 0; i < pinItems.GetSize(); i++) {
					CString item = pinItems[i].Trim();
					int colonPos = item.Find(_T(":"));
					if (colonPos > 0) {
						CString numStr = item.Left(colonPos);
						CString nameStr = item.Mid(colonPos + 1);

						// 去除nameStr前面的数字
						int nonDigitPos = 0;
						while (nonDigitPos < nameStr.GetLength() && _istdigit(nameStr[nonDigitPos])) {
							nonDigitPos++;
						}
						if (nonDigitPos < nameStr.GetLength()) {
							nameStr = nameStr.Mid(nonDigitPos); // 截取非数字部分作为纯名称
						}
						else {
							nameStr.Empty(); // 若全为数字，清空（可根据需求调整）
						}

						int num = _ttoi(numStr);
						netPinList.emplace_back(num, nameStr);
					}
				}
			}
		}
		// 解析Flash引脚
		else if (line.Left(flashPrefix.GetLength()) == flashPrefix) {
			flashPinMap = ParseFlashAndConnectorPins(line);
			
			// 输出调试信息
			CString debugMsg;
			debugMsg.Format(_T("[调试] 解析Flash引脚完成，共找到 %zu 个引脚\r\n"), flashPinMap.size());
			OutputDebugString(debugMsg);
			
			for (const auto& pair : flashPinMap) {
				CString pinInfo;
				pinInfo.Format(_T("[调试] flashPinMap Flash引脚 %d: %s\r\n"), pair.first, pair.second);
				OutputDebugString(pinInfo);
			}
		} else if (line.Left(connectorPrefix.GetLength()) == connectorPrefix) {
			connectorMap = ParseFlashAndConnectorPins(line);
		}
	}
	netFile.Close();

	for (const auto& netPin : netPinList) {
		// 检查是否已存在（按编号+名称去重）
		bool isDuplicate = false;
		for (const auto& uniquePin : uniqueNetPinList) {
			if (uniquePin.first == netPin.first && uniquePin.second.Compare(netPin.second) == 0) {
				isDuplicate = true;
				break;
			}
		}
		if (!isDuplicate) {
			uniqueNetPinList.push_back(netPin);
		}
	}

	// 对比网表引脚与筛选后的validFop/pinVal
	CString netPassMsg, netFailMsg;
	netPassMsg = _T("\r\n=== 网表引脚对比结果（PASS） ===\r\n");
	netFailMsg = _T("\r\n=== 网表引脚对比结果（不一致） ===\r\n");

	CString tmp;
	int netListPassNum = 0;
	map<pair<int, CString>, bool> excelPinMap;
	for (const auto& validPair : validPairList) {
		excelPinMap[{validPair.first, validPair.second}] = true;
	}

	int matchedCount = 0;
	for (const auto& netPin : uniqueNetPinList) {
		bool isMatch = false;

		if (excelPinMap.find({ netPin.first, netPin.second }) != excelPinMap.end()) {
			isMatch = true;
			matchedCount++;
		}
		if (isMatch)
		{
			netListPassNum++;
		}
		else
		{
			tmp.Format(_T("netPinList FOP NO.%d（%s）：与网表引脚不一致\r\n"), netPin.first, netPin.second);
			netFailMsg += tmp;
		}
	}
	if (netListPassNum == uniqueNetPinList.size() && netListPassNum == validPairList.size() - 1) {
		netPassMsg += _T("出Pin栏与netList引脚一致。\r\n");
	}
	else {
		AfxMessageBox(_T(""));
		tmp.Format(_T("出Pin栏对应的netList引脚数量与名称不一致\r\n"));
		netFailMsg += tmp;
	}

	passMsg += netPassMsg;
	failMsg += netFailMsg;
}

void CPinmapCheckToolDlg::ProcessBOMFile()
{
	if (!m_bomPath.IsEmpty()) {
		// 添加调试日志
		OutputDebugString(_T("[调试] 开始读取BOM文件...\r\n"));
		CString debugStr;
		debugStr.Format(_T("[调试] BOM文件路径: %s\r\n"), m_bomPath);
		OutputDebugString(debugStr);

		// 设置B列（Value）和C列（Ref Designator）的索引
		int bomValueIndex = ColNameToIndex(_T("B"));
		int bomRefDesignator = ColNameToIndex(_T("C"));
		if (bomValueIndex <= 0 || bomRefDesignator <= 0)
		{
			AfxMessageBox(_T("列名格式错误（请输入A-Z、AA-AZ等）！"));
			return;
		}

		// 读取B列和C列数据
		vector<CString> bomValueData, bomRefDesignatorData;
		if (!ReadExcelColumn(m_bomPath, 0, bomValueIndex, bomValueData))
		{
			m_reInfo.SetWindowText(_T("读取B列失败！"));
			return;
		}
		if (!ReadExcelColumn(m_bomPath, 0, bomRefDesignator, bomRefDesignatorData))
		{
			m_reInfo.SetWindowText(_T("读取C列失败！"));
			return;
		}

		// 输出读取结果到调试日志
		debugStr.Format(_T("[调试] B列数据数量: %zu, C列数据数量: %zu\r\n"), bomValueData.size(), bomRefDesignatorData.size());
		OutputDebugString(debugStr);

		// 确保数据数量匹配
		size_t minSize = min(bomValueData.size(), bomRefDesignatorData.size());
		for (size_t i = 0; i < minSize; i++)
		{
			CString value = bomValueData[i].Trim();
			CString refDesignator = bomRefDesignatorData[i].Trim();

			// 跳过空行
			if (value.IsEmpty() || refDesignator.IsEmpty())
				continue;

			// 解析Value提取电容和电压
			ComponentSpec spec;
			spec.originalValue = value;

			// 解析逻辑：提取所有电容单位（uF、nF、pF、F）和电压V值
			// 定义所有可能的电容单位
			CString units[] = { _T("uF"), _T("nF"), _T("pF"), _T("F") };
			CString foundUnit = _T("");
			int unitPos = -1;

			// 查找所有可能的电容单位
			for (int i = 0; i < 4; i++) {
				int pos = value.Find(units[i]);
				if (pos > 0) {
					foundUnit = units[i];
					unitPos = pos;
					break;
				}
			}

			int vPos = value.Find(_T("V"));

			// 提取电容
			if (!foundUnit.IsEmpty() && unitPos > 0) {
				// 提取单位前面的数字部分
				CString capStr;
				for (int j = unitPos - 1; j >= 0; j--) {
					TCHAR ch = value[j];
					if (_istdigit(ch) || ch == _T('.') || ch == _T('m') || ch == _T('u') || ch == _T('n') || ch == _T('p')) {
						capStr.Insert(0, ch);
					}
					else {
						break;
					}
				}
				// 确保单位正确添加
				spec.capacitance = capStr + foundUnit;
			}

			// 提取电压
			if (vPos > 0) {
				// 提取V前面的数字部分
				CString voltStr;
				for (int j = vPos - 1; j >= 0; j--) {
					TCHAR ch = value[j];
					if (_istdigit(ch) || ch == _T('.')) {
						voltStr.Insert(0, ch);
					}
					else {
						break;
					}
				}
				spec.voltage = voltStr + _T("V");
			}

			// 拆分Ref Designator（可能包含逗号分隔的多个值）
			int commaPos = 0;
			CString tempRef = refDesignator;
			while (commaPos != -1) {
				commaPos = tempRef.Find(_T(","));
				CString token;
				if (commaPos != -1) {
					token = tempRef.Left(commaPos).Trim();
					tempRef = tempRef.Mid(commaPos + 1).Trim();
				}
				else {
					token = tempRef.Trim();
				}
				if (!token.IsEmpty()) {
					spec.refDesignators.push_back(token);
					// 为每个Ref Designator创建映射
					bomRefToSpecMap[token] = spec;
				}
			}

			// 输出解析结果到调试日志
			debugStr.Format(_T("[调试] 解析结果 - Value: %s, 电容: %s, 电压: %s, Ref Designator: %s\r\n"),
				value, spec.capacitance, spec.voltage, refDesignator);
			OutputDebugString(debugStr);
		}

		// 输出map大小到调试日志
		debugStr.Format(_T("[调试] bomRefToSpecMap大小: %zu\r\n"), bomRefToSpecMap.size());
		OutputDebugString(debugStr);

		// 显示部分解析结果示例
		int count = 0;
		for (auto& pair : bomRefToSpecMap) {
			CString temp;
			temp.Format(_T("Ref: %s, 电容: %s, 电压: %s\r\n"),
				pair.first, pair.second.capacitance, pair.second.voltage);
			debugStr.Format(_T("[调试] lst Ref: %s, 电容: %s, 电压: %s\r\n"), pair.first, pair.second.capacitance, pair.second.voltage);
			OutputDebugString(debugStr);
		}
	}
}

void CPinmapCheckToolDlg::ProcessConfigFile(vector<CString> validPinData, vector<pair<int, CString>> validIcSideList, vector<CString> totalIcSideData)
{
	if (!m_loadCfgPath.IsEmpty()) {
		icSideSpecMap.clear(); // 清空之前的数据
		flashFopMap.clear(); // 清空之前的Flash FOP映射数据

		CApplication app;
		CWorkbooks books;
		CWorkbook book;
		CWorksheets sheets;
		CWorksheet sheet;
		CRange range;
		COleVariant vtMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

		if (app.CreateDispatch(_T("Excel.Application"))) {
			app.put_Visible(FALSE);

			books.AttachDispatch(app.get_Workbooks());
			if (books.m_lpDispatch != NULL) {
				// 打开config文件
				book.AttachDispatch(books.Open(m_loadCfgPath, vtMissing, vtMissing,
					vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
					vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
					vtMissing, vtMissing));
				if (book.m_lpDispatch != NULL) {
					sheets.AttachDispatch(book.get_Worksheets());
					if (sheets.m_lpDispatch != NULL) {
						long sheetCount = sheets.get_Count();
						if (sheetCount >= 6) {
							// 获取第6个Sheet页（索引从1开始）
							sheet.AttachDispatch(sheets.get_Item(COleVariant((long)6)));
							if (sheet.m_lpDispatch != NULL) {
								// 获取使用范围以避免读取过多空行
								CRange usedRange;
								usedRange.AttachDispatch(sheet.get_UsedRange());
								if (usedRange.m_lpDispatch != NULL) {
									CRange rowsRange;
									rowsRange.AttachDispatch(usedRange.get_Rows());
									COleVariant rowsCount;
									if (rowsRange.m_lpDispatch != NULL) {
										rowsCount = rowsRange.get_Count();
										rowsRange.ReleaseDispatch();
									}
									int maxRow = 1000; // 默认最大行数为1000
									if (rowsCount.vt == VT_I4)
										maxRow = rowsCount.lVal;
									else if (rowsCount.vt == VT_R8)
										maxRow = (int)rowsCount.dblVal;

									for (int i = 2; i <= maxRow; i++) {
										CString bCellAddr;
										bCellAddr.Format(_T("B%d"), i);
										CRange bCell;
										bCell.AttachDispatch(sheet.get_Range(COleVariant(bCellAddr), COleVariant(bCellAddr)));
										CString signalName;
										if (bCell.m_lpDispatch != NULL) {
											COleVariant bVal = bCell.get_Value(vtMissing);
											if (bVal.vt == VT_BSTR)
												signalName = bVal.bstrVal;
											bCell.ReleaseDispatch();
										}
										// 读取C列的值（电容/电压信息）
										CString cCellAddr;
										cCellAddr.Format(_T("C%d"), i);
										CRange cCell;
										cCell.AttachDispatch(sheet.get_Range(COleVariant(cCellAddr), COleVariant(cCellAddr)));
										CString componentValue;
										if (cCell.m_lpDispatch != NULL) {
											COleVariant cVal = cCell.get_Value(vtMissing);
											if (cVal.vt == VT_BSTR)
												componentValue = cVal.bstrVal;
											else if (cVal.vt == VT_R8)
												componentValue.Format(_T("%.2f"), cVal.dblVal);
											else if (cVal.vt == VT_I4)
												componentValue.Format(_T("%d"), cVal.lVal);
											cCell.ReleaseDispatch();
										}

										// 解析电容和电压
										if (!signalName.IsEmpty() && !componentValue.IsEmpty()) {
											ComponentSpec spec;
											spec.originalValue = componentValue;

											// 尝试解析格式如 "4.7uF/6.3V" 的值
											int slashPos = componentValue.Find('/');
											if (slashPos > 0) {
												// 提取电容部分（左侧）
												spec.capacitance = componentValue.Left(slashPos).Trim();
												// 提取电压部分（右侧）
												spec.voltage = componentValue.Mid(slashPos + 1).Trim();
											} else {
												// 定义所有可能的电容单位
												CString units[] = { _T("uF"), _T("nF"), _T("pF"), _T("F") };
												bool isCapacitance = false;

												// 检查是否包含任何电容单位
												for (int i = 0; i < 4; i++) {
													if (componentValue.Find(units[i]) >= 0) {
														isCapacitance = true;
														break;
													}
												}

												// 根据内容判断是电容还是电压
												if (isCapacitance) {
													spec.capacitance = componentValue;
													spec.voltage = _T("");
												}
												else if (componentValue.Find(_T("V")) >= 0) {
													spec.capacitance = _T("");
													spec.voltage = componentValue;
												}
											}

											// 检查信号名称是否包含逗号，如果包含则分割为多个独立的key
											int commaPos = signalName.Find(',');
											if (commaPos > 0) {
												// 分割信号名称并为每个部分创建单独的条目
												CString leftPart = signalName.Left(commaPos).Trim();
												CString rightPart = signalName.Mid(commaPos + 1).Trim();

												// 存储分割后的两个部分
												if (!leftPart.IsEmpty()) {
													icSideSpecMap[leftPart] = spec;
												}
												if (!rightPart.IsEmpty()) {
													icSideSpecMap[rightPart] = spec;
												}
											} else {
												// 信号名称不包含逗号，直接存储
												icSideSpecMap[signalName] = spec;
											}
										}
									}
									usedRange.ReleaseDispatch();
								}
							}
							sheet.ReleaseDispatch();
						}

						// 读取第7个Sheet页的Flash引脚与FOP对应关系
						CString debugMsg;
						debugMsg.Format(_T("[调试] 开始读取第7个Sheet页，sheetCount = %ld\r\n"), sheetCount);
						OutputDebugString(debugMsg);
						if (sheetCount >= 7) {
							// 获取第7个Sheet页（索引从1开始）
							sheet.AttachDispatch(sheets.get_Item(COleVariant((long)7)));
							if (sheet.m_lpDispatch != NULL) {
								// 输出当前操作的Sheet名称
								COleVariant sheetNameVar = sheet.get_Name();
								CString sheetName;
								if (sheetNameVar.vt == VT_BSTR) {
									sheetName = sheetNameVar.bstrVal;
								}
								debugMsg.Format(_T("[调试] 成功获取第7个Sheet页，名称为: %s\r\n"), sheetName);
								OutputDebugString(debugMsg);

								// 获取使用范围以避免读取过多空行
								CRange usedRange;
								usedRange.AttachDispatch(sheet.get_UsedRange());
								if (usedRange.m_lpDispatch != NULL) {
									CRange rowsRange;
									rowsRange.AttachDispatch(usedRange.get_Rows());
									COleVariant rowsCount;
									if (rowsRange.m_lpDispatch != NULL) {
										rowsCount = rowsRange.get_Count();
										rowsRange.ReleaseDispatch();
									}
									int maxRow = 1000; // 默认最大行数为1000
									if (rowsCount.vt == VT_I4) {
										maxRow = rowsCount.lVal;
										debugMsg.Format(_T("[调试] 检测到有效行数 (VT_I4): %d\r\n"), maxRow);
										OutputDebugString(debugMsg);
									} else if (rowsCount.vt == VT_R8) {
										maxRow = (int)rowsCount.dblVal;
										debugMsg.Format(_T("[调试] 检测到有效行数 (VT_R8): %d\r\n"), maxRow);
										OutputDebugString(debugMsg);
									} else {
										debugMsg.Format(_T("[调试] 无法确定有效行数，使用默认值: %d\r\n"), maxRow);
										OutputDebugString(debugMsg);
									}

									// 重置flashFopMap
									flashFopMap.clear();

									// 遍历行，从第2行开始（假设第1行是标题）
									for (int i = 2; i <= maxRow; i++) {
										CString aCellAddr, bCellAddr;
										// 读取A列的Flash引脚名称
										aCellAddr.Format(_T("A%d"), i);
										CRange aCell;
										aCell.AttachDispatch(sheet.get_Range(COleVariant(aCellAddr), COleVariant(aCellAddr)));
										CString flashPin;
										if (aCell.m_lpDispatch != NULL) {
											COleVariant aVal = aCell.get_Value(vtMissing);
											// 处理多种可能的数据类型
											if (aVal.vt == VT_BSTR) {
												flashPin = aVal.bstrVal;
											} else if (aVal.vt == VT_R8) {
												flashPin.Format(_T("%.0f"), aVal.dblVal);
											} else if (aVal.vt == VT_I4) {
												flashPin.Format(_T("%d"), aVal.lVal);
											} else {
												flashPin.Empty();
											}
											flashPin.Trim(); // 去除前后空格
											aCell.ReleaseDispatch();
										}

										// 根据用户提示：FOP编号应该是行数-1，例如第二行（i=2）的FOP编号是1
									int fopNumber = i - 1;
									
									// 输出当前行使用的FOP编号到调试日志和界面
									CString debugMsg;
									debugMsg.Format(_T("[调试] 第 %d 行，使用FOP编号: %d\r\n"), i, fopNumber);
									OutputDebugString(debugMsg);
									m_reInfo.SetSel(-1, -1);
									m_reInfo.ReplaceSel(debugMsg);

										// 输出当前行读取的数据
										debugMsg.Format(_T("[调试] 第 %d 行数据 - Flash引脚: '%s', FOP编号: %d\r\n"), i, flashPin, fopNumber);
										OutputDebugString(debugMsg);

										// 存储Flash引脚与FOP的对应关系
										if (!flashPin.IsEmpty() && fopNumber > 0) {
											flashFopMap[flashPin] = fopNumber;
											// 输出成功存储的信息
											debugMsg.Format(_T("[调试] 成功存储Flash引脚与FOP对应关系: %s -> %d\r\n"), flashPin, fopNumber);
											OutputDebugString(debugMsg);
										} else {
											debugMsg.Format(_T("[调试] 跳过第 %d 行: Flash引脚为空或FOP编号无效\r\n"), i);
											OutputDebugString(debugMsg);
										}
									}

									// 输出最终的flashFopMap内容
									debugMsg.Format(_T("[调试] 读取完成，flashFopMap共存储了 %d 条记录\r\n"), flashFopMap.size());
									OutputDebugString(debugMsg);
									for (const auto& pair : flashFopMap) {
										debugMsg.Format(_T("[调试] flashFopMap最终内容: %s -> %d\r\n"), pair.first, pair.second);
										OutputDebugString(debugMsg);
									}

									usedRange.ReleaseDispatch();
								} else {
									debugMsg.Format(_T("[调试] 无法获取第7个Sheet页的UsedRange\r\n"));
									OutputDebugString(debugMsg);
								}
								sheet.ReleaseDispatch();
							} else {
								debugMsg.Format(_T("[调试] 无法获取第7个Sheet页的Dispatch\r\n"));
								OutputDebugString(debugMsg);
							}
						} else {
							debugMsg.Format(_T("[调试] Sheet数量不足7个，当前只有 %ld 个Sheet\r\n"), sheetCount);
							OutputDebugString(debugMsg);
						}

						sheets.ReleaseDispatch();
					}
					book.Close(vtMissing, vtMissing, vtMissing);
					book.ReleaseDispatch();
				}
				books.ReleaseDispatch();
			}
			app.Quit();
			app.ReleaseDispatch();
		}

		// 将IC Side数据与组件规格配对，并显示结果
		if (!icSideSpecMap.empty()) {
			int matchedCount = 0;

			// 遍历validPinData，查找匹配的icSideSpecMap的key
			for (size_t i = 0; i < validPinData.size(); i++) {
				CString pinValue = validPinData[i];
				// 检查icSideSpecMap中是否存在对应的key
				if (icSideSpecMap.find(pinValue) != icSideSpecMap.end()) {
					// 找到匹配，将ComponentSpec存入pinSpecMap
					pinSpecMap[pinValue] = icSideSpecMap[pinValue];
					matchedCount++;

					// 输出匹配信息到调试日志
					CString matchLog;
					ComponentSpec spec = icSideSpecMap[pinValue];
					matchLog.Format(_T("[调试] 找到匹配: Pin=%s, 电容=%s, 电压=%s\r\n"),
						pinValue, spec.capacitance, spec.voltage);
					OutputDebugString(matchLog);
				}
			}

			// 输出配对完成信息到调试日志
			CString finishLog;
			finishLog.Format(_T("[调试] 配对完成，成功匹配数量: %d, pinSpecMap大小: %zu\r\n"),
				matchedCount, pinSpecMap.size());
			OutputDebugString(finishLog);

			// 添加详细的配对信息
			for (auto& pair : pinSpecMap) {
				CString pinName = pair.first;
				ComponentSpec spec = pair.second;
				CString temp;
				temp.Format(_T("Pin: %s, 电容: %s, 电压: %s, 原始值: %s\r\n"),
					pinName, spec.capacitance, spec.voltage, spec.originalValue);
			}
		}
	}
}

bool CPinmapCheckToolDlg::PinMapICSideCheck(vector<int> validRowIndices, vector<CString> validPinData, vector<CString> validIcSideData, vector<double> validFop)
{
	// 校验筛选结果
	int validCount = validRowIndices.size();
	if (validCount == 0) {
		AfxMessageBox(_T("无符合m_fop数值范围的行数据！"));
		m_reInfo.SetWindowText(_T("筛选结果：无符合条件的行（m_fop数值范围：") + m_pinmapinfo_.icSideTopStr + _T("~") + m_pinmapinfo_.icSideBottomStr + _T("）"));
		return false;
	}

	// 6. 对比筛选后的数据（仅对比符合条件的行）
	int passCount = 0;
	passMsg = _T("=== PASS ===\r\n");
	failMsg = _T("=== RECHECK ===\r\n");
	for (int i = 0; i < validCount; i++) {
		int excelRow = validRowIndices[i]; // Excel实际行号
		CString pinVal = validPinData[i];
		CString icSideVal = validIcSideData[i];
		int fopNum = static_cast<int>(validFop[i]);
		CString temp;

		if (pinVal.Compare(icSideVal) == 0) {
			passCount++;
			temp.Format(_T("FOP NO.:%d：一致（值：%s）\r\n"), fopNum, pinVal);
			passMsg += temp;
		} else {
			temp.Format(_T("FOP NO.:%d：不一致（出PIN栏=%s，IC Side栏=%s）\r\n"),
				fopNum, pinVal, icSideVal);
			failMsg += temp;
		}
	}
	// 7. 汇总结果
	CString summary;
	summary.Format(_T("\r\n=== 出pin栏与ic side侧汇总 ===\r\n符合条件总行数：%d，一致行数：%d，不一致行数：%d\r\n"),
		validCount, passCount, validCount - passCount);
	passMsg += summary;
	return true;
}

void CPinmapCheckToolDlg::ResetButtonState()
{
	m_btnCheck.SetFaceColor(RGB(173, 216, 255));
	m_btnCheck.SetWindowText(_T("检查"));
	
	// 隐藏并重置进度条
	m_progressCheck.SetPos(0);
	m_progressCheck.ShowWindow(SW_HIDE);
	UpdateWindow();
}

bool CPinmapCheckToolDlg::ValidateInputParameters(int& sheetIndex, int& pinColIndex, int& icSideColIndex, int& fopColIndex)
{
	// 确认检查
	if (AfxMessageBox(_T("是否检查！"), MB_YESNO | MB_ICONQUESTION) == IDNO) {
		return FALSE;
	}

	// 校验并获取Sheet页编号
	// 不需要再次调用UpdateData
	m_pinmapinfo_.m_sheet_num.GetWindowText(sheetNumStr_);
	sheetNumStr_.Trim();
	if (sheetNumStr_.IsEmpty()) {
		AfxMessageBox(_T("请输入Sheet页编号！"));
		return FALSE;
	}

	sheetIndex = _ttoi(sheetNumStr_);  // Excel索引从0开始
	if (sheetIndex < 0) {
		AfxMessageBox(_T("Sheet页编号不能小于0！"));
		return FALSE;
	}

	// 校验并获取列名
	m_pinmapinfo_.m_pin.GetWindowText(pinColStr_);
	pinColStr_.Trim();
	if (pinColStr_.IsEmpty()) {
		AfxMessageBox(_T("请输入出PIN栏数据！"));
		return FALSE;
	}

	m_pinmapinfo_.m_ic_side.GetWindowText(icSideColStr_);
	icSideColStr_.Trim();
	if (icSideColStr_.IsEmpty()) {
		AfxMessageBox(_T("请输入Ic Side栏数据！"));
		return FALSE;
	}

	m_pinmapinfo_.m_fop.GetWindowText(fopStr_);
	fopStr_.Trim();
	if (fopStr_.IsEmpty()) {
		AfxMessageBox(_T("请输入Fop NO.数据！"));
		return FALSE;
	}

	// 将列名转换为索引
	pinColIndex = ColNameToIndex(pinColStr_);
	icSideColIndex = ColNameToIndex(icSideColStr_);
	fopColIndex = ColNameToIndex(fopStr_);

	if (pinColIndex <= 0 || icSideColIndex <= 0 || fopColIndex <= 0) {
		AfxMessageBox(_T("列名格式错误！"));
		return FALSE;
	}

	m_pinmapinfo_.m_ic_side_top.GetWindowText(m_pinmapinfo_.icSideTopStr);
	m_pinmapinfo_.m_ic_side_bottom.GetWindowText(m_pinmapinfo_.icSideBottomStr);
	m_pinmapinfo_.icSideTopStr.Trim();
	m_pinmapinfo_.icSideBottomStr.Trim();
	if (m_pinmapinfo_.icSideTopStr.IsEmpty() || m_pinmapinfo_.icSideBottomStr.IsEmpty()) {
		AfxMessageBox(_T("请输入IC Side Top/Bottom栏的数值上下限！"));
		return FALSE;
	}

	return TRUE;
}

// 初始化检查过程
void CPinmapCheckToolDlg::InitCheckProcess()
{
	// 初始化校验区
	m_reInfo.SetWindowText(_T(""));
	m_re_check.SetWindowText(_T(""));
	m_text_percent.SetWindowText(_T("0.0"));

	// 设置检查中按钮状态
	m_btnCheck.SetFaceColor(RGB(192, 192, 192));
	m_btnCheck.SetWindowText(_T("检查中..."));

	// 显示并初始化进度条
	m_progressCheck.SetRange(0, 100);
	m_progressCheck.SetPos(0);
	m_progressCheck.ShowWindow(SW_SHOW);
	UpdateWindow();
}

// 验证检查参数
bool CPinmapCheckToolDlg::ValidateCheckParameters(int& sheetIndex, int& pinColIndex, int& icSideColIndex, int& fopColIndex, double& icSideTop, double& icSideBottom)
{
	// 校验Pinmap文件已选择
	if (!m_btnSelPinmap_status || m_strPinmapPath.IsEmpty()) {
		AfxMessageBox(_T("请选择有效的Pinmap文件"));
		return false;
	}

	if (!ValidateInputParameters(sheetIndex, pinColIndex, icSideColIndex, fopColIndex)) {
		return false;
	}

	// 转换为数值（支持整数/小数）
	icSideTop = _tstof(m_pinmapinfo_.icSideTopStr);
	icSideBottom = _tstof(m_pinmapinfo_.icSideBottomStr);
	if (icSideTop > icSideBottom) {
		AfxMessageBox(_T("数值上下限错误（起始值不能大于结束值）！"));
		return false;
	}

	return true;
}

// 筛选有效数据
bool CPinmapCheckToolDlg::FilterValidData(int sheetIndex, int pinColIndex, int icSideColIndex, int fopColIndex, 
	double icSideTop, double icSideBottom,
	vector<int>& validRowIndices, vector<CString>& validPinData, vector<CString>& validIcSideData, vector<double>& validFop,
	vector<pair<int, CString>>& validPairList, vector<pair<int, CString>>& validIcSideList, vector<CString>& totalIcSideData)
{
	// 读取3列数据（pin列、icSide列、m_fop列）
	vector<CString> pinData, icSideData, fopDataStr;
	if (!ReadExcelColumn(m_strPinmapPath, sheetIndex, pinColIndex, pinData)) {
		m_reInfo.SetWindowText(_T("读取pin列失败！"));
		return false;
	}
	if (!ReadExcelColumn(m_strPinmapPath, sheetIndex, icSideColIndex, icSideData)) {
		m_reInfo.SetWindowText(_T("读取icSide列失败！"));
		return false;
	}
	if (!ReadExcelColumn(m_strPinmapPath, sheetIndex, fopColIndex, fopDataStr)) {
		m_reInfo.SetWindowText(_T("读取m_fop列失败！"));
		return false;
	}

	// 校验3列数据行数一致
	int totalDataRows = pinData.size();
	if (icSideData.size() != totalDataRows || fopDataStr.size() != totalDataRows) {
		AfxMessageBox(_T("pin、icSide、m_fop列的行数不一致！"));
		return false;
	}
	if (totalDataRows == 0) {
		AfxMessageBox(_T("所选列无有效数据！"));
		return false;
	}

	// 筛选：仅保留m_fop列数值在[icSideTop, icSideBottom]之间的行
	for (int i = 0; i < totalDataRows; i++) {
		// 转换m_fop列的字符串为数值
		CString fopStr = fopDataStr[i].Trim();
		if (fopStr.IsEmpty()) continue; // 跳过空值
		double fopVal = _tstof(fopStr);
		int fopInt = static_cast<int>(fopVal);
		validPairList.emplace_back(fopInt, pinData[i].Trim());
		validIcSideList.emplace_back(fopInt, icSideData[i].Trim());
		totalIcSideData.push_back(icSideData[i].Trim());
		// 检查是否在数值上下限内
		if (fopVal >= icSideTop && fopVal <= icSideBottom) {
			validFop.push_back(fopVal);
			validRowIndices.push_back(i + 1); // Excel行号从1开始
			validPinData.push_back(pinData[i].Trim());
			validIcSideData.push_back(icSideData[i].Trim());
		}
	}

	return true;
}

// 使用netlist比较组件规格 
void CPinmapCheckToolDlg::CompareComponentSpecs()
{
	if (!pinSpecMap.empty() && !m_netListPath.IsEmpty()) {
		// name ：refDesignators的映射
		map<CString, CString> pinToRefMap;
		
		// 读取netlist文件，查找包含"validPinData的值："格式的行
		CStdioFile netFile;
		if (netFile.Open(m_netListPath, CFile::modeRead | CFile::typeText)) {
			CString line;
			
			// 逐行读取netlist文件
			while (netFile.ReadString(line)) {
				line.Trim();
				
				// 检查行是否包含冒号（格式如"AVDD:C1(2:N2),U2(163:163AVDD)"）
				int colonPos = line.Find(_T(":"));
				if (colonPos > 0) {
					// 提取信号名（如"AVDD"）
					CString signalName = line.Left(colonPos).Trim();
					
					// 提取信号名后面的部分（如"C1(2:N2),U2(163:163AVDD)"）
					CString components = line.Mid(colonPos + 1).Trim();
					
					// 排除无效数据（如"AVDD:AVDD"）
					if (signalName.Compare(components) == 0) {
						continue;
					}
					
					// 提取所有的组件（如C1, U2等）
					CString refDesignators;
					int commaPos = 0;
					CString tempComponents = components;
					while (commaPos != -1) {
						commaPos = tempComponents.Find(_T(","));
						CString token;
						if (commaPos != -1) {
							token = tempComponents.Left(commaPos).Trim();
							tempComponents = tempComponents.Mid(commaPos + 1).Trim();
						} else {
							token = tempComponents.Trim();
						}
						
						if (!token.IsEmpty()) {
							// 提取组件名（如从"C1(2:N2)"中提取"C1"）
							int openParenPos = token.Find(_T("("));
							if (openParenPos > 0) {
								CString compName = token.Left(openParenPos).Trim();
								// 检查是否为电容（以C开头且后面紧跟数字）
								if (compName.GetAt(0) == _T('C')
									&& compName.GetLength() > 1
									&& compName.GetAt(1) >= _T('0') && compName.GetAt(1) <= _T('9')) {
									if (!refDesignators.IsEmpty()) {
										refDesignators += _T(",");
									}
									refDesignators += compName;
								}
							}
						}
					}
					
					// 如果找到了电容组件，存储到映射中
					if (!refDesignators.IsEmpty()) {
						pinToRefMap[signalName] = refDesignators;
						
						// 输出调试日志
						CString debugLog;
						debugLog.Format(_T("[调试] 从netlist中找到: 信号名: %s -> 电容组件: %s\r\n"), signalName, refDesignators);
						OutputDebugString(debugLog);
					}
				}
			}
			netFile.Close();
		}
		
		// 输出pinToRefMap大小到调试日志
		CString mapSizeLog;
		mapSizeLog.Format(_T("[调试] pinToRefMap大小: %zu\r\n"), pinToRefMap.size());
		OutputDebugString(mapSizeLog);
		
		// 输出pinToRefMap内容到调试日志
		for (auto& pair : pinToRefMap) {
			CString contentLog;
			contentLog.Format(_T("[调试] pinToRefMap - 键: %s, 值: %s\r\n"), pair.first, pair.second);
			OutputDebugString(contentLog);
		}
		
		// 9. 使用refDesignators在bomRefToSpecMap中查找组件规格并与pinSpecMap对比
		CString componentPassMsg = _T("\r\n=== 组件规格对比结果（PASS） ===\r\n");
		CString componentFailMsg = _T("\r\n=== 组件规格对比结果（不一致） ===\r\n");
		int componentPassCount = 0;
		int componentFailCount = 0;
		
		for (auto& pinPair : pinSpecMap) {
			CString pinValue = pinPair.first;
			ComponentSpec pinSpec = pinPair.second;
			
			// 查找对应的refDesignators
			if (pinToRefMap.find(pinValue) != pinToRefMap.end()) {
				CString refDesignators = pinToRefMap[pinValue];
				
				// 拆分refDesignators（可能包含逗号分隔的多个值）
				int commaPos = 0;
				CString tempRef = refDesignators;
				while (commaPos != -1) {
					commaPos = tempRef.Find(_T(","));
					CString token;
					if (commaPos != -1) {
						token = tempRef.Left(commaPos).Trim();
						tempRef = tempRef.Mid(commaPos + 1).Trim();
					} else {
						token = tempRef.Trim();
					}
					
					if (!token.IsEmpty()) {
						// 在bomRefToSpecMap中查找该refDesignator
						if (bomRefToSpecMap.find(token) != bomRefToSpecMap.end()) {
							ComponentSpec bomSpec = bomRefToSpecMap[token];
							
							// 比较电容和电压是否一致
							if (pinSpec.capacitance.Compare(bomSpec.capacitance) == 0 && 
								pinSpec.voltage.Compare(bomSpec.voltage) == 0) {
								// 一致，添加到passMsg
								CString temp;
								temp.Format(_T("Pin: %s, Ref: %s, 电容: %s, 电压: %s - 一致\r\n"), 
									pinValue, token, pinSpec.capacitance, pinSpec.voltage);
								componentPassMsg += temp;
								componentPassCount++;
							} else {
								// 不一致，添加到failMsg
								CString temp;
								temp.Format(_T("Pin: %s, Ref: %s - 不一致：pinSpec(电容: %s, 电压: %s) vs bomSpec(电容: %s, 电压: %s)\r\n"), 
									pinValue, token, pinSpec.capacitance, pinSpec.voltage, bomSpec.capacitance, bomSpec.voltage);
								componentFailMsg += temp;
								componentFailCount++;
							}
						} else {
							// 在bomRefToSpecMap中未找到该refDesignator
							CString temp;
							temp.Format(_T("Pin: %s, Ref: %s - 在BOM中未找到\r\n"), pinValue, token);
							componentFailMsg += temp;
							componentFailCount++;
						}
					}
				}
			} else {
				// 在netlist中未找到对应的refDesignators
				CString temp;
				temp.Format(_T("Pin: %s - 在netlist中未找到对应的refDesignators\r\n"), pinValue);
				componentFailMsg += temp;
				componentFailCount++;
			}
		}
		
		// 添加组件规格对比汇总信息
		CString componentSummary;
		componentSummary.Format(_T("\r\n=== 组件规格对比汇总 ===\r\n总对比数: %d, 一致数: %d, 不一致数: %d\r\n"), 
			componentPassCount + componentFailCount, componentPassCount, componentFailCount);
		componentPassMsg += componentSummary;
		
		// 将组件规格对比结果添加到主结果中
		passMsg += componentPassMsg;
		failMsg += componentFailMsg;
	}
}

void CPinmapCheckToolDlg::CompareFlashPins()
{
	// 检查Flash与FOP连接是否正确
	if (!flashPinMap.empty() && !flashFopMap.empty()) {
		CString flashPassMsg = _T("\r\n=== Flash与FOP连接检查结果（PASS） ===\r\n");
		CString flashFailMsg = _T("\r\n=== Flash与FOP连接检查结果（不一致） ===\r\n");
		
		int flashPassCount = 0;
		int flashTotalCount = flashFopMap.size();
		
		// 遍历Flash引脚与FOP的对应关系
		for (const auto& flashFopPair : flashFopMap) {
			CString flashPinName = flashFopPair.first;
			int fopNumber = flashFopPair.second;
			
			// 在Flash引脚映射中查找对应的引脚号
			int matchedPin = -1;
			for (const auto& pinPair : flashPinMap) {
				CString pinName = pinPair.second;
				// 检查引脚名称是否包含目标Flash引脚名称
				if (pinName.Find(flashPinName) >= 0) {
					matchedPin = pinPair.first;
					break;
				}
			}
			
			if (matchedPin != -1) {
				// 连接正确
				flashPassCount++;
				CString tmp;
				tmp.Format(_T("Flash引脚 %s (引脚号 %d) 与 FOP %d 连接正确\r\n"), 
					flashPinName, matchedPin, fopNumber);
				flashPassMsg += tmp;
				
				// 输出调试信息
				CString debugMsg;
				debugMsg.Format(_T("[调试] Flash连接检查通过: %s -> %d (引脚 %d)\r\n"), 
					flashPinName, fopNumber, matchedPin);
				OutputDebugString(debugMsg);
			} else {
				// 连接错误
				CString tmp;
				tmp.Format(_T("Flash引脚 %s 与 FOP %d 连接错误：未找到对应引脚\r\n"), 
					flashPinName, fopNumber);
				flashFailMsg += tmp;
				
				// 输出调试信息
				CString debugMsg;
				debugMsg.Format(_T("[调试] Flash连接检查失败: %s -> %d (未找到引脚)\r\n"), 
					flashPinName, fopNumber);
				OutputDebugString(debugMsg);
			}
		}
		
		// 添加总结信息
		if (flashPassCount == flashTotalCount) {
			flashPassMsg += _T("所有Flash与FOP连接均正确\r\n");
		} else {
			CString summary;
			summary.Format(_T("Flash与FOP连接检查：%d/%d 正确\r\n"), 
				flashPassCount, flashTotalCount);
			flashFailMsg += summary;
		}
		
		// 将结果添加到主结果中
		passMsg += flashPassMsg;
		failMsg += flashFailMsg;
	}
}

void CPinmapCheckToolDlg::CompareConnectorPins()
{
	CString connectorPassMsg = _T("\r\n=== 连接器引脚检查结果（PASS） ===\r\n");
	CString connectorFailMsg = _T("\r\n=== 连接器引脚检查结果（不一致） ===\r\n");
	for (const auto& connector : connectorMap) {
		auto it = connectorConfigMap.find(connector.first);
		if (it == connectorConfigMap.end()) {
			CString msg;
			msg.Format(_T("连接器引脚配置中未找到对应引脚：%d -> %s\r\n"), connector.first, connector.second);
			connectorFailMsg += msg;
		} else if (it->second != connector.second) {
			CString msg;
			msg.Format(_T("连接器引脚名称不匹配：%d -> 网表:%s 配置:%s\r\n"), connector.first, connector.second, it->second);
			connectorFailMsg += msg;
		} else {
			CString msg;
			msg.Format(_T("连接器引脚配置匹配：%d -> %s\r\n"), connector.first, connector.second);
			connectorPassMsg += msg;
		}
	}

	// 检查配置文件中是否有额外的引脚
	for (const auto& config : connectorConfigMap) {
		if (connectorMap.find(config.first) == connectorMap.end()) {
			CString msg;
			msg.Format(_T("网表中未找到配置文件中的引脚：%d -> %s\r\n"), config.first, config.second);
			connectorFailMsg += msg;
		}
	}
	passMsg += connectorPassMsg;
	failMsg += connectorFailMsg;	
}

// 显示检查结果
void CPinmapCheckToolDlg::DisplayCheckResults()
{
	// 显示结果到富文本控件
	m_reInfo.SetWindowText(passMsg);
	m_reInfo.SetSel(0, -1); // 全选（可选）

	m_btnCheck.SetWindowText(_T("Check"));
	m_btnCheck.SetFaceColor(RGB(173, 216, 255));

	// 显示不一致结果（红色）
	if (failMsg.GetLength() > 50) {
		m_re_check.SetWindowText(failMsg);
		CHARFORMAT cfRed;
		cfRed.cbSize = sizeof(CHARFORMAT);
		cfRed.dwMask = CFM_COLOR;
		cfRed.crTextColor = RGB(255, 0, 0);  // 红色
		m_re_check.SetSel(0, -1);
		m_re_check.SetSelectionCharFormat(cfRed);
	} else {
		m_re_check.SetWindowText(_T("所有符合条件的数据均一致，无不一致项！"));
		m_re_check.SetSel(0, -1);
	}
	
	SetMessage();

	m_progressCheck.SetPos(100);
	UpdateWindow();

	MessageBox(_T("Finished"));
}

void CPinmapCheckToolDlg::OnBnClickedCheck()
{
	InitCheckProcess();

	// pinmapExcel的参数验证
	// sheetIndex : Sheet栏
	// pinColIndex : 出pin栏
	// icSideColIndex : IC Side栏
	// fopColIndex : FOP NO.
	int sheetIndex, pinColIndex, icSideColIndex, fopColIndex;
	// IC Side Top和Bottom栏
	double icSideTop, icSideBottom;
	if (!ValidateCheckParameters(sheetIndex, pinColIndex, icSideColIndex, fopColIndex, icSideTop, icSideBottom)) {
		ResetButtonState();
		return;
	}
	// 更新进度条：参数验证完成
	m_progressCheck.SetPos(10);
	UpdateWindow();

	// 数据筛选
	vector<int> validRowIndices;
	vector<CString> validPinData, validIcSideData, totalIcSideData;
	vector<double> validFop;
	vector<pair<int, CString>> validPairList, validIcSideList;
	if (!FilterValidData(sheetIndex, pinColIndex, icSideColIndex, fopColIndex, 
		icSideTop, icSideBottom,
		validRowIndices, validPinData, validIcSideData, validFop,
		validPairList, validIcSideList, totalIcSideData)) {
		ResetButtonState();
		return;
	}
	// 更新进度条：数据筛选完成
	m_progressCheck.SetPos(30);
	UpdateWindow();

	// 执行检查和处理
	if (!PinMapICSideCheck(validRowIndices, validPinData, validIcSideData, validFop)) {
		ResetButtonState();
		return;
	}
	// 更新进度条：PinMap ICSide检查完成
	m_progressCheck.SetPos(40);
	UpdateWindow();
	
	ProcessNetlistFile(validPairList);
	// 更新进度条：Netlist文件处理完成
	m_progressCheck.SetPos(50);
	UpdateWindow();

	ProcessBOMFile();
	// 更新进度条：BOM文件处理完成
	m_progressCheck.SetPos(60);
	UpdateWindow();

	// mock connectorConfigMap数据
	for (int i = 1; i <= 40; i++) {
		CString strConnector;
		strConnector.Format(_T("J%d"), i);
		connectorConfigMap[i] = strConnector;
	}
	ProcessConfigFile(validPinData, validIcSideList, totalIcSideData);
	// 更新进度条：Config文件处理完成
	m_progressCheck.SetPos(70);
	UpdateWindow();

	// 比较组件规格
	CompareComponentSpecs();
	// 更新进度条：组件规格比较完成
	m_progressCheck.SetPos(75);
	UpdateWindow();

	//比较Flash
	CompareFlashPins();
	// 更新进度条：Flash比较完成
	m_progressCheck.SetPos(85);
	UpdateWindow();

	// 比较连接器引脚
	CompareConnectorPins();
	// 更新进度条：连接器引脚比较完成
	m_progressCheck.SetPos(90);
	UpdateWindow();

	// 显示结果
	DisplayCheckResults();
	// 更新进度条：检查完成
	
	// 重置按钮状态和隐藏进度条
	ResetButtonState();
}

void CPinmapCheckToolDlg::SetMessage()
{
	// 计算高度
	SCROLLINFO si;
	si.cbSize = sizeof(SCROLLINFO);
	si.fMask = SIF_RANGE | SIF_PAGE;
	m_re_check.GetScrollInfo(SB_VERT, &si);
	m_totalHeight = si.nMax;

	CRect clientRect;
	m_re_check.GetClientRect(clientRect);
	m_displayedHeight = clientRect.Height();
	// 计算当前显示区域与总内容的比例
	double percentage = 0.0;
	if (m_totalHeight > 1) {
		percentage = (double)m_displayedHeight / (m_totalHeight) * 100;
	}
	if (percentage > 100.0) {
		percentage = 100.0;
	}

	CString strPercent;
	strPercent.Format(_T("%.1f"), percentage);
	m_text_percent.SetWindowText(strPercent);
}

void CPinmapCheckToolDlg::OnBnClickedSaveLog()
{
	// 预处理字符串，移除可能导致截断的空字符
	CString strPassMsg = passMsg;
	strPassMsg.Replace(_T("\0"), _T("")); // 移除所有\0字符
	CString strFailMsg = failMsg;
	strFailMsg.Replace(_T("\0"), _T("")); // 移除所有\0字符

	// 定义文件过滤器
	CString filter = _T("文本文件 (*.txt)|*.txt|");

	// 创建保存文件对话框（第一个参数为FALSE表示保存模式）
	CFileDialog dlg(FALSE, _T("txt"), _T("新建文件.txt"),
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		filter, this);

	// 显示对话框
	if (dlg.DoModal() == IDOK) {
		CString filePath = dlg.GetPathName(); // 获取完整路径

		// 打开文件进行写入 - 使用CTextFile确保正确的文本编码处理
		CStdioFile file;
		if (file.Open(filePath, CFile::modeCreate | CFile::modeWrite | CFile::typeText)) {
			// 写入标题信息
			CTime currentTime = CTime::GetCurrentTime();
			CString timeStr = currentTime.Format(_T("%Y-%m-%d %H:%M:%S\r\n"));
			CString title = _T("PinMap Check Tool 日志文件\r\n");
			CString separator = _T("========================================\r\n");

			file.WriteString(title);
			file.WriteString(timeStr);
			file.WriteString(separator);
			
			// 写入PASS信息
			file.WriteString(_T("\r\nPASS 信息:\r\n"));
			file.WriteString(separator);
			
			if (strPassMsg.IsEmpty() || strPassMsg.GetLength() < 10) {
				file.WriteString(_T("暂无PASS信息，请先执行检查操作！\r\n"));
			} else {
				file.WriteString(strPassMsg);
			}
			
			// 写入ReCheck信息
			file.WriteString(_T("\r\nReCheck 信息:\r\n"));
			file.WriteString(separator);
			
			if (strFailMsg.IsEmpty() || strFailMsg.GetLength() < 10) {
				file.WriteString(_T("暂无ReCheck信息，请先执行检查操作！\r\n"));
			} else {
				file.WriteString(strFailMsg);
			}
			
			// 确保所有数据都写入磁盘
			file.Flush();
			file.Close(); // 只在成功打开时关闭文件
			
			AfxMessageBox(_T("日志保存成功！\n文件路径:\n") + filePath);
		} else {
			AfxMessageBox(_T("文件打开失败，无法保存日志！"));
		}
	}
}

void CPinmapCheckToolDlg::OnTcnSelchangeTab(NMHDR* pNMHDR, LRESULT* pResult)
{
	int nSel = m_tabCtrl.GetCurSel();
	*pResult = 0;

	m_pinmapinfo_.ShowWindow(SW_HIDE);
	m_general_.ShowWindow(SW_HIDE);
	m_component_.ShowWindow(SW_HIDE);
	m_hint_.ShowWindow(SW_HIDE);

	switch (nSel) {
	case 0:
		m_pinmapinfo_.ShowWindow(SW_SHOW);
		break;
	case 1:
		m_general_.ShowWindow(SW_SHOW);
		break;
	case 2:
		m_component_.ShowWindow(SW_SHOW);
		break;
	case 3:
		m_hint_.ShowWindow(SW_SHOW);
		break;
	default:
		break;
	}
}

HBRUSH CPinmapCheckToolDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// 设置特定控件的颜色
	if (pWnd == &m_edtSheetName) {
		// 设置Sheet Name编辑框的文本颜色为红色
		pDC->SetTextColor(RGB(255, 0, 0));
	}

	return hbr;
}

void CPinmapCheckToolDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX) {
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	} else {
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CPinmapCheckToolDlg::OnPaint()
{
	if (IsIconic()) {
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	} else {
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CPinmapCheckToolDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CPinmapCheckToolDlg::OnEnChangeEdtSheetName2()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnLvnItemchangedListPinmap(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: 在此添加控件通知处理程序代码
	*pResult = 0;
}

void CPinmapCheckToolDlg::OnEnChangeEdtSheetName()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnAcnStartAnimate2()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnEnChangeEdtMessage()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnStnClickedStcMessage()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnStnClickedTextPercent()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnStnClickedListPinmap()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnStnClickedPinmapName()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnEnChangeRichedit21()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}

void CPinmapCheckToolDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialogEx::OnSize(nType, cx, cy);
	if (cx == 0 || cy == 0 || !m_is_init) {
		return;
	}

	CRect recheckRect;
	recheckRect.left = m_re_check_rect.left;
	recheckRect.top = m_re_check_rect.top;
	recheckRect.right = (m_re_check_rect.right * cx) / m_orig_cx;
	recheckRect.bottom = (m_re_check_rect.bottom * cy) / m_orig_cy;
	m_re_check.MoveWindow(recheckRect);

	CRect informationRect;
	informationRect.left = m_information_rect.left;
	informationRect.top = m_information_rect.top;
	informationRect.right = (m_information_rect.right * cx) / m_orig_cx;
	informationRect.bottom = (m_information_rect.bottom * cy) / m_orig_cy;
	m_information.MoveWindow(informationRect);

	CRect passRect;
	passRect.left = m_pass_rect.left;
	passRect.top = (m_pass_rect.top * cy) / m_orig_cy;
	passRect.right = (m_pass_rect.right * cx) / m_orig_cx;
	passRect.bottom = (m_pass_rect.bottom * cy) / m_orig_cy;
	m_pass.MoveWindow(passRect);
	
	CRect reInfoRect;
	reInfoRect.left = m_reInfo_rect.left;
	reInfoRect.top = (m_reInfo_rect.top * cy) / m_orig_cy;
	reInfoRect.right = (m_reInfo_rect.right * cx) / m_orig_cx;
	reInfoRect.bottom = (m_reInfo_rect.bottom * cy) / m_orig_cy;
	m_reInfo.MoveWindow(reInfoRect);
	SetMessage();

	if (m_splitter.GetSafeHwnd()) {
		m_splitter.MoveWindow(0, 0, cx, cy);  // 根据窗口大小调整分割窗口的位置和大小
	}
}

void CPinmapCheckToolDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值

	CDialogEx::OnGetMinMaxInfo(lpMMI);
	lpMMI->ptMinTrackSize.y = m_nMinWindowHeight;
	lpMMI->ptMinTrackSize.x = m_nMinWindowWidth;
}
