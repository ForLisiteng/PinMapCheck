
// PinmapCheckToolDlg.h: 头文件
//

#pragma once
#include "CDlgPinMapInfo.h"
#include "CDlgGeneral.h"
#include "CDlgComponent.h"
#include "CDlgHint.h"
#include <afxlayout.h> 
#include <map>
#include <vector>
#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CRange.h"
using namespace std;

// CPinmapCheckToolDlg 对话框
class CPinmapCheckToolDlg : public CDialogEx
{
// 构造
public:
	CPinmapCheckToolDlg(CWnd* pParent = nullptr);	// 标准构造函数

    void InitGeneralComboBox();
    CString passMsg, failMsg;
    bool PinMapICSideCheck(vector<int> validRowIndices, vector<CString> validPinData, vector<CString> validIcSideData, vector<double> validFop);
    void ResetButtonState();
    bool ValidateInputParameters(int& sheetIndex, int& pinColIndex, int& icSideColIndex, int& fopColIndex);
    void ProcessNetlistFile(vector<pair<int, CString>> validPairList);
    void ProcessBOMFile();
    void ProcessConfigFile(vector<CString> validPinData, vector<pair<int, CString>> validIcSideList, vector<CString> totalIcSideData);

    // 新增的封装函数
    void InitCheckProcess();
    bool ValidateCheckParameters(int& sheetIndex, int& pinColIndex, int& icSideColIndex, int& fopColIndex, double& icSideTop, double& icSideBottom);
    bool FilterValidData(int sheetIndex, int pinColIndex, int icSideColIndex, int fopColIndex, 
        double fopMin, double fopMax,
        vector<int>& validRowIndices, vector<CString>& validPinData, vector<CString>& validIcSideData, vector<double>& validFop,
        vector<pair<int, CString>>& validPairList, vector<pair<int, CString>>& validIcSideList, vector<CString>& totalIcSideData);
    void CompareComponentSpecs();
    void CompareConnectorPins();
    void DisplayCheckResults();
    void CompareFlashPins();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PINMAPCHECKTOOL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
    // 控件变量
    CMFCButton m_btnSelPinmap;
    bool m_btnSelPinmap_status = false;
    CMFCButton m_btnLoadNetlist;
    bool m_btnLoadNetlist_status = false;
    CMFCButton m_btnLoadBOM;
    bool m_btnLoadBOM_status = false;
    CMFCButton m_btnLoadCfg;
    bool m_btnLoadCfg_status = false;
    CMFCButton m_btnSaveCfg;
    CMFCButton m_btnCheck;
    CMFCButton m_btnSaveLog;

    CTabCtrl m_tabCtrl;

    CStatic m_stcSheetName;
    CEdit m_edtSheetName;
    CEdit m_edtSheetName2;
    CStatic m_stcMessage;
    CEdit m_edtMessage;
    CRichEditCtrl m_reInfo;
    CRichEditCtrl m_re_check;
    CRect m_re_check_rect;
    CProgressCtrl m_progressCheck;
    int m_orig_cx;               // 保存对话框初始宽度
    int m_orig_cy;               // 保存对话框初始高度
    

    // 消息处理函数
    afx_msg void OnBnClickedSelPinmap();
    afx_msg void OnBnClickedLoadNetlist();
    afx_msg void OnBnClickedLoadBom();
    afx_msg void OnBnClickedLoadCfg();
    afx_msg void OnBnClickedSaveCfg();
    afx_msg void OnBnClickedCheck();
    afx_msg void OnBnClickedSaveLog();
    afx_msg void OnTcnSelchangeTab(NMHDR* pNMHDR, LRESULT* pResult);
    afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
    afx_msg void OnEnChangeEdtSheetName2();
    afx_msg void OnLvnItemchangedListPinmap(NMHDR* pNMHDR, LRESULT* pResult);
    afx_msg void OnEnChangeEdtSheetName();
    afx_msg void OnAcnStartAnimate2();
    afx_msg void OnEnChangeEdtMessage();
    afx_msg void OnStnClickedStcMessage();
    BOOL ReadExcelColumn(const CString& strFilePath, int nSheetIndex, int nColIndex, vector<CString>& data);
    int ColNameToIndex(const CString& colName);
    void SaveGeneralSheet(CWorksheet& sheet, CWorksheets& worksheets, CWorkbook& workbook, CWorkbooks& workbooks, CApplication& excelApp);
    void SaveComponentSheet(CWorksheet& sheet, CWorksheets& worksheets, CWorkbook& workbook, CWorkbooks& workbooks, CApplication& excelApp);
    void SavePinMapSheet(CWorksheet& sheet, CWorksheets& worksheets, CWorkbook& workbook, CWorkbooks& workbooks, CApplication& excelApp);
    
private:
    void SetMessage();
    struct ControlLayout {
        CRect originalRect;
        float leftRatio, topRatio, rightRatio, bottomRatio;
    };

    CRect m_rect;
    map<int, vector<double>> tar;				//存储所有控件比例信息    key = 控件ID  value = 控件比例
    map<int, vector<double>>::iterator key;		//迭代器查找结果
    map<int, vector<double>>::iterator end;		//存储容器的最后一个元素
    int m_nMinWindowWidth;  // 左半部分固定宽度
    int m_nMinWindowHeight; // 窗口最小高度（左侧默认高度）

    CRect m_initialClientRect;
    float m_initialAspectRatio;
    std::map<UINT, ControlLayout> m_controlLayouts;
    bool m_maintainAspectRatio;

    CStatic m_text_percent;
    CDlgPinMapInfo m_pinmapinfo_;
    CString m_strPinmapPath;// 存储Pinmap文件的完整路径
    CDlgGeneral m_general_;
    CDlgComponent m_component_;
    CDlgHint m_hint_;

    // sheet页
    CString sheetNumStr_ = _T("0");
    // 出PIN栏
    CString pinColStr_ = _T("E");
    // ICSide栏
    CString icSideColStr_ = _T("O");
    // IC Side顶部行号
    CString topRowStr_ = _T("115");
    // IC Side底部行号
    CString bottomRowStr_ = _T("835");
    // FOP NO.
    CString fopStr_ = _T("C");
    // 到玻璃栏
    CString glassRailingStr_ = _T("E");
    // OLB栏
    CString olbStr_ = _T("AC");

    int m_totalHeight = 0;
    int m_displayedHeight = 0;
public:
    afx_msg void OnStnClickedTextPercent();
    CStatic m_picture;
    afx_msg void OnStnClickedListPinmap();
    CStatic m_staticPinmapName;
    afx_msg void OnStnClickedPinmapName();
    afx_msg void OnEnChangeRichedit21();
    CStatic m_netList;
    CString m_netListPath;// 存储netList文件的完整路径
    CStatic m_bom;
    CString m_bomPath;
    afx_msg void OnSize(UINT nType, int cx, int cy);
    afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
    CStatic m_setting;
    CStatic m_information;
    CRect m_information_rect;
    bool m_is_init;
    CStatic m_pass;
    CRect m_pass_rect;
    CRect m_reInfo_rect;
    CSplitterWnd m_splitter;
    CString m_loadCfgPath;
    struct ComponentSpec {
        CString capacitance; // 电容
        CString voltage;     // 电压(V)
        CString originalValue; // 原始值，用于显示
        vector<CString> refDesignators; // 参考标识符列表
    };
    
    // 用于组件规格到FOP NO.的映射
    struct SpecKey {
        CString capacitance;
        CString voltage;
        
        // 重载operator<以便在map中使用
        bool operator<(const SpecKey& other) const {
            int capCompare = capacitance.Compare(other.capacitance);
            if (capCompare != 0) return capCompare < 0;
            return voltage.Compare(other.voltage) < 0;
        }
    };
    
    // 类型定义以简化代码
    typedef map<CString, ComponentSpec> ComponentSpecMap;
    typedef vector<pair<int, CString>> FopDataList;
    
    map<CString, int> flashFopMap; // Config中Flash引脚与FOP的对应关系
    
    map<int, CString> flashPinMap; // netList的Flash引脚映射
    map<int, CString> connectorMap; // netList的连接器引脚映射
    map<int, CString> connectorConfigMap; // Config的连接器引脚映射
    
    // 成员变量
    FopDataList uniqueNetPinList; // 存储网表中DDIC ID的引脚信息：<FOP NO., Name>
    ComponentSpecMap bomRefToSpecMap; // <Ref Designator, 组件规格> bom文件对应的关系 (C30:2.2uF/6.3V)
    ComponentSpecMap icSideSpecMap;  // <IC Side值, 组件规格> 全部config中的name与spec的组件规格映射 (VDDB_VCI:2.2uF/6.3V)
    ComponentSpecMap pinSpecMap;     // <validPinData的值, 组件规格> 从config中获取的在icSide上下界内的 name与spec的组件规格映射 (VDDB_VCI:2.2uF/6.3V)
};
