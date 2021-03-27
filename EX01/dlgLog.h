#pragma once

#include"DB/Oracle.h"
#include"tool.h"
// dlgLog 对话框

class dlgLog : public CDialogEx
{
	DECLARE_DYNAMIC(dlgLog)

public:
	dlgLog(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~dlgLog();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	//afx_msg void OnEnChangeEdit1();
	afx_msg void OnBnClickedDlglogCancel();
	afx_msg void OnBnClickedDlglogLog();
};
