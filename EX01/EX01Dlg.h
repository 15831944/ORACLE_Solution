
// EX01Dlg.h: 头文件
//

#pragma once


// CEX01Dlg 对话框
class CEX01Dlg : public CDialogEx
{
	// 构造
public:
	CEX01Dlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EX01_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
private:
	CString  getInput(UINT nid);
public:
	afx_msg void OnBnClickedButtonBooksinInsert();
	afx_msg void OnBnClickedBooksoutInsert();
};
