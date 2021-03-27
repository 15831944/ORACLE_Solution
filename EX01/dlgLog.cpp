// dlgLog.cpp: 实现文件
//

#include "pch.h"
#include "EX01.h"
#include "dlgLog.h"
#include "afxdialogex.h"


// dlgLog 对话框

IMPLEMENT_DYNAMIC(dlgLog, CDialogEx)

dlgLog::dlgLog(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG1, pParent)
{

}

dlgLog::~dlgLog()
{
}

void dlgLog::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(dlgLog, CDialogEx)
	ON_BN_CLICKED(IDC_DlgLog_Cancel, &dlgLog::OnBnClickedDlglogCancel)
	ON_BN_CLICKED(IDC_DlgLog_Log, &dlgLog::OnBnClickedDlglogLog)
END_MESSAGE_MAP()


// dlgLog 消息处理程序


void dlgLog::OnBnClickedDlglogCancel()
{
	OnCancel();
	// TODO: 在此添加控件通知处理程序代码
}


void dlgLog::OnBnClickedDlglogLog()
{
	CString accContent = getInput(this, IDC_DlgLog_Account);
	CString pwdContent = getInput(this,IDC_DlgLog_Password);
	CString sSql;
	sSql.Format("select * ");
	OnOK();
}
