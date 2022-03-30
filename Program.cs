using System.Text;
using System.Diagnostics;
using Win32;
using mshtml;

IntPtr MainHwnd;

static bool EnumChildWindowsProc(IntPtr hWnd, IntPtr lParam)
{
  var count = WinNative.GetWindowTextLength(hWnd);
  var sb = new StringBuilder(500);
  var ret = WinNative.GetWindowText(hWnd, sb, 500);

  StringBuilder className = new(500);
  IntPtr ChildHwnd = hWnd;
  int ChildRet = WinNative.GetClassName(ChildHwnd, className, 500);

  StringBuilder TitleName = new(500);

  int ChildTitle = WinNative.GetWindowText(ChildHwnd, TitleName, 500);

  Console.WriteLine("{0} 0x{1:X8} - {2} {3} {4}", new string('+', ((int)lParam)), hWnd, sb.ToString(), className, TitleName);

  // さらに自身の子ウインドウを列挙
  WinNative.EnumChildWindows(hWnd, EnumChildWindowsProc, lParam + 1);

  return true;
}

bool EnumChildProcIES(IntPtr hWnd, IntPtr lParam)
{
  var count = WinNative.GetWindowTextLength(hWnd);
  var sb = new StringBuilder(500);
  var ret = WinNative.GetWindowText(hWnd, sb, 500);
  IntPtr IESHandle = new IntPtr(0);

  StringBuilder className = new(500);
  IntPtr ChildHwnd = hWnd;
  int ChildRet = WinNative.GetClassName(ChildHwnd, className, 500);

  StringBuilder TitleName = new(500);

  int ChildTitle = WinNative.GetWindowText(ChildHwnd, TitleName, 500);

  //Console.WriteLine("{0} 0x{1:X8} - {2} {3} {4}", new string('+', ((int)lParam)), hWnd, sb.ToString(), className, TitleName);
  if(className.ToString() == "Internet Explorer_Server"){
    MainHwnd = hWnd;
  }
  object obj = WinNative.GetHTMLDocumentFromIES(MainHwnd);
  if(obj is null){
  }else{
    IHTMLDocument3 ReHtml = (IHTMLDocument3)obj;
    Console.WriteLine(ReHtml.getElementsByName("q").length);
  }


  // さらに自身の子ウインドウを列挙
  WinNative.EnumChildWindows(hWnd, EnumChildProcIES, lParam + 1);

  return true;

}

//全てのプロセスを列挙する
foreach (Process p in
    Process.GetProcesses())
{
  //メインウィンドウのタイトルがある時だけ列挙する
  if (p.MainWindowTitle.Length != 0)
  {
    if (p.ProcessName == "LegacyBrowser")
    {
      Console.WriteLine("プロセス名:" + p.ProcessName);
      Console.WriteLine("タイトル名:" + p.MainWindowTitle);
      MainHwnd = p.MainWindowHandle;

      int length = WinNative.GetWindowTextLength(MainHwnd);
      WinNative.EnumChildWindows(MainHwnd, EnumChildProcIES, new IntPtr(1));

      StringBuilder className = new(500);
      int ret = WinNative.GetClassName(MainHwnd, className, 500);
    }
  }
}