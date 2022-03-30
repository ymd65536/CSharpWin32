using System.Runtime.InteropServices;
using System.Text;
// using System.Windows.Forms;

namespace Win32
{
  class WinNative
  {
     public delegate bool EnumWindowsProc(
       IntPtr hWnd
       , IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumWindows(
      EnumWindowsProc lpEnumFunc
      , IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumChildWindows(
        IntPtr hWndParent,
        EnumWindowsProc lpEnumFunc,
        IntPtr lParam
    );

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowTextLength(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern Int32 GetWindowText(
        IntPtr hWnd
        ,StringBuilder lpString
        ,Int32 nMaxCount
    );

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(
      IntPtr hWnd
    , StringBuilder lpClassName
    , int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern uint RegisterWindowMessage(string lpString);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(
      IntPtr hwnd
      ,uint msg 
      , IntPtr wParam
      , IntPtr lParam
      , SendMessageTimeoutFlags flags
      , uint uTimeout
      , out IntPtr lpdwResult);

    [Flags]
    public enum SendMessageTimeoutFlags : uint
    {
        SMTO_NORMAL = 0x0
        ,SMTO_BLOCK = 0x1
        ,SMTO_ABORTIFHUNG = 0x2
        ,SMTO_NOTIMEOUTIFNOTHUNG = 0x8
    }

    // [DllImport("ole32.dll", CharSet = CharSet.Auto,PreserveSig=false)]
    // public static extern Guid IIDFromString(string lpsz);

    [DllImport("ole32.dll", CharSet = CharSet.Auto)]
    public static extern int IIDFromString(string lpsz, out Guid lpiid);


/*
    [DllImport("oleacc.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr ObjectFromLresult(
        IntPtr lResult
        , Guid riid 
        ,IntPtr wParam
        , object ppvObject);
*/
    [DllImport("oleacc.dll", PreserveSig=false)]
    [return: MarshalAs(UnmanagedType.Interface)]
    public static extern object ObjectFromLresult(
      IntPtr lResult
      , [MarshalAs(UnmanagedType.LPStruct)] Guid refiid
      , IntPtr wParam);

    public static object GetHTMLDocumentFromIES(IntPtr hwnd)
    {
      uint msg;
      IntPtr res;
      //Guid[] iid = new Guid[3];

      msg = WinNative.RegisterWindowMessage("WM_HTML_GETOBJECT");
      res = new IntPtr(0);
      WinNative.SendMessageTimeout(
        hwnd
        ,msg
        ,new IntPtr(0)
        ,new IntPtr(0)
        ,SendMessageTimeoutFlags.SMTO_ABORTIFHUNG
        ,1000
        ,out res);

      object obj = null;

      if(res.ToInt64() > 0)
      {
        string IID_IHTMLDocument3 = "{3050f485-98b5-11cf-bb82-00aa00bdce0b}";
        Guid IID_IHTMLDocument = new Guid(IID_IHTMLDocument3);
        if(res != IntPtr.Zero)
        {
          obj = WinNative.ObjectFromLresult(res,IID_IHTMLDocument,new IntPtr(0));
          if(obj is null){

          }else{
            if(obj == DBNull.Value){
              
            }else{
              return obj;
            }
          }
        }
      }
      return obj;
    }
  }
}