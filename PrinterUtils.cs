namespace PrinterUtils
{

    using System.ComponentModel;
    using System.Collections;
    using System.Runtime.InteropServices;
    using System;


    public class PrinterSettings
    {
        #region "Private Variables"
        private IntPtr hPrinter = new System.IntPtr();
        private PRINTER_DEFAULTS PrinterValues = new PRINTER_DEFAULTS();
        private PRINTER_INFO_2 pinfo = new PRINTER_INFO_2();
        private DEVMODE dm;
        private IntPtr ptrDM;
        private IntPtr ptrPrinterInfo;
        private int sizeOfDevMode = 0;
        private int lastError;
        private int nBytesNeeded;
        private long nRet;
        private int intError;
        private System.Int32 nJunk;
        private IntPtr yDevModeData;

        #endregion
        #region "Win API Def"
        [DllImport("kernel32.dll", EntryPoint = "GetLastError", SetLastError = false,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern Int32 GetLastError();
        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool ClosePrinter(IntPtr hPrinter);
        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesA", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter,
        [MarshalAs(UnmanagedType.LPStr)] string pDeviceNameg,
        IntPtr pDevModeOutput, ref IntPtr pDevModeInput, int fMode);
        [DllImport("winspool.Drv", EntryPoint = "GetPrinterA", SetLastError = true,
            CharSet = CharSet.Ansi, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        private static extern bool GetPrinter(IntPtr hPrinter, Int32 dwLevel,
        IntPtr pPrinter, Int32 dwBuf, out Int32 dwNeeded);
        /*[DllImport("winspool.Drv", EntryPoint="OpenPrinterA", 
            SetLastError=true, CharSet=CharSet.Ansi, 
            ExactSpelling=true, CallingConvention=CallingConvention.StdCall)]
        static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, 
        out IntPtr hPrinter, ref PRINTER_DEFAULTS pd)

        [ DllImport( "winspool.drv",CharSet=CharSet.Unicode,ExactSpelling=false,
        CallingConvention=CallingConvention.StdCall )]
        public static extern long OpenPrinter(string pPrinterName, 
                 ref IntPtr phPrinter, int pDefault);*/

        /*[DllImport("winspool.Drv", EntryPoint="OpenPrinterA", 
            SetLastError=true, CharSet=CharSet.Ansi, 
            ExactSpelling=true, CallingConvention=CallingConvention.StdCall)]
        static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, 
        out IntPtr hPrinter, ref PRINTER_DEFAULTS pd);
        */
        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA",
            SetLastError = true, CharSet = CharSet.Ansi,
            ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool
            OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter,
            out IntPtr hPrinter,
            //ref PRINTER_DEFAULTS pd);
            IntPtr ptr);
        [DllImport("winspool.drv", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern bool SetPrinter(IntPtr hPrinter, int Level, IntPtr
        pPrinter, int Command);

        /*[DllImport("winspool.drv", CharSet=CharSet.Ansi, SetLastError=true)]
        private static extern bool SetPrinter(IntPtr hPrinter, int Level, IntPtr 
        pPrinter, int Command);*/

        // Wrapper for Win32 message formatter.
        /*[DllImport("kernel32.dll", CharSet=System.Runtime.InteropServices.CharSet.Auto)]
        private unsafe static extern int FormatMessage( int dwFlags,
        ref IntPtr pMessageSource,
        int dwMessageID,
        int dwLanguageID,
        ref string lpBuffer,
        int nSize,
        IntPtr* pArguments);*/
        #endregion
        #region "Data structure"
        [StructLayout(LayoutKind.Sequential)]
        public struct PRINTER_DEFAULTS
        {
            public IntPtr pDatatype;
            public IntPtr pDevMode;
            public int DesiredAccess;
        }
        /*[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
        public struct PRINTER_DEFAULTS
        {
        public int pDataType;
        public IntPtr pDevMode;
        public ACCESS_MASK DesiredAccess;
        }*/

        [StructLayout(LayoutKind.Sequential)]
        private struct PRINTER_INFO_2
        {
            [MarshalAs(UnmanagedType.LPStr)] public string pServerName;
            [MarshalAs(UnmanagedType.LPStr)] public string pPrinterName;
            [MarshalAs(UnmanagedType.LPStr)] public string pShareName;
            [MarshalAs(UnmanagedType.LPStr)] public string pPortName;
            [MarshalAs(UnmanagedType.LPStr)] public string pDriverName;
            [MarshalAs(UnmanagedType.LPStr)] public string pComment;
            [MarshalAs(UnmanagedType.LPStr)] public string pLocation;
            public IntPtr pDevMode;
            [MarshalAs(UnmanagedType.LPStr)] public string pSepFile;
            [MarshalAs(UnmanagedType.LPStr)] public string pPrintProcessor;
            [MarshalAs(UnmanagedType.LPStr)] public string pDatatype;
            [MarshalAs(UnmanagedType.LPStr)] public string pParameters;
            public IntPtr pSecurityDescriptor;
            public Int32 Attributes;
            public Int32 Priority;
            public Int32 DefaultPriority;
            public Int32 StartTime;
            public Int32 UntilTime;
            public Int32 Status;
            public Int32 cJobs;
            public Int32 AveragePPM;
        }
        private const short CCDEVICENAME = 32;
        private const short CCFORMNAME = 32;
        [StructLayout(LayoutKind.Sequential)]
        public struct DEVMODE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCDEVICENAME)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public short dmOrientation;
            public short dmPaperSize;
            public short dmPaperLength;
            public short dmPaperWidth;
            public short dmScale;
            public short dmCopies;
            public short dmDefaultSource;
            public short dmPrintQuality;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCFORMNAME)]
            public string dmFormName;
            public short dmUnusedPadding;
            public short dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
        }
        #endregion
        #region "Constants"
        private const int DM_DUPLEX = 0x1000;
        private const int DM_IN_BUFFER = 8;
        private const int DM_OUT_BUFFER = 2;
        private const int PRINTER_ACCESS_ADMINISTER = 0x4;
        private const int PRINTER_ACCESS_USE = 0x8;
        private const int STANDARD_RIGHTS_REQUIRED = 0xF0000;
        private const int PRINTER_ALL_ACCESS =
            (STANDARD_RIGHTS_REQUIRED | PRINTER_ACCESS_ADMINISTER
            | PRINTER_ACCESS_USE);
        #endregion

        #region "OpenPrinter Modifier"
        public static bool OpenPrinter(string szPrinter, out IntPtr hPrinter, ref PRINTER_DEFAULTS pd)
        {
            GCHandle ptr = GCHandle.Alloc(pd, GCHandleType.Pinned);

            bool bRet = PrinterSettings.OpenPrinter(szPrinter, out hPrinter, ptr.AddrOfPinnedObject());

            ptr.Free();

            return bRet;
        }
        #endregion


        #region "Function to change printer settings" 
        public bool ChangePrinterSetting(string PrinterName, PrinterData PS)

        {

            if (((int)PS.Duplex < 1) || ((int)PS.Duplex > 3))
            {
                throw new ArgumentOutOfRangeException("nDuplexSetting",
                                         "nDuplexSetting is incorrect.");
            }
            else
            {
                dm = this.GetPrinterSettings(PrinterName);
                dm.dmDefaultSource = (short)PS.source;
                dm.dmOrientation = (short)PS.Orientation;
                dm.dmPaperSize = (short)PS.Size;
                dm.dmDuplex = (short)PS.Duplex;
                Marshal.StructureToPtr(dm, yDevModeData, true);
                pinfo.pDevMode = yDevModeData;
                pinfo.pSecurityDescriptor = IntPtr.Zero;
                /*update driver dependent part of the DEVMODE
                1 = DocumentProperties(IntPtr.Zero, hPrinter, sPrinterName, yDevModeData
                , ref pinfo.pDevMode, (DM_IN_BUFFER | DM_OUT_BUFFER));*/
                Marshal.StructureToPtr(pinfo, ptrPrinterInfo, false);
                lastError = Marshal.GetLastWin32Error();
                nRet = Convert.ToInt16(SetPrinter(hPrinter, 2, ptrPrinterInfo, 0));
                if (nRet == 0)
                {
                    //Unable to set shared printer settings.
                    lastError = Marshal.GetLastWin32Error();
                    //string myErrMsg = GetErrorMessage(lastError);
                    throw new Win32Exception(Marshal.GetLastWin32Error());

                }
                if (hPrinter != IntPtr.Zero)
                    ClosePrinter(hPrinter);
                return Convert.ToBoolean(nRet);
            }
        }
        private DEVMODE GetPrinterSettings(string PrinterName)
        {
            PrinterData PData = new PrinterData();
            DEVMODE dm;
            const int PRINTER_ACCESS_ADMINISTER = 0x4;
            const int PRINTER_ACCESS_USE = 0x8;
            const int PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED |
                       PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE);


            PrinterValues.pDatatype = IntPtr.Zero;
            PrinterValues.pDevMode = IntPtr.Zero;
            PrinterValues.DesiredAccess = PRINTER_ALL_ACCESS;
            nRet = Convert.ToInt64(OpenPrinter(PrinterName,
                           out hPrinter, ref PrinterValues));
            if (nRet == 0)
            {
                lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out nBytesNeeded);
            if (nBytesNeeded <= 0)
            {
                throw new System.Exception("Unable to allocate memory");
            }
            else
            {
                // Allocate enough space for PRINTER_INFO_2... 
                { //ptrPrinterIn fo =
                    Marshal.AllocCoTaskMem(nBytesNeeded);
                };
                ptrPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded);
                // The second GetPrinter fills in all the current settings, so all you 
                // need to do is modify what you're interested in...
                nRet = Convert.ToInt32(GetPrinter(hPrinter, 2,
                    ptrPrinterInfo, nBytesNeeded, out nJunk));
                if (nRet == 0)
                {
                    lastError = Marshal.GetLastWin32Error();
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                pinfo = (PRINTER_INFO_2)Marshal.PtrToStructure(ptrPrinterInfo,
                                                      typeof(PRINTER_INFO_2));
                IntPtr Temp = new IntPtr();
                if (pinfo.pDevMode == IntPtr.Zero)
                {
                    // If GetPrinter didn't fill in the DEVMODE, try to get it by calling
                    // DocumentProperties...
                    IntPtr ptrZero = IntPtr.Zero;
                    //get the size of the devmode structure
                    sizeOfDevMode = DocumentProperties(IntPtr.Zero, hPrinter,
                                       PrinterName, ptrZero, ref ptrZero, 0);

                    ptrDM = Marshal.AllocCoTaskMem(sizeOfDevMode);
                    int i;
                    i = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, ptrDM,
                    ref ptrZero, DM_OUT_BUFFER);
                    if ((i < 0) || (ptrDM == IntPtr.Zero))
                    {
                        //Cannot get the DEVMODE structure.
                        throw new System.Exception("Cannot get DEVMODE data");
                    }
                    pinfo.pDevMode = ptrDM;
                }
                intError = DocumentProperties(IntPtr.Zero, hPrinter,
                          PrinterName, IntPtr.Zero, ref Temp, 0);
                //IntPtr yDevModeData = Marshal.AllocCoTaskMem(i1);
                yDevModeData = Marshal.AllocHGlobal(intError);
                intError = DocumentProperties(IntPtr.Zero, hPrinter,
                         PrinterName, yDevModeData, ref Temp, 2);
                dm = (DEVMODE)Marshal.PtrToStructure(yDevModeData, typeof(DEVMODE));
                //nRet = DocumentProperties(IntPtr.Zero, hPrinter, sPrinterName, yDevModeData
                // , ref yDevModeData, (DM_IN_BUFFER | DM_OUT_BUFFER));
                if ((nRet == 0) || (hPrinter == IntPtr.Zero))
                {
                    lastError = Marshal.GetLastWin32Error();
                    //string myErrMsg = GetErrorMessage(lastError);
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                return dm;
            }
            #endregion
        }

        public PrinterData GetPrinterSettingsInPrinterData(string PrinterName)
        {
            DEVMODE dm = GetPrinterSettings(PrinterName);
            PrinterData printerData = new PrinterData();
            printerData.Duplex = dm.dmDuplex;
            printerData.source = dm.dmDefaultSource;
            printerData.Orientation = dm.dmOrientation;
            printerData.Size = dm.dmPaperSize;

            return printerData;
        }
    }

    public class PrinterData
    {
        public PrinterData()
        {
            Duplex = 1;
            source = 1;
            Orientation = 1;
            Size = 1;
        }

        public int Duplex { get; set; }
        public int source { get; set; }
        public int Orientation { get; set; }
        public int Size { get; set; }
    }
}
