using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace itConference
{
    class ClassSenseLock
    {
        //elitee library interface declaration, 
        //return value and some constant definition.

        // 设备打开模式
        static public uint ELE_EXCLUSIVE_MODE = 0x00000000;// 独占方式打开
        static public uint ELE_SHARE_MODE = 0x00000001;// 共享方式打开

        // 设备通讯模式
        static public uint ELE_COMM_USB_MODE = 0x00000000;// USB模式
        static public uint ELE_COMM_HID_MODE = 0x000000AA;// HID模式

        // 设备版本信息
        static public uint ELE_V0201 = 0x00000201;// 精锐E设备2.01版



        // 设备类型
        static public uint ELE_LOCAL_DEVICE = 0x00000000;// 单机设备
        static public uint ELE_NET_DEVICE = 0x00000001;// 网络设备
        static public uint ELE_NORMAL_DEVICE = 0x00000000;// 无时钟设备

        static public uint ELE_RTL_DEVICE = 0x00000002;// 实时时钟设备
        static public uint ELE_USER_DEVICE = 0x00000000;// 用户设备
        static public uint ELE_MASTER_DEVICE = 0x00000080;// 升级控制设备

        // 设备的信息

        static public uint ELE_GET_DEVICE_SERIAL = 0x00000001;// 设备序列号，8字节
        static public uint ELE_GET_VENDOR_DESC = 0x00000002;// 厂商描述，8字节
        static public uint ELE_GET_CURRENT_TIME = 0x00000004;// 设备当前的时间，仅对时钟锁有效，返回的值与crt中的_time函数返回的值意义一致

        static public uint ELE_GET_DEVICE_VERSION = 0x00000007;// 设备的版本号
        static public uint ELE_GET_DEVICE_TYPE = 0x00000008;// 设备的类型

        static public uint ELE_GET_MANUFACTURE_TIME = 0x00000009;// 设备的生产日期，返回的值与crt中的_time函数返回值意义一致

        static public uint ELE_GET_MODIFY_TIME = 0x0000000A;// 设备的修改日期，返回的值与crt中的_time函数返回值意义一致

        static public uint ELE_GET_COMM_MODE = 0x0000000B;// 设备的通讯模式，USB的，或者是HID的

        static public uint ELE_GET_DEVELOPER_NUMBER = 0x00000012;// 开发商编号
        static public uint ELE_GET_MODULE_COUNT = 0x00000013;// 模块总数
        static public uint ELE_GET_MODULE_SIZE = 0x00000014;// 模块大小


        // 设置相关命令
        static public uint ELE_RESET_DEVICE = 0x00000000;// 重置设备，PIN码验证后的安全状态丢失

        static public uint ELE_SET_LCD_UP = 0x00000001;// 设置灯亮的状态

        static public uint ELE_SET_LCD_DOWN = 0x00000002;// 设置灯灭的状态

        static public uint ELE_SET_LCD_FLASH = 0x00000003;// 设置灯闪烁的状态    
        static public uint ELE_SET_COMM_MODE = 0x00000005;// 设置通讯模式
        static public uint ELE_SET_VENDOR_DESC = 0x00000006;// 设置厂商描述



        // 错误返回值

        static public uint ELE_SUCCESS = 0x00000000;// 成功
        static public uint ELE_INVALID_PARAMETER = 0x00000001;// 参数错误
        static public uint ELE_INSUFFICIENT_BUFFER = 0x00000002;// 缓冲区不够

        static public uint ELE_NOT_ENOUGH_MEMORY = 0x00000003;// 内存不够
        static public uint ELE_INVALID_DEVICE_HANDLE = 0x00000004;// 设备句柄无效
        static public uint ELE_COMM_ERROR = 0x00000005;// 通讯错误

        static public uint ELE_INVALID_SHARE_MODE = 0x00000006;// 打开设备时，使用了无效的共享方式
        static public uint ELE_UNSUPPORTED_OS = 0x00000007;// 精锐E不支持该操作系统
        static public uint ELE_ENUMERATE_DEVICE_FAILED = 0x00000008;// 列举设备失败
        static public uint ELE_NO_MORE_DEVICE = 0x00000009;// 没有符合条件的设备



        static public uint ELE_NOT_RTL_DEVICE = 0x00000024;// 不是时钟设备
        static public uint ELE_NOT_MASTER_DEVICE = 0x00000025;// 不是升级控制锁

        static public uint ELE_MODULE_NOT_FOUND = 0x00000027;// 指定的模块不存在
        static public uint ELE_MODULE_SIZE_BEYOND = 0x00000028;// 执行操作时超越模块区域空间

        static public uint ELE_INVALID_PACKAGE = 0x0000002C;// 无效的升级包，可能网络传输过程造成破坏（Hash验证错误）

        static public uint ELE_INVALID_STRUCTURE_SIZE = 0x00000066;// 结构体的Size值设置不正确
        static public uint ELE_CLOSE_DEVICE_FAILED = 0x00000067;// 关闭设备失败
        static public uint ELE_VERIFY_SIGNATURE_ERROR = 0x00000068;// 验证签名失败 


        //Hardware's error
        static public uint ELE_WRITE_FLASH_ERROR = 0x00008001;// 写精锐E硬件产生错误
        static public uint ELE_UNSUPPORTED_FUNCTION = 0x00008002;// 设备功能不支持

        static public uint ELE_RTC_ERROR = 0x00008003;// 时钟锁错误

        static public uint ELE_RTC_POWER_OFF = 0x00008004;// 时钟锁电源已断电 
        static public uint ELE_WRONG_COMMAND_INS = 0x00008101;// 传输命令参数中INS错误
        static public uint ELE_WRONG_COMMAND_CLA = 0x00008102;// 传输命令参数中CLA错误
        static public uint ELE_WRONG_COMMAND_LC = 0x00008301;// 传输命令参数中LC错误
        static public uint ELE_WRONG_COMMAND_DATA = 0x00008302;// 传输命令参数中DATA错误
        static public uint ELE_WRONG_COMMAND_LE = 0x00008303;// 传输命令参数中LE错误
        static public uint ELE_WRONG_COMMAND_PP = 0x00008306;// 传输命令参数中PP错误
        static public uint ELE_INVALID_OFFSET = 0x00008307;// 无效的偏移量
        static public uint ELE_INVALID_MODULE = 0x00008501;// 无效的模块

        static public uint ELE_MACHINE_CODE_MISMATCH = 0x00008502;// 机器码不匹配
        static public uint ELE_INVALID_HEAD_SIZE = 0x00008503;// 可选头的大小无效

        static public uint ELE_INVALID_SECTION_NUMBER = 0x00008504;// 段的编号无效
        static public uint ELE_SECTION_NOT_FOUND = 0x00008505;// 段不存在
        static public uint ELE_INVALID_SECTION_OFFSET = 0x00008506;// 段的偏移量无效

        static public uint ELE_INVALID_SECTION_SIZE = 0x00008507;// 段的大小无效
        static public uint ELE_SECTION_TYPE_MISMATCH = 0x00008508;// 段的类型不匹配

        static public uint ELE_ANONYMOUS_USER = 0x00008800;// 当前用户是匿名的用户
        static public uint ELE_DEVICE_STATE_MISMATCH = 0x00008801;// 设备状态不匹配
        static public uint ELE_SECURITY_STATE_MISMATCH = 0x00008802;// 安全状态不匹配
        static public uint ELE_DEVICE_PIN_BLOCK = 0x00008803;// 设备PIN码已锁死
        static public uint ELE_SECURITY_MESSAGE_ERROR = 0x00008808;// 升级包中安全信息不正确

        static public uint ELE_DEVICE_PIN_ERROR = 0x000088c0;// 88cx设备PIN码错误，还有x次尝试机会

        static public uint ELE_OBJECT_NOT_FOUND = 0x00008901;// 对象不存在

        static public uint ELE_NO_CURRENT_MODULE = 0x00008902;// 指定的模块不存在
        static public uint ELE_OBJECT_NAME_ALREADY_EXIST = 0x00008903;// 对象名已经存在

        static public uint ELE_BEYOND_OBJECT_SIZE = 0x00008905;// 超过对象的大小

        static public uint ELE_INVALID_OBJECT_HANDLE = 0x00008906;// 无效的对象句柄

        static public uint ELE_DEVICE_NOT_IN_WHITE_LIST = 0x00008a01;// 升级的设备在白名单中不存在

        static public uint ELE_BLOCK_CIPHER_ERROR = 0x00008a02;// 升级block的密钥错误

        static public uint ELE_BLOCK_SIGNATURE_ERROR = 0x00008a03;// 升级block的签名错误

        static public uint ELE_BLOCK_MISMATCH = 0x00008a04;// 升级block不匹配

        static public uint ELE_ERROR_UNKNOWN = 0x00008fff;// 未知错误

        static public uint ELE_ELC_ERROR_SES_EEPROM = 0x00009001;// 写精锐E硬件产生错误
        static public uint ELE_ELC_ERROR_SES_UNSUPPORT = 0x00009002;// 设备功能不支持

        static public uint ELE_ELC_ERROR_SES_RTC = 0x00009003;// 时钟锁错误

        static public uint ELE_ELC_ERROR_SES_RTC_POWER = 0x00009004;// 时钟锁电源已断电
        static public uint ELE_ELC_ERROR_SES_MEMORY = 0x00009201;// 内存地址越界
        static public uint ELE_ELC_ERROR_SES_PARAM = 0x00009204;// SES参数错误
        static public uint ELE_ELC_ERROR_SES_OBJECT = 0x00009901;// 对象不存在  


        static public uint ELE_VM_W_CODE_RANGE = 0x0000A001;// ROM地址越界,
        static public uint ELE_VM_W_INST_RSV = 0x0000A002;// 非法的指令编码，
        static public uint ELE_VM_W_IDATA_RANGE = 0x0000A003;// RAM地址越界，

        static public uint ELE_VM_W_BIT_RANGE = 0x0000A004;// BIT地址越界，

        static public uint ELE_VM_W_SFR_RANGE = 0x0000A005;// SFR地址非法，

        static public uint ELE_VM_W_XRAM_RANGE = 0x0000A006;// XRAM地址越界，




        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]

        public struct ELE_DEVICE_CONTEXT
        {
            public uint ulSize;                    // 这里需要设为sizeof(ELE_DEVICE_CONTEXT)
            public uint ulFinger;                  // 标识，内部填充


            public uint ulMask;                    // 表示那些条件有效1，2，4分别表示开发商编号，厂商描述，序列号


            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] ucDevNumber;             // 保存开发商编号
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] ucDesp;                 // 保存厂商描述
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] ucSerialNumber;         // 保存序列号


            public uint ulShareMode;               // 设备打开方式，共享模式或是独占模式


            public uint ulIndex;                   // 这里的Index值表示当前打开的锁在全部设备（无论是不是符合条件的）列表中的位置


            public uint ulDriverMode;              // 驱动模式 有驱动模式或是无驱模式


            public int hDevice;           // 设备句柄
            public int hMutex;            // 互斥体句柄


        }

        public struct ELE_LIBVERSIONINFO
        {
            public uint ulVersionInfoSize;          // 本结构体的大小，传给函数前要设置为sizeof(ELE_LIBVERSIONINFO)
            public uint ulMajorVersion;            // 主版本号，本公版版本中设置为2
            public uint ulMinorVersion;            // 次版本号，本公版版本中设置为0
            public uint ulClientID;                // 客户定制的编号，对于公版产品，此号码为0
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 128)]
            public sbyte[] acDesp;               // 对该库的补充描述字符串


        }



        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleSign(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] pucSerial, byte[] pcModuleName, byte[] pucModuleContent, uint ulModuleSize, byte[] pPkgBuffer, uint ulPkgBufferLen, ref uint pulActualPkgLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleUpdate(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] pucPkgContent, uint ulPkgLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleChangePin(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] pucOldPin, byte[] pucNewPin);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleReadModule(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] pcModuleName, byte[] pucModuleBuffer, uint ulModuleBufferLen, ref uint pulActuralModuleLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleChangeModuleName(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] pcOldModuleName, byte[] pcNewModuleName);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleGetFirstModuleName(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] pcModuleNameBuffer, uint ulModuleNameBufferLen, ref uint pulModuleNameLen, ref uint pulIndex);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleNextModuleName(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] pcModuleNameBuffer, uint ulModuleNameBufferLen, ref uint pulModuleNameLen, ref uint pulIndex);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleOpenFirstDevice(byte[] DeviceNumber, byte[] Desc, byte[] SerialNumber, uint ShareMode, ref ELE_DEVICE_CONTEXT EleDeviceContext);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleVerifyPin(ref ELE_DEVICE_CONTEXT EleDeviceContext, byte[] Pin);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleWriteModule(ref ELE_DEVICE_CONTEXT EleDeviceContext, string ModuleName, byte[] ModuleContent, uint ModuleContentLen, ref uint ActualWrittenLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        public static extern bool EleExecute(ref ELE_DEVICE_CONTEXT EleDeviceContext, string ModuleName, byte[] Input, uint InputLen, byte[] Output, uint OutputLen, ref uint ActualOutputLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleClose(ref ELE_DEVICE_CONTEXT EleDeviceContext);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern uint EleGetLastError();
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleOpenNextDevice(ref ELE_DEVICE_CONTEXT EleDeviceContext);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleGetDeviceInfo(ref ELE_DEVICE_CONTEXT EleDeviceContext, uint ulFlag, ref int pBuffer, uint ulBufferLen, ref uint pulActualLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleSetDevieInfo(ref ELE_DEVICE_CONTEXT EleDeviceContext, uint ulFlag, ref int pBuffer, uint ulBufferLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleControl(ref ELE_DEVICE_CONTEXT EleDeviceContext, uint ulCtrlCode, byte[] pucInput, uint ulInputLen, byte[] pucOutput, uint ulOutputLen, ref uint pulActualLen);
        [DllImport("elitee.dll", CharSet = CharSet.Auto)]
        private static extern bool EleGetVersion(ref ELE_LIBVERSIONINFO pEleLibVersionInfo);



        static public uint ELE_T2_SUCCESS = 0x000020000;
        static public uint ELE_T2_NO_MORE_DEVICE = 0x00020001;
        static public uint ELE_T2_INVALID_PASSWORD = 0x00020002;
        static public uint ELE_T2_INSUFFICIENT_BUFFER = 0x00020003;
        static public uint ELE_T2_BEYOND_DATA_SIZE = 0x00020004;

        static public uint MAXLENGTH = 16;

        public string SenseLock = "";
        public string strModal = "";
        public string sVersion = "";

        public uint iStart = 0, iLength = 16;

        [DllImport("elet2.dll", CharSet = CharSet.Auto)]
        public static extern uint EleT2Read(uint usOffset, byte[] pucOutbuffer, uint usOutbufferLen, ref uint usReadLen);
        [DllImport("elet2.dll", CharSet = CharSet.Auto)]
        public static extern uint EleT2Write(int usOffset, byte[] pcPassword, byte[] pucInbuffer, int usInbufferLen, byte[] usWrittenLen);
        [DllImport("elet2.dll", CharSet = CharSet.Auto)]
        public static extern uint EleT2ChangePassword(byte[] pcOldPass, byte[] pcNewPass);

        public int ReadSenseLock()
        {
            if (strModal == "")
                return 1;

            uint ulErrRet = 0;
            bool bRet = false;

            ELE_DEVICE_CONTEXT edc = new ELE_DEVICE_CONTEXT();
            edc.ulSize = (uint)Marshal.SizeOf(typeof(ELE_DEVICE_CONTEXT));

            //open first elite e
            bRet = EleOpenFirstDevice(null, null, null, ELE_SHARE_MODE, ref edc);
            if (!bRet)
            {
                ulErrRet = EleGetLastError();
                return -1;
            }

            //verify pin of the found device   
            byte[] Pin = { (byte)'1', (byte)'3', (byte)'5', (byte)'7', (byte)'9', (byte)'2', (byte)'4', (byte)'6', (byte)'8', (byte)'0', (byte)'8', (byte)'8', (byte)'6', (byte)'6', (byte)'2', (byte)'2' };

            bRet = EleVerifyPin(ref edc, Pin);
            if (!bRet)
            {
                ulErrRet = EleGetLastError();
                return -1;
            }

            //close device
            bRet = EleClose(ref edc);
            if (!bRet)
            {
                ulErrRet = EleGetLastError();
                return -1;
            }



            byte[] rbuff = new byte[MAXLENGTH];
            uint rbLen = 0;

            uint uLRet = 0;


            uLRet = EleT2Read(iStart, rbuff, iLength, ref rbLen);

            if (uLRet != ELE_T2_SUCCESS)
                return 1;

            System.Text.UTF8Encoding u8Encoding = new UTF8Encoding();
            SenseLock = u8Encoding.GetString(rbuff);

            int i = SenseLock.IndexOf(strModal);
            if (i == -1)
                return -1;

            if (i > SenseLock.Length - 1)
            {
                sVersion = "";
                return 0;
            }
            else
            {
                sVersion = SenseLock.Substring(i + strModal.Length, SenseLock.Length - strModal.Length);
                try
                {
                    sVersion = Int32.Parse(sVersion).ToString();
                }
                catch
                {
                    sVersion = "";
                }
            }
            return 0;


        }

    }
}
