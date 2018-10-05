using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.PointOfService;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;
using System.ComponentModel;

namespace PosDeviceHelperLib
{
    public class PosDeviceHelper
    {
        PosExplorer _posExplorer;
        PosCommon _posCommonScanner;
        PosCommon _posCommonScale;
        PosCommon _posCommonCashDrawer;
        PosCommon _posCommonPrinter;
        Control _cDataCode;
        Control _cDataWeight;
        Scanner _Scanner;
        Scale _Scale;
        CashDrawer _CashDrawer;
        PosPrinter _PosPrinter;
        int iTimeOut;
        string sBarCodeLabel = "";
        double dbWeightLabel = 0.0;
        string sNumberFormat;
        string sMessengeError;
        int iCount = 0;
        Thread CurrentPosThread;
        public PosDeviceHelper(Form frm)
        {
            _posExplorer = new PosExplorer(frm);
            iTimeOut = 1000;
            sNumberFormat = "1000";
        }
        public string DataCodeLabel
        {
            get { return sBarCodeLabel; }
        }
        public double DataWeightLabel
        {
            get
            {
                return dbWeightLabel;
            }
        }
        public string MessengeError
        {
            get { return sMessengeError; }
        }
        public void ResetDataCodeLabel()
        {
            sBarCodeLabel = string.Empty;
        }
        public void LoadDevice(DeviceLevel strDeviceLevelName, Dictionary<string, string> dicDrivers, Control cDataCode = null, Control cWeight = null)
        {
            try
            {
                foreach (KeyValuePair<string, string> item in dicDrivers)
                {
                    switch (item.Key)
                    {
                        case "Scanner":
                            LoadScannerDevice(strDeviceLevelName, item.Value, cDataCode);
                            break;
                        case "Scale":
                            LoadScaleDevice(strDeviceLevelName, item.Value, cWeight);
                            break;
                        case "CashDrawer":
                            LoadCashDrawerDevice(strDeviceLevelName, item.Value);
                            break;
                        case "Printer":
                            LoadPosPrinterDevice(strDeviceLevelName, item.Value);
                            break;
                    }
                }
                ClaimPosCommon();
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        public void LoadDevice(DeviceLevel strDeviceLevelName, string strScannerName, string strScaleName, string strCashDrawerName, string strPosPrinterName, Control cDataCode = null, Control cWeight = null)
        {
            try
            {
                LoadScannerDevice(strDeviceLevelName, strScannerName, cDataCode);
                LoadScaleDevice(strDeviceLevelName, strScaleName, cWeight);
                LoadCashDrawerDevice(strDeviceLevelName, strCashDrawerName);
                LoadPosPrinterDevice(strDeviceLevelName, strPosPrinterName);
                ClaimPosCommon();
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        public void LoadScannerDevice(DeviceLevel strDeviceLevelName, string strScannerName, Control cDataCode = null)
        {
            try
            {
                OpenDevice(ref _posCommonScanner, strScannerName, Convert.ToString(strDeviceLevelName));
                if (_Scanner == null)
                {
                    if (_posCommonScanner != null)
                    {
                        _Scanner = (Scanner)_posCommonScanner;
                        _Scanner.DataEvent += new DataEventHandler(_Scanner_DataEvent);
                        _Scanner.DecodeData = true;
                    }
                }
                if (_cDataCode == null)
                    _cDataCode = cDataCode;
            }
            catch (Exception ex)
            {
                //sMessengeError = ex.Message;
                throw ex;
            }
        }
        public void LoadScaleDevice(DeviceLevel strDeviceLevelName, string strScaleName, Control cWeight = null)
        {
            try
            {
                OpenDevice(ref _posCommonScale, strScaleName, Convert.ToString(strDeviceLevelName));
                CurrentPosThread = new Thread(new ThreadStart(() =>
                {
                    if (_Scale == null)
                    {
                        if (_posCommonScale != null)
                        {
                            _Scale = (Scale)_posCommonScale;
                            _Scale.DataEvent += new DataEventHandler(_Scale_DataEvent);
                            _Scale.StatusUpdateEvent += new StatusUpdateEventHandler(_Scale_StatusUpdateEvent);
                            //_Scale.StatusNotify = StatusNotify.Enabled;
                        }
                    }
                }));
                CurrentPosThread.IsBackground = true;
                CurrentPosThread.Start();
                if (_cDataWeight == null)
                    _cDataWeight = cWeight;
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        public void LoadCashDrawerDevice(DeviceLevel strDeviceLevelName, string strCashDrawerName)
        {
            try
            {
                OpenDevice(ref _posCommonCashDrawer, strCashDrawerName, Convert.ToString(strDeviceLevelName));
                if (_CashDrawer == null)
                {
                    if (_posCommonCashDrawer != null)
                    {
                        _CashDrawer = (CashDrawer)_posCommonCashDrawer;
                        _CashDrawer.DeviceEnabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        public void LoadPosPrinterDevice(DeviceLevel strDeviceLevelName, string strPosPrinterName)
        {
            try
            {
                OpenDevice(ref _posCommonPrinter, strPosPrinterName, Convert.ToString(strDeviceLevelName));
                if (_PosPrinter == null)
                {
                    if (_posCommonPrinter != null)
                    {
                        _PosPrinter = (PosPrinter)_posCommonPrinter;
                    }
                }
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        private void OpenDevice(ref PosCommon posCommon, string strDeviceName, string strDeviceLevel)
        {
            if (string.IsNullOrEmpty(strDeviceLevel) || string.IsNullOrEmpty(strDeviceName)) return;
            DeviceCollection devices = _posExplorer.GetDevices((DeviceCompatibilities)Enum.Parse(typeof(DeviceCompatibilities), strDeviceLevel, false));
            DeviceInfo drv = devices.Cast<DeviceInfo>().Where(x => x.ServiceObjectName.Equals(strDeviceName)).FirstOrDefault();
            try
            {
                posCommon = (PosCommon)_posExplorer.CreateInstance(drv);
                posCommon.Open();
                SetDataEventEnabledPosCommon(ref posCommon, true);
                HookUpEvents(posCommon, true);
            }
            catch (Exception ex)
            {
                //sMessengeError = ex.Message;
                throw ex;
            }
        }
        private void SetDataEventEnabledPosCommon(ref PosCommon posCommon, bool isEnabled)
        {
            try
            {
                PropertyInfo dataEventEnabled = posCommon.GetType().GetProperty("DataEventEnabled");
                if (dataEventEnabled != null)
                    dataEventEnabled.SetValue(posCommon, isEnabled, null);
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        private void HookUpEvents(PosCommon pc, bool attach)
        {
            EventInfo dataEvent = pc.GetType().GetEvent("DataEvent");
            if (dataEvent != null)
            {
                if (attach)
                    dataEvent.AddEventHandler(pc, new DataEventHandler(co_OnDataEvent));
                else
                    dataEvent.RemoveEventHandler(pc, new DataEventHandler(co_OnDataEvent));
            }
        }
        private void ClaimPosCommon()
        {
            try
            {
                if (_posCommonScanner != null)
                {
                    _posCommonScanner.Claim(iTimeOut);
                    _posCommonScanner.DeviceEnabled = true;
                }
                if (_posCommonScale != null)
                {
                    _posCommonScale.Claim(iTimeOut);
                    ((Scale)_posCommonScale).StatusNotify = StatusNotify.Enabled;
                    _posCommonScale.DeviceEnabled = true;
                }
                if (_CashDrawer != null)
                    _CashDrawer.Claim(iTimeOut);
                if (_PosPrinter != null)
                    _PosPrinter.Claim(iTimeOut);
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        private void SetValueControl(Control ctrl, object oValue)
        {
            if (ctrl == null) return;
            ctrl.Text = Convert.ToString(oValue);
        }

        #region Events
        private void co_OnDataEvent(object sender, DataEventArgs e)
        {
            PosCommon posCommon = sender as PosCommon;
            SetDataEventEnabledPosCommon(ref posCommon, true);
        }
        private void _Scale_DataEvent(object sender, DataEventArgs e)
        {
        }
        private void _Scale_StatusUpdateEvent(object sender, StatusUpdateEventArgs e)
        {
            Thread.Sleep(1000);
            double dbVal = GetScaleWeight();
            if (dbVal > 0)
            {
                SetValueControl(_cDataWeight, dbVal);
                iCount = 0;
                //Thread.Sleep(5000);
            }
            else
            {
                if (iCount == 0)
                    SetValueControl(_cDataWeight, dbVal);
                iCount++;
            }
        }
        private void _Scanner_DataEvent(object sender, DataEventArgs e)
        {
            byte[] b = _Scanner.ScanDataLabel;
            string str = "";
            for (int i = 0; i < b.Length; i++)
                str += (char)b[i];
            sBarCodeLabel = str;
            SetValueControl(_cDataCode, sBarCodeLabel);
            SetValueControl(_cDataWeight, GetScaleWeight());
        }
        public double GetScaleWeight()
        {
            if (_Scale == null) return 0;
            decimal weight = 0;
            try
            {
                weight = _Scale.ReadWeight(Int32.Parse(sNumberFormat, System.Globalization.CultureInfo.CurrentCulture));
            }
            catch { weight = 0; }
            dbWeightLabel = Convert.ToDouble(weight);
            return dbWeightLabel;
        }
        public void OpenCashDrawerCase()
        {
            try
            {
                if (_CashDrawer != null)
                    _CashDrawer.OpenDrawer();
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        public void Print(string strContent)
        {
            try
            {
                if (_PosPrinter != null)
                    _PosPrinter.PrintNormal(PrinterStation.Receipt, ValidateDataPrint(strContent));
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        private string ValidateDataPrint(string strContent)
        {
            return strContent.Replace("ESC", ((char)27).ToString()) + "\x1B|1lF";
        }
        public void CutReceipt(int iPercentage = 100)
        {
            try
            {
                if (_PosPrinter != null)
                    _PosPrinter.CutPaper(iPercentage);
            }
            catch (Exception ex)
            {
                sMessengeError = ex.Message;
            }
        }
        #endregion
    }
    public enum DeviceLevel
    {
        CompatibilityLevel1,
        Opos,
        OposAndCompatibilityLevel1
    }
}
