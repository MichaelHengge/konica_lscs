import sys
import os
import time
from typing import Dict, List, Optional, Tuple
from datetime import datetime

try:
    import clr
    import System
    from System import AppDomain, DateTime as NetDateTime
    from System.Reflection import Assembly
    from System.IO import Directory
except ImportError:
    raise ImportError("This library requires pythonnet. Please run: pip install pythonnet")


class KonicaDevice:
    """
    Complete controller for Konica Minolta CS-150, CS-160, LS-150, and LS-160 devices.
    
    Implements all major functions from the LC-MISDK Reference Manual.
    
    Basic Usage:
        device = KonicaDevice()
        device.connect()
        device.measure()
        lv = device.get_luminance()
        device.disconnect()
    
    Or use context manager:
        with KonicaDevice() as device:
            device.connect()
            device.measure()
            print(device.get_color())
    """
    
    def __init__(self):
        self.sdk = None
        self._sdk_bin_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'sdk_bin'))
        
        if not os.path.exists(self._sdk_bin_path):
            raise FileNotFoundError(f"SDK binary folder not found at: {self._sdk_bin_path}")
        
        self._original_cwd = os.getcwd()
        os.chdir(self._sdk_bin_path)
        sys.path.insert(0, self._sdk_bin_path)
        
        try:
            self._setup_environment()
            self._initialize_sdk()
        except Exception as e:
            os.chdir(self._original_cwd)
            raise RuntimeError(f"Failed to initialize Konica SDK: {e}")

    def _setup_environment(self):
        """Setup environment for SDK"""
        calc_path = os.path.join(self._sdk_bin_path, "calccolor", "x64")
        os.environ['PATH'] = f"{self._sdk_bin_path};{calc_path};{os.environ.get('PATH', '')}"
        
        try:
            domain = AppDomain.CurrentDomain
            domain.SetData("APPBASE", self._sdk_bin_path)
            Directory.SetCurrentDirectory(self._sdk_bin_path)
        except:
            pass

    def _initialize_sdk(self):
        """Initialize the Konica SDK"""
        clr.AddReference(os.path.join(self._sdk_bin_path, "LC-MISDK.dll"))
        
        from Konicaminolta import (
            LightColorMISDK, MeasStatus, XYZ, Lvxy, Lvudvd, LvTcpDuv, LvDwPe, Lv,
            LuminanceUnit, ErrorDefine, MeasurementData, MeasurementTime, MeasTimeMode,
            MeasurementFrequency, SyncMode, PeakValley, CloseUpLensType, CalibType,
            ColorCorrectionFactor, CCFMode, AutoPowerOff, BackLightMode, BackLightLevel,
            DispType, DisplayLanguage, DateFormat, ColorMode, ColorModeDisplay,
            DataSaveMode, PeriodicCalibNotify, ToggleStatus, TriggerStatus, DeviceInfo,
            UserCalibData
        )
        
        self.types = {
            'MeasStatus': MeasStatus,
            'XYZ': XYZ,
            'Lvxy': Lvxy,
            'Lvudvd': Lvudvd,
            'LvTcpDuv': LvTcpDuv,
            'LvDwPe': LvDwPe,
            'Lv': Lv,
            'Unit': LuminanceUnit,
            'Error': ErrorDefine,
            'MeasurementData': MeasurementData,
            'MeasurementTime': MeasurementTime,
            'MeasTimeMode': MeasTimeMode,
            'MeasurementFrequency': MeasurementFrequency,
            'SyncMode': SyncMode,
            'PeakValley': PeakValley,
            'CloseUpLensType': CloseUpLensType,
            'CalibType': CalibType,
            'ColorCorrectionFactor': ColorCorrectionFactor,
            'CCFMode': CCFMode,
            'AutoPowerOff': AutoPowerOff,
            'BackLightMode': BackLightMode,
            'BackLightLevel': BackLightLevel,
            'DispType': DispType,
            'DisplayLanguage': DisplayLanguage,
            'DateFormat': DateFormat,
            'ColorMode': ColorMode,
            'ColorModeDisplay': ColorModeDisplay,
            'DataSaveMode': DataSaveMode,
            'PeriodicCalibNotify': PeriodicCalibNotify,
            'ToggleStatus': ToggleStatus,
            'TriggerStatus': TriggerStatus,
            'DeviceInfo': DeviceInfo,
            'UserCalibData': UserCalibData
        }
        
        self.sdk = LightColorMISDK.GetInstance()

        self.error_map = {
            0: "The processing was completed normally (KmSuccess)",
            10: "No instrument is connected to the specified virtual COM port (ErNoConnect)",
            25: "The assigned parameter is incorrect (ErInvalidParameter)",
            30: "This model does not support the specified command (ErCannotCommand)",
            45: "No specified data (ErNoData)",
            50: "Calculation result is out of the support range of chromaticity or luminance (ErOutOfRangeValue)",
            60: "The instrument cannot receive any commands because it is in the middle of a process (ErInstrumentProcessing)",
            100: "Failed to connect to the instrument (ErConnectFailed)",
            110: "There are one or more ports that cannot be disconnected (ErDisConnectFailed)",
            120: "Failed to set the setting (ErSetFailed)",
            130: "Failed to obtain the setting (ErGetFailed)",
            140: "Calculation failed (ErCalcFailed)",
            150: "Cancellation failed (ErCancelFailed)",
            160: "Write failed (ErWriteFailed)",
            170: "Read failed (ErReadFailed)",
            180: "Failed to delete the setting(s) (ErDeleteFailed)",
            200: "Measurement failed (ErMeasurementFailed)",
        }

    def _check_error(self, ret, operation: str = "Operation"):
        """Check return code and raise exception if error, providing a detailed description."""
        error_code = ret.errorCode
        
        if error_code != self.types['Error'].KmSuccess:
            base_message = self.error_map.get(error_code, "Unknown error code")
            sdk_messages = ', '.join(ret.errorMessage) if hasattr(ret, 'errorMessage') and ret.errorMessage else ''
            
            if sdk_messages:
                full_message = f"{base_message}. SDK Details: {sdk_messages}"
            else:
                full_message = base_message
                
            raise RuntimeError(f"{operation} failed with error {error_code}: {full_message}")

    # ==================== CONNECTION ====================
    
    def connect(self, com_port: int = 0):
        """
        Connect to the device.
        
        Args:
            com_port: Virtual COM port number. 0 = auto-detect first available device
        """
        if not self.sdk:
            raise RuntimeError("SDK not initialized")
        
        ret = self.sdk.Connect(com_port)
        self._check_error(ret, "Connect")

    def disconnect(self, com_port: int = 0):
        """Disconnect from the device"""
        if self.sdk:
            try:
                ret = self.sdk.DisConnect(com_port)
            except: # COntinue with cleanup even if error
                pass
        
        if self._original_cwd: # env cleanup
            os.chdir(self._original_cwd)

    def get_device_list(self) -> Dict[int, str]:
        """
        Get list of connected devices.
        
        Returns:
            Dictionary mapping COM port number to device string (e.g., "CS-150(12345678)")
        """
        result = self.sdk.GetDeviceList()
        if isinstance(result, tuple):
            ret, device_dict = result
            self._check_error(ret, "GetDeviceList")
            
            devices = {}
            for port in device_dict.Keys:
                devices[int(port)] = str(device_dict[port])
            return devices
        
        raise RuntimeError("Unexpected return from GetDeviceList")

    # ==================== MEASUREMENT ====================
    
    def measure(self, wait: bool = True, com_port: int = 0):
        """
        Take a measurement.
        
        Args:
            wait: If True, waits for measurement to complete
            com_port: COM port number (0 = default)
        """
        ret = self.sdk.Measure(com_port)
        self._check_error(ret, "Measure")
        
        if wait:
            self.wait_for_idle(com_port)
            time.sleep(0.2)

    def polling_measurement(self, com_port: int = 0) -> tuple[str, int]:
        """
        Polls the device to check the current measurement status.
        """
        
        result = self.sdk.PollingMeasurement(comPort=com_port) 
        
        if not isinstance(result, tuple):
            raise RuntimeError("Unexpected return format from PollingMeasurement")

        ret, state = result
        
        self._check_error(ret, "PollingMeasurement")
        
        if state == self.types['MeasStatus'].Idling:
            status_str = 'idling'
        elif state == self.types['MeasStatus'].Measuring:
            status_str = 'measuring'
        else:
            status_str = 'unknown'

        return status_str, ret.errorCode

    def wait_for_idle(self, com_port: int = 0): # Not official KM function
        """Wait for measurement to complete"""
        while True:
            state, _ = self.polling_measurement(com_port)
            
            if state != 'measuring':
                break
                
            time.sleep(0.1)

    def cancel_measurement(self, com_port: int = 0):
        """Cancel ongoing measurement"""
        ret = self.sdk.CancelMeasurement(com_port)
        self._check_error(ret, "CancelMeasurement")

    # ==================== DATA READING ====================

    def read_latest_data(self, color_space_class, com_port: int = 0) -> dict[str, float]: # STILL HAS ISSUES
        """
        Read the latest data in a specified color space.
        
        Args:
            color_space_class: The desired color space class instance (e.g., self.types['XYZ']()).
            com_port: COM port number (0 = default).
            
        Returns:
            Dictionary with color values.
        """
        data = color_space_class
        ret = self.sdk.ReadLatestData(data, com_port)
        self._check_error(ret, f"ReadLatestData (for {data.__class__.__name__})")
        
        values = {}
        if hasattr(data, 'ColorSpaceValue'):
            for key in data.ColorSpaceValue.Keys:
                values[str(key)] = float(data.ColorSpaceValue[key])
        
        return values
    
    def read_display_value(self, com_port: int = 0) -> Dict[str, float]: # USE THIS RATHER THAN READ_LATEST_DATA
        """
        Read the value currently displayed on the device.
        This is the most reliable method for reading measurements.
        
        Returns:
            Dictionary with color values (keys depend on current color mode)
        """
        result = self.sdk.ReadDisplayValue()
        
        if not isinstance(result, tuple):
            raise RuntimeError("Unexpected return from ReadDisplayValue")
        
        ret, display_data = result
        self._check_error(ret, "ReadDisplayValue")
        
        values = {}
        if hasattr(display_data, 'ColorSpaceValue'):
            for key in display_data.ColorSpaceValue.Keys:
                values[str(key)] = float(display_data.ColorSpaceValue[key])
        
        return values
    
    # ------------------- Additional helpful functions (not official) -------------------

    def read_latest_data_xyz(self, com_port: int = 0) -> dict[str, float]:
        """
        Read latest measurement as XYZ tristimulus values (raw reading).
        
        Returns:
            Dictionary with keys "X", "Y", "Z".
        """
        data_instance = self.types['XYZ']()
        
        return self.read_latest_data(data_instance, com_port)

    def get_luminance(self, com_port: int = 0) -> float:
        """
        Get luminance value in cd/m².
        
        Returns:
            Luminance value
        """
        values = self.read_display_value(com_port)
        if 'Lv' in values:
            return values['Lv']
        elif 'Y' in values:
            return values['Y']
        else:
            raise ValueError("No luminance value found in measurement")

    def get_color(self, com_port: int = 0) -> Dict[str, float]:
        """
        Get all available color values from last measurement.
        
        Returns:
            Dictionary with all color values
        """
        return self.read_display_value(com_port)

    # ==================== STORED MEASUREMENT DATA ====================
    
    def get_number_of_samples(self, com_port: int = 0) -> int:
        """Get number of measurements stored in device"""
        result = self.sdk.GetNumberOfSampleData(com_port)
        if isinstance(result, tuple):
            ret, count = result
            self._check_error(ret, "GetNumberOfSampleData")
            return int(count)
        raise RuntimeError("Unexpected return from GetNumberOfSampleData")

    def read_sample_data(self, data_number: int, com_port: int = 0) -> Dict[str, float]: # Only one version implemented yet
        """
        Read stored measurement data from device.
        
        Args:
            data_number: Sample number (1-based index)
            com_port: COM port number
            
        Returns:
            Dictionary with color values
            
        Raises:
            RuntimeError: If no data exists or read fails
        """

        data = self.types['XYZ']()
        ret = self.sdk.ReadSampleData(data_number, data, com_port)
        
        if ret.errorCode == self.types['Error'].KmSuccess:
            return {"X": data.X, "Y": data.Y, "Z": data.Z}
        elif ret.errorCode == self.types['Error'].ErNoData:
            data_lvxy = self.types['Lvxy'](self.types['Unit'].cdm2)
            ret = self.sdk.ReadSampleData(data_number, data_lvxy, com_port)
            
            if ret.errorCode == self.types['Error'].KmSuccess:
                return {"Lv": data_lvxy.Lv, "x": data_lvxy.x, "y": data_lvxy.y}
        
        # If we get here, reading failed
        self._check_error(ret, "ReadSampleData")

    def delete_sample_data(self, data_number: int = -1, com_port: int = 0):
        """
        Delete stored measurement data.
        
        Args:
            data_number: Sample number to delete, or -1 to delete all
            com_port: COM port number
        """
        ret = self.sdk.DeleteSampleData(data_number, com_port)
        self._check_error(ret, "DeleteSampleData")

    # ==================== TARGET VALUES (NOT TESTED YET) ====================

    def set_target_channel(self, target_channel: int, com_port: int = 0):
        """
        Sets the target channel to be used by the instrument.
        
        Note: The target channel must already contain a saved target value.
        
        Args:
            target_channel: The target channel number.
            com_port: Virtual COM Port number (0 = default).
        """
        ret = self.sdk.SetTargetCh(target_channel, com_port)
        self._check_error(ret, "SetTargetCh")

    def get_target_channel(self, com_port: int = 0) -> int:
        """
        Obtains the currently specified target channel setting from the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current target channel number (Int32).
        """
        result = self.sdk.GetTargetCh(comPort=com_port)
        
        if not isinstance(result, tuple):
            raise RuntimeError("Unexpected return format from GetTargetCh")
            
        ret, target_channel = result
        self._check_error(ret, "GetTargetCh")
        
        return int(target_channel)
    
    def read_target_data(self, target_channel: int, color_space: str = 'Lvxy', com_port: int = 0) -> dict[str, float]:
        """
        Reads the target data stored in a specific channel in the desired color space.
        
        Args:
            target_channel: Target channel number (Int32).
            color_space: The desired color space ('XYZ', 'Lvxy', 'Lvudvd', etc.). Defaults to 'Lvxy'.
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            Dictionary with color values for the target data.
        """
        if color_space not in self.types:
            raise ValueError(f"Invalid color_space '{color_space}'. Must be one of: {list(self.types.keys())}")
            
        if color_space == 'Lvxy':
            data_instance = self.types['Lvxy'](self.types['Unit'].cdm2)
        else:
            data_instance = self.types[color_space]()

        ret = self.sdk.ReadTargetData(target_channel, data_instance, com_port)
        self._check_error(ret, f"ReadTargetData (channel {target_channel})")
        
        values = {}
        if hasattr(data_instance, 'ColorSpaceValue'):
            for key in data_instance.ColorSpaceValue.Keys:
                values[str(key)] = float(data_instance.ColorSpaceValue[key])
        
        return values
    
    def write_target_data(self, target_channel: int, target_values: dict[str, any], com_port: int = 0):
        """
        Writes target data (color and ID) to the specified channel.
        
        Args:
            target_channel: Target channel number (Int32).
            target_values: Dictionary containing 'Lv', 'x', 'y' values and optional 'Id'.
            com_port: Virtual COM Port number (0 = default).
        """
        
        Lvxy = self.types['Lvxy']
        
        lvxy_target = Lvxy(
            self.types['Unit'].cdm2,
            target_values.get('Lv', 0.0),
            target_values.get('x', 0.0),
            target_values.get('y', 0.0)
        )
        
        if 'Id' in target_values:
            lvxy_target.Id = str(target_values['Id'])
            
        ret = self.sdk.WriteTargetData(target_channel, lvxy_target, com_port)
        self._check_error(ret, f"WriteTargetData (channel {target_channel})")

    def delete_target_data(self, target_channel: int = -1, com_port: int = 0):
        """
        Deletes the target value stored in the instrument.
        
        Args:
            target_channel: Target channel to delete. Use -1 to delete all target data.
            com_port: Virtual COM Port number (0 = default).
        """
        ret = self.sdk.DeleteTargetData(target_channel, com_port)
        self._check_error(ret, "DeleteTargetData")

    # ==================== MEASUREMENT SETTINGS ====================
    
    def set_measurement_time(self, mode: str, manual_time: float = 1.0, com_port: int = 0):
        """
        Set measurement time.
        
        Args:
            mode: 'auto' or 'manual'
            manual_time: Integration time in seconds (if mode='manual')
            com_port: COM port number
        """
        meas_time = self.types['MeasurementTime']()
        
        if mode.lower() == 'auto':
            meas_time.MeasTimeMode = self.types['MeasTimeMode'].Auto
        elif mode.lower() == 'manual':
            meas_time.MeasTimeMode = self.types['MeasTimeMode'].Manual
            meas_time.ManualMeasurementTime = manual_time
        else:
            raise ValueError("mode must be 'auto' or 'manual'")
        
        ret = self.sdk.SetMeasurementTime(meas_time)
        self._check_error(ret, "SetMeasurementTime")

    def get_measurement_time(self, com_port: int = 0) -> Tuple[str, float]:
        """
        Get measurement time settings.
        
        Returns:
            Tuple of (mode, manual_time)
        """
        result = self.sdk.GetMeasurementTime()
        if isinstance(result, tuple):
            ret, meas_time = result
            self._check_error(ret, "GetMeasurementTime")
            
            mode = "auto" if meas_time.MeasTimeMode == self.types['MeasTimeMode'].Auto else "manual"
            return mode, meas_time.ManualMeasurementTime
        
        raise RuntimeError("Unexpected return from GetMeasurementTime")

    def set_sync_mode(self, sync: bool, frequency: float = 60.0, com_port: int = 0):
        """
        Set synchronization mode.
        
        Args:
            sync: True for sync ON, False for sync OFF
            frequency: Sync frequency in Hz (if sync=True)
            com_port: COM port number
        """
        sync_data = self.types['MeasurementFrequency']()
        sync_data.SyncMode = self.types['SyncMode'].Sync if sync else self.types['SyncMode'].Async
        sync_data.Frequency = frequency
        
        ret = self.sdk.SetSyncMode(sync_data)
        self._check_error(ret, "SetSyncMode")

    def get_sync_mode(self, com_port: int = 0) -> Tuple[bool, float]:
        """
        Get synchronization mode.
        
        Returns:
            Tuple of (sync_enabled, frequency)
        """
        result = self.sdk.GetSyncMode()
        if isinstance(result, tuple):
            ret, sync_data = result
            self._check_error(ret, "GetSyncMode")
            
            sync = sync_data.SyncMode == self.types['SyncMode'].Sync
            return sync, sync_data.Frequency
        
        raise RuntimeError("Unexpected return from GetSyncMode")

    def set_peak_valley(self, mode: str, com_port: int = 0):
        """
        Set peak/valley measurement mode.
        
        Args:
            mode: 'off', 'peak', or 'valley'
            com_port: COM port number
        """
        mode_map = {
            'off': self.types['PeakValley'].OFF,
            'peak': self.types['PeakValley'].Peak,
            'valley': self.types['PeakValley'].Valley
        }
        
        if mode.lower() not in mode_map:
            raise ValueError("mode must be 'off', 'peak', or 'valley'")
        
        ret = self.sdk.SetPeakValley(mode_map[mode.lower()])
        self._check_error(ret, "SetPeakValley")

    def get_peak_valley(self, com_port: int = 0) -> str:
        """Get peak/valley mode"""
        result = self.sdk.GetPeakValley()
        if isinstance(result, tuple):
            ret, pv = result
            self._check_error(ret, "GetPeakValley")
            
            if pv == self.types['PeakValley'].OFF:
                return 'off'
            elif pv == self.types['PeakValley'].Peak:
                return 'peak'
            else:
                return 'valley'
        
        raise RuntimeError("Unexpected return from GetPeakValley")

    # ==================== LENS SETTINGS ====================
    
    def set_close_up_lens(self, lens_type: str, com_port: int = 0):
        """
        Set close-up lens type.
        
        Args:
            lens_type: 'standard', 'no153', 'no135', 'no122', or 'no110'
            com_port: COM port number
        """
        lens_map = {
            'standard': self.types['CloseUpLensType'].Standard,
            'no153': self.types['CloseUpLensType'].No153,
            'no135': self.types['CloseUpLensType'].No135,
            'no122': self.types['CloseUpLensType'].No122,
            'no110': self.types['CloseUpLensType'].No110
        }
        
        if lens_type.lower() not in lens_map:
            raise ValueError(f"Invalid lens_type. Must be one of: {list(lens_map.keys())}")
        
        ret = self.sdk.SetCloseUpLens(lens_map[lens_type.lower()])
        self._check_error(ret, "SetCloseUpLens")

    def get_close_up_lens(self, com_port: int = 0) -> str:
        """Get current close-up lens setting"""
        result = self.sdk.GetCloseUpLens()
        if isinstance(result, tuple):
            ret, lens = result
            self._check_error(ret, "GetCloseUpLens")
            
            lens_map = {
                self.types['CloseUpLensType'].Standard: 'standard',
                self.types['CloseUpLensType'].No153: 'no153',
                self.types['CloseUpLensType'].No135: 'no135',
                self.types['CloseUpLensType'].No122: 'no122',
                self.types['CloseUpLensType'].No110: 'no110'
            }
            return lens_map.get(lens, 'unknown')
        
        raise RuntimeError("Unexpected return from GetCloseUpLens")
    
    # ==================== USER CALIBRATION (NOT TESTED YET) ====================

    def set_calibration_channel(self, channel: int, com_port: int = 0):
        """
        Specifies the user calibration channel to be used for measurements.
        
        Channel 0 means no user calibration is applied.
        
        Args:
            channel: The user calibration channel number (Int32).
            com_port: Virtual COM Port number (0 = default).
        """
        ret = self.sdk.SetCalibrationCh(channel, com_port)
        self._check_error(ret, "SetCalibrationCh")

    def get_calibration_channel(self, com_port: int = 0) -> int:
        """
        Obtains the currently specified user calibration channel.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current user calibration channel number (Int32).
        """
        result = self.sdk.GetCalibrationCh(comPort=com_port)
        
        if not isinstance(result, tuple):
            raise RuntimeError("Unexpected return format from GetCalibrationCh")
            
        ret, calibration_channel = result
        self._check_error(ret, "GetCalibrationCh")
        
        return int(calibration_channel)
    
    def set_matrix_calibration( # NEEDS GOOD CHECKING. DO NOT USE FOR NOW
        self, 
        channel: int, 
        measured_values: list[dict[str, float]], 
        true_values: list[dict[str, float]], 
        calib_type: str, 
        id_str: str, 
        com_port: int = 0
    ):
        """
        Sets the user calibration coefficients (matrix) on the instrument.
        
        Args:
            channel: User calibration channel (1-10).
            measured_values: List of dicts (e.g., [{'Lv': 10.0, 'x': 0.3, 'y': 0.3}]).
            true_values: List of dicts corresponding to the measured_values.
            calib_type: 'onepoint', 'rgb', or 'wrgb'.
            id_str: ID string (max 12 characters).
            com_port: Virtual COM Port number (0 = default).
            
        Raises:
            ValueError: If input format or list lengths are incorrect.
            RuntimeError: If calibration fails (e.g., inverse matrix does not exist).
        """
        
        type_map = {
            'onepoint': self.types['CalibType'].OnePoint,
            'rgb': self.types['CalibType'].RGB,
            'wrgb': self.types['CalibType'].WRGB
        }
        
        calib_enum = type_map.get(calib_type.lower())
        if not calib_enum:
            raise ValueError(f"Invalid calib_type: must be one of {list(type_map.keys())}")

        Lvxy = self.types['Lvxy']
        Unit = self.types['Unit']
        
        def to_lvxy_list(data_list):
            csharp_list = System.Collections.Generic.List[Lvxy]()
            for data in data_list:
                lvxy_instance = Lvxy(
                    Unit.cdm2, 
                    data.get('Lv', 0.0), 
                    data.get('x', 0.0), 
                    data.get('y', 0.0)
                )
                csharp_list.Add(lvxy_instance)
            return csharp_list
            
        # Convert data lists
        csharp_measured = to_lvxy_list(measured_values)
        csharp_true = to_lvxy_list(true_values)

        ret = self.sdk.SetMatrixCalib(
            channel, 
            csharp_measured, 
            csharp_true, 
            calib_enum, 
            id_str, 
            com_port
        )
        self._check_error(ret, f"SetMatrixCalib ({calib_type})")

    def get_calibration_data(self, channel: int, com_port: int = 0) -> dict:
        """
        Obtains the calibration parameters (measured values, true values, and coefficients)
        set for the specified user calibration channel.
        
        Args:
            channel: User calibration channel (Int32).
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            A dictionary containing structured calibration data.
        """
        result = self.sdk.GetCalibData(channel, comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 4:
            raise RuntimeError("Unexpected return format from GetCalibData. Expected 4-element tuple.")
            
        ret, measured_list, true_list, calib_data = result
        self._check_error(ret, f"GetCalibData (channel {channel})")
        
        def extract_data_list(csharp_list):
            py_list = []
            for item in csharp_list:
                data = {str(k): float(v) for k, v in item.ColorSpaceValue.Items}
                data['Id'] = str(item.Id)
                py_list.append(data)
            return py_list

        calib_type_map = {0: 'OnePoint', 1: 'RGB', 2: 'WRGB'}
        
        return {
            'measured_values': extract_data_list(measured_list),
            'true_values': extract_data_list(true_list),
            'user_calib_data': {
                'Id': str(calib_data.Id),
                'Date': str(calib_data.Date), # .NET DateTime object converted to string
                'CalibType': calib_type_map.get(calib_data.CalibType, 'Unknown'),
                'Coef': [float(c) for c in calib_data.Coef] # Calibration coeff.
            }
        }
    
    def delete_calibration_data(self, channel: int = -1, com_port: int = 0):
        """
        Deletes the user calibration coefficients stored in the instrument.
        
        Args:
            channel: Calibration channel to delete. Use -1 to delete all data (default).
            com_port: Virtual COM Port number (0 = default).
        """
        ret = self.sdk.DeleteCalibData(channel, com_port)
        self._check_error(ret, "DeleteCalibData")

    # ==================== COLOR CORRECTION FACTOR (NOT TESTED YET) ====================

    def set_ccf(self, factor_value: float, mode: str, com_port: int = 0):
        """
        Specifies the CCF (Color Correction Factor) setting and turns the mode on or off.
        
        Args:
            factor_value: The Color Correction Factor value (float).
            mode: 'on' or 'off'.
            com_port: Virtual COM Port number (0 = default).
        """
        
        CCFMode = self.types['CCFMode']
        
        if mode.lower() == 'on':
            ccf_mode = CCFMode.ON
        elif mode.lower() == 'off':
            ccf_mode = CCFMode.OFF
        else:
            raise ValueError("mode must be 'on' or 'off'")

        CCF = self.types['ColorCorrectionFactor']
        ccf_data = CCF()
        ccf_data.Coef = factor_value
        ccf_data.CcfMode = ccf_mode
        
        ret = self.sdk.SetCCF(ccf_data, com_port)
        self._check_error(ret, "SetCCF")

    def get_ccf(self, com_port: int = 0) -> dict:
        """
        Obtains the CCF (Color Correction Factor) setting, including the factor value
        and the on/off mode.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            A dictionary with keys 'factor_value' and 'mode' ('on' or 'off').
        """
        result = self.sdk.GetCCF(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetCCF.")
            
        ret, ccf_data = result
        self._check_error(ret, "GetCCF")
        
        if ccf_data.CcfMode == self.types['CCFMode'].ON:
            mode_str = 'on'
        elif ccf_data.CcfMode == self.types['CCFMode'].OFF:
            mode_str = 'off'
        else:
            mode_str = 'unknown'

        return {
            'factor_value': float(ccf_data.Coef),
            'mode': mode_str
        }
    
    # ==================== POWER SETTINGS =====================

    def set_auto_power_off(self, enable: bool, com_port: int = 0):
        """
        Specifies the auto power off setting for the instrument.
        
        Args:
            enable: True to turn auto power off ON, False to turn it OFF.
            com_port: Virtual COM Port number (0 = default).
        """
        AutoPowerOff = self.types['AutoPowerOff']
        
        if enable:
            mode = AutoPowerOff.On
        else:
            mode = AutoPowerOff.Off
            
        ret = self.sdk.SetAutoPowerOff(mode, com_port)
        self._check_error(ret, "SetAutoPowerOff")

    def get_auto_power_off(self, com_port: int = 0) -> bool:
        """
        Obtains the auto power off setting for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            True if auto power off is ON, False if it is OFF.
        """
        result = self.sdk.GetAutoPowerOff(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetAutoPowerOff.")
            
        ret, auto_power_off_mode = result
        self._check_error(ret, "GetAutoPowerOff")
        
        return auto_power_off_mode == self.types['AutoPowerOff'].On

    # ==================== BACKLIGHT SETTINGS ====================
    
    def set_backlight(self, on: bool, com_port: int = 0): # removed _on_off from name for simplicity
        """Turn device backlight on or off"""
        mode = self.types['BackLightMode'].On if on else self.types['BackLightMode'].Off
        ret = self.sdk.SetBackLightOnOff(mode)
        self._check_error(ret, "SetBackLightOnOff")

    def get_backlight(self, com_port: int = 0) -> bool: # removed _on_off from name for simplicity
        """Get backlight status"""
        result = self.sdk.GetBackLightOnOff()
        if isinstance(result, tuple):
            ret, mode = result
            self._check_error(ret, "GetBackLightOnOff")
            return mode == self.types['BackLightMode'].On
        raise RuntimeError("Unexpected return from GetBackLightOnOff")

    def set_backlight_level(self, level: int, com_port: int = 0):
        """
        Set backlight brightness level.
        
        Args:
            level: Brightness level (1-5, where 1=darkest, 5=brightest)
            com_port: COM port number
        """
        if not 1 <= level <= 5:
            raise ValueError("level must be between 1 and 5")
        
        level_map = {
            1: self.types['BackLightLevel'].Level1,
            2: self.types['BackLightLevel'].Level2,
            3: self.types['BackLightLevel'].Level3,
            4: self.types['BackLightLevel'].Level4,
            5: self.types['BackLightLevel'].Level5
        }
        
        ret = self.sdk.SetBackLightLevel(level_map[level])
        self._check_error(ret, "SetBackLightLevel")

    def get_backlight_level(self, com_port: int = 0) -> int:
        """Get backlight brightness level (1-5)"""
        result = self.sdk.GetBackLightLevel()
        if isinstance(result, tuple):
            ret, level = result
            self._check_error(ret, "GetBackLightLevel")
            
            level_map = {
                self.types['BackLightLevel'].Level1: 1,
                self.types['BackLightLevel'].Level2: 2,
                self.types['BackLightLevel'].Level3: 3,
                self.types['BackLightLevel'].Level4: 4,
                self.types['BackLightLevel'].Level5: 5
            }
            return level_map.get(level, 1)
        raise RuntimeError("Unexpected return from GetBackLightLevel")
    
    # ==================== DISPLAY DIGITS ====================

    def set_color_display_digits(self, digit_number: int, com_port: int = 0):
        """
        Specifies the number of color display digits (precision) on the instrument.
        
        Args:
            digit_number: Color display digits (typically 3 or 4).
            com_port: Virtual COM Port number (0 = default).
        """
        if digit_number not in [3, 4]:
            raise ValueError("digit_number must be 3 or 4.")
            
        ret = self.sdk.SetColorDispDigit(digit_number, com_port)
        self._check_error(ret, "SetColorDispDigit")

    def get_color_display_digits(self, com_port: int = 0) -> int:
        """
        Obtains the color display digits setting (precision) for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current number of display digits (Int32).
        """
        result = self.sdk.GetColorDispDigit(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetColorDispDigit.")
            
        ret, digit_number = result
        self._check_error(ret, "GetColorDispDigit")
        
        return int(digit_number)
    
    # ==================== DISPLAY TYPE ====================

    def set_display_type(self, display_type: str, com_port: int = 0):
        """
        Specifies the type of value displayed on the instrument: Absolute, Difference, or Ratio.
        
        Args:
            display_type: 'absolute', 'difference', or 'ratio'.
            com_port: Virtual COM Port number (0 = default).
        """
        DispType = self.types['DispType']
        
        type_map = {
            'absolute': DispType.Abs,
            'difference': DispType.Diff,
            'ratio': DispType.Ratio
        }
        
        disp_enum = type_map.get(display_type.lower())
        
        if not disp_enum:
            raise ValueError(f"Invalid display_type: must be one of {list(type_map.keys())}")
            
        ret = self.sdk.SetDisplayType(disp_enum, com_port)
        self._check_error(ret, "SetDisplayType")

    def get_display_type(self, com_port: int = 0) -> str:
        """
        Obtains the current display type setting for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current display type ('absolute', 'difference', or 'ratio').
        """
        result = self.sdk.GetDisplayType(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetDisplayType.")
            
        ret, disp_type_enum = result
        self._check_error(ret, "GetDisplayType")
        
        DispType = self.types['DispType']
        
        if disp_type_enum == DispType.Abs:
            return 'absolute'
        elif disp_type_enum == DispType.Diff:
            return 'difference'
        elif disp_type_enum == DispType.Ratio:
            return 'ratio'
        else:
            return 'unknown'
        
    # ==================== LANGUAGE DATE TIME ====================

    def set_display_language(self, language: str, com_port: int = 0):
        """
        Specifies the language setting for the instrument's display.
        
        Args:
            language: 'english', 'japanese', or 'chinese'.
            com_port: Virtual COM Port number (0 = default).
        """
        DisplayLanguage = self.types['DisplayLanguage']
        
        lang_map = {
            'english': DisplayLanguage.ENG,
            'japanese': DisplayLanguage.JPN,
            'chinese': DisplayLanguage.CHI
        }
        
        lang_enum = lang_map.get(language.lower())
        
        if not lang_enum:
            raise ValueError(f"Invalid language: must be one of {list(lang_map.keys())}")
            
        ret = self.sdk.SetDisplayLanguage(lang_enum, com_port)
        self._check_error(ret, "SetDisplayLanguage")

    def get_display_language(self, com_port: int = 0) -> str:
        """
        Obtains the current display language setting for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current language ('english', 'japanese', or 'chinese').
        """
        result = self.sdk.GetDisplayLanguage(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetDisplayLanguage.")
            
        ret, lang_enum = result
        self._check_error(ret, "GetDisplayLanguage")
        
        DisplayLanguage = self.types['DisplayLanguage']
        
        if lang_enum == DisplayLanguage.ENG:
            return 'english'
        elif lang_enum == DisplayLanguage.JPN:
            return 'japanese'
        elif lang_enum == DisplayLanguage.CHI:
            return 'chinese'
        else:
            return 'unknown'
        
    def set_datetime(self, dt: datetime, com_port: int = 0):
        """
        Specifies the date/time setting for the instrument.
        
        Args:
            dt: A Python datetime.datetime object.
            com_port: Virtual COM Port number (0 = default).
        """
        net_datetime = NetDateTime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)
            
        ret = self.sdk.SetDateTime(net_datetime, com_port)
        self._check_error(ret, "SetDateTime")

    def get_datetime(self, com_port: int = 0) -> datetime:
        """
        Obtains the date/time setting from the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            A Python datetime.datetime object.
        """
        result = self.sdk.GetDateTime(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetDateTime.")
            
        ret, net_datetime = result
        self._check_error(ret, "GetDateTime")
        
        return datetime(
            net_datetime.Year, 
            net_datetime.Month, 
            net_datetime.Day, 
            net_datetime.Hour, 
            net_datetime.Minute, 
            net_datetime.Second
        )
    
    def set_date_format(self, date_format: str, com_port: int = 0):
        """
        Specifies the date format settings for the instrument's display.
        
        Args:
            date_format: 'ymd', 'mdy', or 'dmy'.
            com_port: Virtual COM Port number (0 = default).
        """
        DateFormat = self.types['DateFormat']
        
        format_map = {
            'ymd': DateFormat.YYMMDD,
            'mdy': DateFormat.MMDDYY,
            'dmy': DateFormat.DDMMYY
        }
        
        format_enum = format_map.get(date_format.lower())
        
        if not format_enum:
            raise ValueError(f"Invalid date_format: must be one of {list(format_map.keys())}")
            
        ret = self.sdk.SetDateFormat(format_enum, com_port)
        self._check_error(ret, "SetDateFormat")

    def get_date_format(self, com_port: int = 0) -> str:
        """
        Obtains the current date format setting for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current date format ('ymd', 'mdy', or 'dmy').
        """
        result = self.sdk.GetDateFormat(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetDateFormat.")
            
        ret, format_enum = result
        self._check_error(ret, "GetDateFormat")
        
        DateFormat = self.types['DateFormat']
        
        if format_enum == DateFormat.YYMMDD:
            return 'ymd'
        elif format_enum == DateFormat.MMDDYY:
            return 'mdy'
        elif format_enum == DateFormat.DDMMYY:
            return 'dmy'
        else:
            return 'unknown'
        
    # ==================== DEVICE COLOR MODE ====================

    def set_color_mode(self, mode: str, com_port: int = 0):
        """
        Set color mode displayed on device.
        
        Args:
            mode: 'lvxy', 'lvudvd', 'lvtcpduv', 'xyz', 'lvdwpe', or 'lv'
            com_port: COM port number
        """
        mode_map = {
            'lvxy': self.types['ColorMode'].Lvxy,
            'lvudvd': self.types['ColorMode'].Lvudvd,
            'lvtcpduv': self.types['ColorMode'].LvTcpDuv,
            'xyz': self.types['ColorMode'].XYZ,
            'lvdwpe': self.types['ColorMode'].LvDwPe,
            'lv': self.types['ColorMode'].Lv
        }
        
        if mode.lower() not in mode_map:
            raise ValueError(f"Invalid mode. Must be one of: {list(mode_map.keys())}")
        
        ret = self.sdk.SetColorMode(mode_map[mode.lower()])
        self._check_error(ret, "SetColorMode")

    def get_color_mode(self, com_port: int = 0) -> str:
        """Get current color mode"""
        result = self.sdk.GetColorMode()
        if isinstance(result, tuple):
            ret, mode = result
            self._check_error(ret, "GetColorMode")
            
            mode_map = {
                self.types['ColorMode'].Lvxy: 'lvxy',
                self.types['ColorMode'].Lvudvd: 'lvudvd',
                self.types['ColorMode'].LvTcpDuv: 'lvtcpduv',
                self.types['ColorMode'].XYZ: 'xyz',
                self.types['ColorMode'].LvDwPe: 'lvdwpe',
                self.types['ColorMode'].Lv: 'lv'
            }
            return mode_map.get(mode, 'unknown')
        raise RuntimeError("Unexpected return from GetColorMode")
    
    def set_color_mode_display(self, mode: str, enable: bool, com_port: int = 0):
        """
        Turns a specific color mode on or off, making it available (or unavailable) 
        for display selection on the instrument.
        
        Args:
            mode: The color mode to enable/disable ('lvxy', 'lvudvd', etc.).
            enable: True to turn display ON, False to turn display OFF.
            com_port: Virtual COM Port number (0 = default).
        """
        ColorMode = self.types['ColorMode']
        ColorModeDisplay = self.types['ColorModeDisplay']

        mode_map = {
            'lvxy': ColorMode.Lvxy,
            'lvudvd': ColorMode.Lvudvd,
            'lvtcpduv': ColorMode.LvTcpDuv,
            'xyz': ColorMode.XYZ,
            'lvdwpe': ColorMode.LvDwPe,
            'lv': ColorMode.Lv
        }
        
        mode_enum = mode_map.get(mode.lower())
        if not mode_enum:
            raise ValueError(f"Invalid mode. Must be one of: {list(mode_map.keys())}")
            
        on_off_enum = ColorModeDisplay.On if enable else ColorModeDisplay.Off
            
        ret = self.sdk.SetColorModeDisplayOnOff(mode_enum, on_off_enum, com_port)
        self._check_error(ret, f"SetColorModeDisplayOnOff ({mode.upper()})")

    def get_color_mode_display(self, mode: str, com_port: int = 0) -> bool:
        """
        Obtains the display on/off setting for a specific color mode.
        
        Args:
            mode: The color mode to check ('lvxy', 'lvudvd', etc.).
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            True if the color mode is set to display, False otherwise.
        """
        ColorMode = self.types['ColorMode']
        ColorModeDisplay = self.types['ColorModeDisplay']

        mode_map = {
            'lvxy': ColorMode.Lvxy,
            'lvudvd': ColorMode.Lvudvd,
            'lvtcpduv': ColorMode.LvTcpDuv,
            'xyz': ColorMode.XYZ,
            'lvdwpe': ColorMode.LvDwPe,
            'lv': ColorMode.Lv
        }
        
        mode_enum = mode_map.get(mode.lower())
        if not mode_enum:
            raise ValueError(f"Invalid mode. Must be one of: {list(mode_map.keys())}")

        result = self.sdk.GetColorModeDisplayOnOff(mode_enum, comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetColorModeDisplayOnOff.")
            
        ret, on_off_enum = result
        self._check_error(ret, "GetColorModeDisplayOnOff")
        
        return on_off_enum == ColorModeDisplay.On
    
    # ==================== DATA SAVE MODE ====================

    def set_data_save_mode(self, mode: str, com_port: int = 0):
        """
        Specifies the data save mode setting for the instrument: Auto or Manual.
        
        Args:
            mode: 'auto' or 'manual'.
            com_port: Virtual COM Port number (0 = default).
        """
        DataSaveMode = self.types['DataSaveMode']
        
        mode_map = {
            'auto': DataSaveMode.AutoSave,
            'manual': DataSaveMode.ManuSave
        }
        
        mode_enum = mode_map.get(mode.lower())
        
        if not mode_enum:
            raise ValueError(f"Invalid mode. Must be one of: {list(mode_map.keys())}")
            
        ret = self.sdk.SetDataSaveMode(mode_enum, com_port)
        self._check_error(ret, "SetDataSaveMode")

    def get_data_save_mode(self, com_port: int = 0) -> str:
        """
        Obtains the current data save mode setting for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current data save mode ('auto' or 'manual').
        """
        result = self.sdk.GetDataSaveMode(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetDataSaveMode.")
            
        ret, mode_enum = result
        self._check_error(ret, "GetDataSaveMode")
        
        DataSaveMode = self.types['DataSaveMode']
        
        if mode_enum == DataSaveMode.AutoSave:
            return 'auto'
        elif mode_enum == DataSaveMode.ManuSave:
            return 'manual'
        else:
            return 'unknown'
        
    # ==================== PERIODIC CALIBRATION SETTING (BE REALLY CAREFUL WITH THAT ONE) ====================

    def set_periodic_calibration_notify(self, enable: bool, com_port: int = 0):
        """
        Turns the periodic calibration alert message on or off.
        
        Args:
            enable: True to turn the alert ON, False to turn it OFF.
            com_port: Virtual COM Port number (0 = default).
        """
        PeriodicCalibNotify = self.types['PeriodicCalibNotify']
        
        mode = PeriodicCalibNotify.On if enable else PeriodicCalibNotify.Off
            
        ret = self.sdk.SetPeriodicCalibNotify(mode, com_port)
        self._check_error(ret, "SetPeriodicCalibNotify")

    def get_periodic_calibration_notify(self, com_port: int = 0) -> bool:
        """
        Obtains the periodic calibration alert on/off setting for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            True if the periodic calibration alert is ON, False if it is OFF.
        """
        result = self.sdk.GetPeriodicCalibNotify(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetPeriodicCalibNotify.")
            
        ret, mode_enum = result
        self._check_error(ret, "GetPeriodicCalibNotify")
        
        return mode_enum == self.types['PeriodicCalibNotify'].On
    
    # ==================== BUTTON SETTINGS ====================

    def set_toggle(self, enable: bool, com_port: int = 0):
        """
        Specifies the toggle on/off setting for the instrument's measurement button.
        
        - If ON (True), press to start, press again to stop (Toggle mode).
        - If OFF (False), press and hold to measure, release to stop (Standard mode).
        
        Args:
            enable: True for Toggle mode, False for Standard mode.
            com_port: Virtual COM Port number (0 = default).
        """
        ToggleStatus = self.types['ToggleStatus']
        
        mode = ToggleStatus.On if enable else ToggleStatus.Off
            
        ret = self.sdk.SetToggleOnOff(mode, com_port)
        self._check_error(ret, "SetToggleOnOff")

    def get_toggle(self, com_port: int = 0) -> bool:
        """
        Obtains the toggle on/off setting for the instrument's measurement button.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            True if Toggle mode is ON, False if it is OFF (Standard mode).
        """
        result = self.sdk.GetToggleOnOff(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetToggleOnOff.")
            
        ret, mode_enum = result
        self._check_error(ret, "GetToggleOnOff")
        
        return mode_enum == self.types['ToggleStatus'].On
    
    def set_trigger(self, enable: bool, com_port: int = 0):
        """
        Specifies the enabled/disabled setting for the instrument's measurement button (trigger).
        
        Note: The trigger is disabled automatically after disconnection.
        
        Args:
            enable: True to enable the measurement button, False to disable it.
            com_port: Virtual COM Port number (0 = default).
        """
        TriggerStatus = self.types['TriggerStatus']
        
        mode = TriggerStatus.On if enable else TriggerStatus.Off
            
        ret = self.sdk.SetTriggerOnOff(mode, com_port)
        self._check_error(ret, "SetTriggerOnOff")

    def get_trigger(self, com_port: int = 0) -> bool:
        """
        Obtains the enabled/disabled status for the instrument's measurement button (trigger).
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            True if the button is enabled, False if it is disabled.
        """
        # Note: Use explicit keyword argument for robustness with pythonnet
        result = self.sdk.GetTriggerOnOff(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetTriggerOnOff.")
            
        ret, mode_enum = result
        self._check_error(ret, "GetTriggerOnOff")
        
        return mode_enum == self.types['TriggerStatus'].On

    # ==================== DEVICE SDK INFO ====================
    
    def get_device_info(self, com_port: int = 0) -> Dict[str, any]:
        """
        Get device information.
        
        Returns:
            Dictionary with device info (product name, serial, firmware version, etc.)
        """
        result = self.sdk.GetDeviceInfo()
        if isinstance(result, tuple):
            ret, info = result
            self._check_error(ret, "GetDeviceInfo")
            
            return {
                'product_name': str(info.ProductName),
                'serial_number': str(info.SerialNumber),
                'firmware_major': str(info.SoftMajorVersion),
                'firmware_minor': str(info.SoftMinorVersion),
                'firmware_free': str(info.SoftFreeVersion),
                'calibration_expiration': info.PeriodicCalibrationExpirationDate,
                'calibration_warning': bool(info.PeriodicCalibrationWarningStatus)
            }
        raise RuntimeError("Unexpected return from GetDeviceInfo")

    def get_sdk_version(self) -> str:
        """Get SDK version string"""
        result = self.sdk.GetSDKVersion()
        if isinstance(result, tuple):
            ret, version = result
            self._check_error(ret, "GetSDKVersion")
            return str(version)
        raise RuntimeError("Unexpected return from GetSDKVersion")
    
    # ==================== MEASUREMENT UNIT SETTINGS ====================

    def set_luminance_unit(self, unit: str, com_port: int = 0):
        """
        Specifies the luminance unit setting for the instrument.
        
        Args:
            unit: 'cdm2' (candelas/m^2) or 'fl' (foot-Lamberts).
            com_port: Virtual COM Port number (0 = default).
        """
        LuminanceUnit = self.types['Unit']
        
        unit_map = {
            'cdm2': LuminanceUnit.cdm2,
            'fl': LuminanceUnit.other
        }
        
        unit_enum = unit_map.get(unit.lower())
        
        if not unit_enum:
            raise ValueError(f"Invalid unit. Must be one of: {list(unit_map.keys())}")
            
        ret = self.sdk.SetLuminanceUnit(unit_enum, com_port)
        self._check_error(ret, "SetLuminanceUnit")

    def get_luminance_unit(self, com_port: int = 0) -> str:
        """
        Obtains the luminance unit setting for the instrument.
        
        Args:
            com_port: Virtual COM Port number (0 = default).
            
        Returns:
            The current luminance unit ('cdm2' or 'fl').
        """
        result = self.sdk.GetLuminanceUnit(comPort=com_port)
        
        if not isinstance(result, tuple) or len(result) != 2:
            raise RuntimeError("Unexpected return format from GetLuminanceUnit.")
            
        ret, unit_enum = result
        self._check_error(ret, "GetLuminanceUnit")
        
        LuminanceUnit = self.types['Unit']
        
        if unit_enum == LuminanceUnit.cdm2:
            return 'cdm2'
        elif unit_enum == LuminanceUnit.other:
            return 'fl'
        else:
            return 'unknown'

    # ==================== CONTEXT MANAGER ====================
    
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
        return False

    def __repr__(self):
        return f"KonicaDevice(sdk_path='{self._sdk_bin_path}')"