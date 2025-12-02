# Konica Minolta CS/LS Device Control Library

Python library for controlling Konica Minolta CS-150, CS-160, LS-150, and LS-160 chroma meters and luminance meters. All functions still under development.

## Features

- Complete device control via USB
- Luminance and color measurements
- All measurement modes (Lvxy, XYZ, Lvudvd, etc.)
- Device configuration (integration time, sync, peak/valley)
- Display and backlight control
- Close-up lens configuration
- Read/delete stored measurements
- Device information and diagnostics

## Requirements

- Python 3.6+
- pythonnet >= 3.0.0
- Windows 10/11 (might work on different OS, not tested yet)
- Konica Minolta LC-MISDK (included in `sdk_bin` folder)
- Device USB drivers (see candelameter manual and section below)

## Installation

### 1. Driver

The device needs to be installed with the correct drivers. If it shows up in the device manager as just `USB Serial Device` it might not work. These steps did work for me:

1. Open the device manager
2. Right click the device under (COM & LPT)
3. Click update drivers
4. Click `Browse my computer for driver software`
5. Select the driver folder and `Have Disk...`
6. Choose the .INF file"

The candelameter should now show up as somethign like `Measurement Device (COMx)`

### 2. Install pythonnet

```bash
pip install pythonnet
```

### 3. Set up library structure

For this project to work properly you need to replicate the project structure from a .net application. The easiest way is to build a minimal working .net example (in the Konica Minolta LC-MISDK there is some example code, additional dlls need to be installed via nuget) and then copy the debug folder with all the .dll files. For copyright reasons I can't include any files from the LC-MISDK in this project. In the future I will test in more detail which files are actually neccessary.

This is a tree view of my working solution :

```
ROOT FOLDER
│   libtest.py
│
└───src
    └───konica_lscs
        │   __init__.py
        │
        └───sdk_bin
            │   App.config
            │   CalculateColor.dll
            │   CalculateDominantWavelength.dll
            │   CalculateRequiredData.dll
            │   CalculateTduvJIS.dll
            │   CalculateXYZ.dll
            │   CalculateYuvD.dll
            │   CalculateYxy.dll
            │   Kmop.BusinessCore.dll
            │   Kmop.BusinessProxy.dll
            │   Kmop.BusinessWorkflow.dll
            │   Kmop.ColorCore.dll
            │   Kmop.ColorProxy.dll
            │   Kmop.CommunicationCore.dll
            │   Kmop.CommunicationProxy.dll
            │   Kmop.Constants.dll
            │   Kmop.ContextServices.dll
            │   Kmop.DAL.Utilities.dll
            │   Kmop.DataAccessCore.dll
            │   Kmop.DataAccessProxy.dll
            │   Kmop.DummyDAL.dll
            │   Kmop.HostConn.dll
            │   Kmop.IBusiness.dll
            │   Kmop.IBusinessServicesContract.dll
            │   Kmop.IColorService.dll
            │   Kmop.ICommunicationServicesContract.dll
            │   Kmop.IDAL.dll
            │   Kmop.IDataAccessServiceContract.dll
            │   Kmop.IDeviceCommand.dll
            │   Kmop.IInstrumentServiceContract.dll
            │   Kmop.IKmopDAL.dll
            │   Kmop.Instrument.ConnectDevice.dll
            │   Kmop.Instrument.CS100P.dll
            │   Kmop.InstrumentControl.dll
            │   Kmop.InstrumentCore.dll
            │   Kmop.InstrumentProxy.dll
            │   Kmop.IRuleEngine.dll
            │   Kmop.KmopDummyDAL.dll
            │   Kmop.Rest.DTO.dll
            │   Kmop.Rest.Interface.dll
            │   Kmop.RuleEngine.CS100P.dll
            │   Kmop.RuleEngineProxy.dll
            │   Kmop.SignalRClient.dll
            │   Kmop.Utilities.dll
            │   LC-MISDK.dll
            │   log4net.dll
            │   MathNet.Numerics.dll
            │   Microsoft.AspNet.SignalR.Client.dll
            │   Newtonsoft.Json.dll
            │   ServiceStack.Common.dll
            │   ServiceStack.dll
            │   ServiceStack.Interfaces.dll
            │   ServiceStack.OrmLite.dll
            │   ServiceStack.OrmLite.MySql.dll
            │   ServiceStack.ServiceInterface.dll
            │   ServiceStack.Text.dll
            │   SimpleInjector.Diagnostics.dll
            │   SimpleInjector.dll
            │   WeightValue.dll
            │   WordingList.dll
            │
            ├───calccolor
            │   └───x64
            │           CalculateColor.dll
            │           CalculateDominantWavelength.dll
            │           CalculateRequiredData.dll
            │           CalculateTduvJIS.dll
            │           CalculateXYZ.dll
            │           CalculateYuvD.dll
            │           CalculateYxy.dll
            │           OBS_10_1nm.obs
            │           OBS_10_5nm.obs
            │           OBS_2_1nm.obs
            │           OBS_2_5nm.obs
            │           WeightValue.dll
            │           WordingList.dll
            │           WP_USER01.uwp
            │
            └───IDMap
                    CS-150.xml
                    CS-160.xml
                    LS-150.xml
                    LS-160.xml
```

## Quick Start

Included in the project there is a `libtest.py` file that runs through some basic examples. Below are some very basic use-cases.

```python
from konica_lscs import KonicaDevice

# Basic measurement
with KonicaDevice() as device:
    device.connect()
    device.measure()
    
    luminance = device.get_luminance()
    print(f"Luminance: {luminance} cd/m²")
```

## Usage Examples

### 1. Basic Measurement

```python
from konica_lscs import KonicaDevice

device = KonicaDevice()
device.connect()

# Take a measurement
device.measure()

# Get luminance
lv = device.get_luminance()
print(f"Luminance: {lv:.2f} cd/m²")

# Get all color values
color = device.get_color()
print(f"Color: {color}")

device.disconnect()
```

### 2. Using a Context Manager (Recommended)

```python
with KonicaDevice() as device:
    device.connect()
    device.measure()
    lv = device.get_luminance()
    print(f"Luminance: {lv:.2f} cd/m²")
# Automatically disconnects device
```

### 3. Device Information

```python
with KonicaDevice() as device:
    device.connect()
    
    info = device.get_device_info()
    print(f"Device: {info['product_name']}")
    print(f"Serial: {info['serial_number']}")
    print(f"Firmware: {info['firmware_major']}.{info['firmware_minor']}")
    
    sdk_version = device.get_sdk_version()
    print(f"SDK: {sdk_version}")
```

### 4. Measurement Settings

```python
with KonicaDevice() as device:
    device.connect()
    
    # Set auto integration time
    device.set_measurement_time('auto')
    
    # Or manual integration time (1.5 seconds)
    device.set_measurement_time('manual', manual_time=1.5)
    
    # Enable sync mode at 60 Hz
    device.set_sync_mode(sync=True, frequency=60.0)
    
    # Set peak measurement
    device.set_peak_valley('peak')  # 'peak', 'valley', or 'off'
```

### 5. Display Control

```python
with KonicaDevice() as device:
    device.connect()
    
    # Backlight control
    device.set_backlight(True)
    device.set_backlight_level(5)  # 1-5 (1=darkest, 5=brightest)
    
    # Color mode
    device.set_color_mode('lvxy')  # 'lvxy', 'xyz', 'lvudvd', etc.
```

### 6. Continuous Monitoring

```python
with KonicaDevice() as device:
    device.connect()
    device.set_measurement_time('auto')
    
    for i in range(10):
        device.measure()
        lv = device.get_luminance()
        print(f"Measurement {i+1}: {lv:.2f} cd/m²")
```

### 7. Working with Stored Data

```python
with KonicaDevice() as device:
    device.connect()
    
    # Get number of stored measurements
    count = device.get_number_of_samples()
    print(f"Stored samples: {count}")
    
    # Read stored measurements
    for i in range(1, count + 1):
        sample = device.read_sample_data(i)
        print(f"Sample {i}: {sample}")
    
    # Delete all stored data
    device.delete_sample_data(-1)  # -1 = delete all
```

### 8. Multiple Color Spaces

```python
with KonicaDevice() as device:
    device.connect()
    device.measure()
    
    # Read as XYZ
    device.set_color_mode('xyz')
    xyz = device.get_color()
    print(f"XYZ: {xyz}")
    
    # Read as Lvxy
    device.set_color_mode('lvxy')
    lvxy = device.get_color()
    print(f"Lvxy: {lvxy}")
    
    # Or read XYZ directly
    xyz_direct = device.read_latest_data_xyz()
    print(f"XYZ direct: {xyz_direct}")
```

## Functions Reference

There are all the functions from the official manual included in the library. However so far I have only tested the ones below:

### Connection Methods

- `connect(com_port=0)` - Connect to device (0 = auto-detect)
- `disconnect(com_port=0)` - Disconnect from device
- `get_device_list()` - Get dictionary of connected devices

### Measurement Methods

- `measure(wait=True, com_port=0)` - Take a measurement
- `wait_for_idle(com_port=0)` - Wait for measurement to complete
- `cancel_measurement(com_port=0)` - Cancel ongoing measurement

### Data Reading Methods

- `get_luminance(com_port=0)` - Get luminance in cd/m²
- `get_color(com_port=0)` - Get all color values as dict
- `read_display_value(com_port=0)` - Read device display values
- `read_latest_data_xyz(com_port=0)` - Read as XYZ tristimulus values

### Stored Data Methods

- `get_number_of_samples(com_port=0)` - Get count of stored measurements
- `read_sample_data(data_number, com_port=0)` - Read stored measurement
- `delete_sample_data(data_number=-1, com_port=0)` - Delete stored data (-1 = all)

### Measurement Settings

- `set_measurement_time(mode, manual_time=1.0, com_port=0)` - Set integration time
- `get_measurement_time(com_port=0)` - Get integration time settings
- `set_sync_mode(sync, frequency=60.0, com_port=0)` - Set sync mode
- `get_sync_mode(com_port=0)` - Get sync mode settings
- `set_peak_valley(mode, com_port=0)` - Set peak/valley mode ('off', 'peak', 'valley')
- `get_peak_valley(com_port=0)` - Get peak/valley mode

### Lens Settings

- `set_close_up_lens(lens_type, com_port=0)` - Set close-up lens ('standard', 'no153', 'no135', 'no122', 'no110')
- `get_close_up_lens(com_port=0)` - Get current lens setting

### Display Settings

- `set_backlight(on, com_port=0)` - Turn backlight on/off
- `get_backlight(com_port=0)` - Get backlight status
- `set_backlight_level(level, com_port=0)` - Set brightness (1-5)
- `get_backlight_level(com_port=0)` - Get brightness level
- `set_color_mode(mode, com_port=0)` - Set color mode ('lvxy', 'xyz', 'lvudvd', 'lvtcpduv', 'lvdwpe', 'lv')
- `get_color_mode(com_port=0)` - Get current color mode

### Device Information

- `get_device_info(com_port=0)` - Get device info dict
- `get_sdk_version()` - Get SDK version string

## Color Modes

The device supports multiple color spaces:

- **Lvxy** - Luminance (cd/m²), chromaticity x, y
- **XYZ** - CIE XYZ tristimulus values
- **Lvudvd** - Luminance, CIE 1976 UCS chromaticity u', v'
- **LvTcpDuv** - Luminance, correlated color temperature, Duv
- **LvDwPe** - Luminance, dominant wavelength, excitation purity
- **Lv** - Luminance only (LS-series devices)


## Supported Devices

- CS-150 (Chroma Meter)
- CS-160 (Chroma Meter)
- LS-150 (Luminance Meter)
- LS-160 (Luminance Meter)

The LS-series devices only support luminance measurements, not color.

## Troubleshooting

### "SDK not initialized" error
- Ensure all DLL files are in the `sdk_bin` folder
- Check that pythonnet is properly installed

### "Connect Failed" error
- Have you tried turning it off and on again? Is it plugged in?
- Ensure device drivers are installed (see manual section 5.3)
- Try specifying COM port explicitly: `device.connect(com_port=3)`

### "Read Data Failed" error
- Ensure measurement completed before reading: `device.measure(wait=True)`
- Check that color mode is supported by your device model

## Advanced Usage

### Multiple Devices

```python
# Get list of connected devices
with KonicaDevice() as device:
    devices = device.get_device_list()
    print(f"Found devices: {devices}")
    
    # Connect to specific device
    device.connect(com_port=3)
```

### Custom Integration Time

```python
with KonicaDevice() as device:
    device.connect()
    
    # Fast measurements (auto mode)
    device.set_measurement_time('auto')
    
    # Or precise measurements (manual mode)
    device.set_measurement_time('manual', manual_time=2.0)
```

## License

This library interfaces with the Konica Minolta LC-MISDK. See the SDK license agreement for terms. The python code is licensed under the MIT License.


## Contributing

Contributions are very welcome. Report any issues via Github. If anybody would like to further help improve this project I'm happy about any help.

## Known Issues

- `read_latest_data` function not working properly. For now use `read_display_value`

## ToDo

- Check if all functions work
- Check which files are actually neccessary, at the moment all are included
- Re-work into actually installable library
- Implement proper tests

## Version History

- **0.1.0** - Initial release