"""
Comprehensive examples for using the konica_lscs library.

This script demonstrates all major features of the Konica Minolta
CS/LS series device control library.
"""

import sys
import os
import time

# Setup path
current_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(current_dir, "src")
sys.path.insert(0, src_path)

from konica_lscs import KonicaDevice


def example_basic_measurement():
    """Example 1: Basic measurement"""
    print("\n" + "="*60)
    print("EXAMPLE 1: Basic Measurement")
    print("="*60)
    
    device = KonicaDevice()
    device.connect()
    
    print("Taking measurement...")
    device.measure()
    
    lv = device.get_luminance()
    print(f"✓ Luminance: {lv:.2f} cd/m²")
    
    color = device.get_color()
    print(f"✓ Color values: {color}")
    
    device.disconnect()


def example_context_manager():
    """Example 2: Using context manager"""
    print("\n" + "="*60)
    print("EXAMPLE 2: Context Manager (Recommended)")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        device.measure()
        
        lv = device.get_luminance()
        print(f"✓ Luminance: {lv:.2f} cd/m²")


def example_device_info():
    """Example 3: Get device information"""
    print("\n" + "="*60)
    print("EXAMPLE 3: Device Information")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        # Get device info
        info = device.get_device_info()
        print(f"Device: {info['product_name']}")
        print(f"Serial: {info['serial_number']}")
        print(f"Firmware: {info['firmware_major']}.{info['firmware_minor']}.{info['firmware_free']}")
        print(f"Calibration due: {info['calibration_expiration']}")
        print(f"Calibration warning: {info['calibration_warning']}")
        
        # Get SDK version
        sdk_ver = device.get_sdk_version()
        print(f"SDK Version: {sdk_ver}")
        
        # Get device list
        devices = device.get_device_list()
        print(f"\nAvailable devices: {devices}")


def example_measurement_settings():
    """Example 4: Configure measurement settings"""
    print("\n" + "="*60)
    print("EXAMPLE 4: Measurement Settings")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        # Set auto measurement time
        print("Setting auto measurement time...")
        device.set_measurement_time('auto')
        mode, time_val = device.get_measurement_time()
        print(f"✓ Measurement time: {mode}")
        
        # Set manual measurement time
        print("\nSetting manual measurement time (1.5s)...")
        device.set_measurement_time('manual', manual_time=1.5)
        mode, time_val = device.get_measurement_time()
        print(f"✓ Measurement time: {mode}, {time_val}s")
        
        # Set sync mode
        print("\nSetting sync mode (50 Hz)...")
        device.set_sync_mode(sync=True, frequency=50.0)
        sync, freq = device.get_sync_mode()
        print(f"✓ Sync: {sync}, Frequency: {freq} Hz")
        
        # Set peak/valley
        print("\nSetting peak measurement...")
        device.set_peak_valley('peak')
        pv = device.get_peak_valley()
        print(f"✓ Peak/Valley: {pv}")


def example_display_settings():
    """Example 5: Display and backlight settings"""
    print("\n" + "="*60)
    print("EXAMPLE 5: Display Settings")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        # Backlight control
        print("Turning backlight ON...")
        device.set_backlight(True)
        time.sleep(1)
        
        print("Setting backlight level to 5 (brightest)...")
        device.set_backlight_level(5)
        level = device.get_backlight_level()
        print(f"✓ Backlight level: {level}")
        time.sleep(1)
        
        print("Turning backlight OFF...")
        device.set_backlight(False)
        
        # Color mode
        print("\nSetting color mode to Lvxy...")
        device.set_color_mode('lvxy')
        mode = device.get_color_mode()
        print(f"✓ Color mode: {mode}")


def example_lens_settings():
    """Example 6: Close-up lens settings"""
    print("\n" + "="*60)
    print("EXAMPLE 6: Lens Settings")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        # Get current lens
        lens = device.get_close_up_lens()
        print(f"Current lens: {lens}")
        
        # Set to standard (no close-up lens)
        print("Setting to standard lens...")
        device.set_close_up_lens('standard')
        lens = device.get_close_up_lens()
        print(f"✓ Lens: {lens}")


def example_stored_data():
    """Example 7: Working with stored measurements"""
    print("\n" + "="*60)
    print("EXAMPLE 7: Stored Measurement Data")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        # Check how many samples are stored
        count = device.get_number_of_samples()
        print(f"Stored samples: {count}")
        
        if count > 0:
            # Read first stored sample
            print(f"\nReading sample 1...")
            try:
                sample = device.read_sample_data(1) # Samples seem to be srating from 1 not 0
                print(f"Sample 1: {sample}")
                
                # Read all samples
                print(f"\nReading all {count} samples:")
                for i in range(1, min(count + 1, 6)):  # Show first 5
                    try:
                        sample = device.read_sample_data(i)
                        if 'Lv' in sample:
                            print(f"  Sample {i}: Lv={sample['Lv']:.2f} cd/m²")
                        elif 'Y' in sample:
                            print(f"  Sample {i}: Y={sample['Y']:.2f}")
                        else:
                            print(f"  Sample {i}: {sample}")
                    except RuntimeError as e:
                        print(f"  Sample {i}: Could not read ({e})")
            except RuntimeError as e:
                print(f"Could not read stored data: {e}")
        else:
            print("No stored samples. Take a measurement on the device to store data.")
        
        # Uncomment to delete stored data
        #device.delete_sample_data(-1)  # -1 = delete all
        #print("✓ Deleted all stored samples")


def example_continuous_monitoring():
    """Example 8: Continuous measurement monitoring"""
    print("\n" + "="*60)
    print("EXAMPLE 8: Continuous Monitoring (10 measurements)")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        # Set to auto measurement time for faster readings
        device.set_measurement_time('auto')
        
        print("\nTaking 10 measurements...")
        print("#  |  Time (s) |  Δt (s) | Luminance (cd/m²)")
        print("-" * 55)
        
        start_time = time.time()
        last_time = start_time
        
        for i in range(10):
            device.measure()
            lv = device.get_luminance()
            
            now = time.time()
            elapsed = now - start_time
            delta_t = now - last_time
            last_time = now
            
            print(f"{i+1:2} | {elapsed:9.2f} | {delta_t:7.2f} | {lv:10.2f}")
            
            time.sleep(0.5)  # Wait between measurements


def example_multiple_color_modes():
    """Example 9: Reading different color spaces"""
    print("\n" + "="*60)
    print("EXAMPLE 9: Multiple Color Spaces")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        # Take one measurement
        print("Taking measurement...")
        device.measure()
        
        # Read in different color modes
        modes = ['lvxy', 'xyz', 'lvudvd']
        
        for mode in modes:
            try:
                device.set_color_mode(mode)
                values = device.get_color()
                print(f"\n{mode.upper()}: {values}")
            except Exception as e:
                print(f"\n{mode.upper()}: Not available ({e})")


def example_xyz_direct():
    """Example 10: Direct XYZ reading"""
    print("\n" + "="*60)
    print("EXAMPLE 10: Direct XYZ Reading")
    print("="*60)
    
    with KonicaDevice() as device:
        device.connect()
        
        print("Taking measurement...")
        device.measure()
        
        # Read as XYZ directly
        xyz = device.read_latest_data_xyz()
        print(f"\nXYZ Tristimulus values:")
        print(f"  X = {xyz['X']:.4f}")
        print(f"  Y = {xyz['Y']:.4f}")
        print(f"  Z = {xyz['Z']:.4f}")
        
        # Y is luminance
        print(f"\nLuminance (Y) = {xyz['Y']:.2f} cd/m²")


def run_all_examples():
    """Run all examples"""
    print("\n" + "#"*60)
    print("# KONICA MINOLTA CS/LS DEVICE CONTROL EXAMPLES")
    print("#"*60)
    
    try:
        example_basic_measurement()
        example_context_manager()
        example_device_info()
        example_measurement_settings()
        example_display_settings()
        example_lens_settings()
        example_stored_data()
        example_continuous_monitoring()
        example_multiple_color_modes()
        example_xyz_direct()
        
        print("\n" + "="*60)
        print("✓ All examples completed successfully!")
        print("="*60)
        
    except Exception as e:
        print(f"\n✗ Example failed: {e}")
        import traceback
        traceback.print_exc()


def run_single_example(example_num: int):
    """Run a single example"""
    examples = {
        1: example_basic_measurement,
        2: example_context_manager,
        3: example_device_info,
        4: example_measurement_settings,
        5: example_display_settings,
        6: example_lens_settings,
        7: example_stored_data,
        8: example_continuous_monitoring,
        9: example_multiple_color_modes,
        10: example_xyz_direct
    }
    
    if example_num in examples:
        try:
            examples[example_num]()
            print("\n✓ Example completed!")
        except Exception as e:
            print(f"\n✗ Example failed: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"Example {example_num} not found. Available: 1-{len(examples)}")


def print_menu():
    """Print the example menu"""
    print("\n" + "="*60)
    print("KONICA MINOLTA CS/LS DEVICE CONTROL EXAMPLES")
    print("="*60)
    print("\nAvailable Examples:")
    print("  0  - Run all examples")
    print("  1  - Basic Measurement")
    print("  2  - Context Manager (Recommended)")
    print("  3  - Device Information")
    print("  4  - Measurement Settings")
    print("  5  - Display Settings")
    print("  6  - Lens Settings")
    print("  7  - Stored Measurement Data")
    print("  8  - Continuous Monitoring (10 measurements)")
    print("  9  - Multiple Color Spaces")
    print("  10 - Direct XYZ Reading")
    print("  q  - Quit")
    print("="*60)


if __name__ == "__main__":
    import sys
    
    # Check if command line argument provided (non-interactive mode)
    if len(sys.argv) > 1:
        try:
            choice = int(sys.argv[1])
            if choice == 0:
                run_all_examples()
            elif 1 <= choice <= 10:
                run_single_example(choice)
            else:
                print(f"Invalid choice: {choice}. Please choose 0-10.")
                sys.exit(1)
        except ValueError:
            print("Error: Please provide a number between 0 and 10")
            print("Usage: python examples.py [example_number]")
            sys.exit(1)
    else:
        # Interactive mode with loop
        while True:
            print_menu()
            
            try:
                user_input = input("\nEnter example number (0-10) or 'q' to quit: ").strip()
                
                # Check for quit
                if user_input.lower() in ['q', 'quit', 'exit']:
                    print("\nExiting... Goodbye!")
                    break
                
                # Try to parse as integer
                try:
                    choice = int(user_input)
                except ValueError:
                    print(f"\nInvalid input: '{user_input}'. Please enter a number (0-10) or 'q'.")
                    input("Press Enter to continue...")
                    continue
                
                # Run selected example(s)
                if choice == 0:
                    run_all_examples()
                elif 1 <= choice <= 10:
                    run_single_example(choice)
                else:
                    print(f"\nInvalid choice: {choice}. Please choose 0-10.")
                
                # Wait for user before showing menu again
                input("\nPress Enter to return to menu...")
                
            except KeyboardInterrupt:
                print("\n\nInterrupted by user. Exiting...")
                break
            except EOFError:
                print("\n\nEOF detected. Exiting...")
                break