import ctypes
import os
from ctypes import POINTER, byref, c_int, c_long, c_ushort, c_void_p, cast
from ctypes.wintypes import DWORD

from comtypes import BSTR

dll = ctypes.CDLL(os.path.join(os.getcwd(), "mlsdk64.dll"))


class SAFEARRAYBOUND(ctypes.Structure):
    _fields_ = [("cElements", c_long), ("lLbound", c_long)]


class SAFEARRAY(ctypes.Structure):
    _fields_ = [
        ("cDims", c_ushort),
        ("fFeatures", c_ushort),
        ("cbElements", c_long),
        ("cLocks", c_long),
        ("pvData", c_void_p),
        ("rgsabound", SAFEARRAYBOUND * 1),
    ]


# int MLAPI_GetErrorMessage(int ErrorCode, BSTR* pDesc)
def error_message(error_code):

    # Prints the error message corresponding to the given error code.

    # Parameters:
    # error_code (int): The error code for which the corresponding error message needs to be retrieved.

    # Returns:
    # None. The function prints the error message directly.

    dll.MLAPI_GetErrorMessage.argtypes = [c_int, POINTER(BSTR)]
    dll.MLAPI_GetErrorMessage.restype = c_int
    error_desc = BSTR()
    dll.MLAPI_GetErrorMessage(error_code, byref(error_desc))
    print(error_desc.value)


# int MLAPI_Initialize()
def initialize_dll():

    # Initializes the Mystic Light SDK DLL.

    # This function calls the MLAPI_Initialize function from the DLL to initialize the SDK.
    # It prints a success message if the initialization is successful, or calls the error_message
    # function to print the corresponding error message if the initialization fails.

    # Parameters:
    # None.

    # Returns:
    # None. The function prints the result directly.

    dll.MLAPI_Initialize.argtypes = []
    dll.MLAPI_Initialize.restype = c_int
    status = dll.MLAPI_Initialize()
    if status == 0:
        print("DLL Initialized successfully.")
    else:
        error_message(status)
        exit()


# int MLAPI_Release()
def release_dll():

    # Releases the Mystic Light SDK DLL.

    # This function calls the MLAPI_Release function from the DLL to release the SDK.
    # It prints a success message if the release is successful, or calls the error_message
    # function to print the corresponding error message if the release fails.

    # Parameters:
    # None.

    # Returns:
    # None. The function prints the result directly.

    dll.MLAPI_Release.argtypes = []
    dll.MLAPI_Release.restype = c_int
    status = dll.MLAPI_Release()
    if status == 0:
        print("DLL released successfully.")
    else:
        error_message(status)
        exit()


# int MLAPI_GetDeviceInfo(SAFEARRAY** pDevType, SAFEARRAY** pLedCount)
def get_device_info():

    # Retrieves information about connected devices and their LED counts.

    # This function calls the MLAPI_GetDeviceInfo function from the DLL to obtain
    # the types of devices and the number of LEDs each device has. It then prompts
    # the user to select a device from the list.

    # Returns:
    #     tuple: A tuple containing the selected device type (BSTR) and the number
    #            of LEDs (int) for the selected device. Returns None if the device
    #            selection is invalid or an error occurs.

    dll.MLAPI_GetDeviceInfo.argtypes = [
        POINTER(POINTER(SAFEARRAY)),
        POINTER(POINTER(SAFEARRAY)),
    ]
    dll.MLAPI_GetDeviceInfo.restype = c_int

    pDevType = POINTER(SAFEARRAY)()
    pLedCount = POINTER(SAFEARRAY)()
    status = dll.MLAPI_GetDeviceInfo(byref(pDevType), byref(pLedCount))
    if status != 0:
        error_message(status)
        exit()
    else:
        device_types = cast(pDevType.contents.pvData, POINTER(BSTR))
        led_counts = cast(pLedCount.contents.pvData, POINTER(BSTR))
        for i in range(pDevType.contents.rgsabound[0].cElements):
            print(f"{i}: Device Type: {device_types[i]}, LED Count: {led_counts[i]}")
        try:
            n = int(input("Choose device: "))
            return device_types[n], int(led_counts[n])
        except (ValueError, TypeError):
            print("The device is specified incorrectly.")
            return None


# int MLAPI_GetDeviceName(BSTR type, SAFEARRAY** pDevName)
def get_device_name(device_type):

    # Retrieves the name of a device based on its type.
    # (Strangely doesn`t work for me, I used next function to get the name.)

    # This function calls the MLAPI_GetDeviceName function from the DLL to obtain
    # the name of a device specified by the given device type.

    # Parameters:
    # device_type (BSTR): The type of the device for which the name is to be retrieved.

    # Returns:
    # None. The function prints the device name directly. If an error occurs, it prints
    # the error message and exits the program.

    dll.MLAPI_GetDeviceName.argtypes = [
        BSTR,
        POINTER(POINTER(SAFEARRAY)),
    ]
    dll.MLAPI_GetDeviceName.restype = c_int
    pDevName = POINTER(SAFEARRAY)()
    status = dll.MLAPI_GetDeviceName(device_type, byref(pDevName))
    if status != 0:
        error_message(status)
        exit()
    else:
        print(pDevName.contents.pvData)


# int MLAPI_GetDeviceNameEx(BSTR type, DWORD index, BSTR* pDevName)
def get_device_name_ex(device_type, index):

    # Retrieves the name of a specific device based on its type and index.

    # This function calls the MLAPI_GetDeviceNameEx function from the DLL to obtain
    # the name of a device specified by the given device type and index.

    # Parameters:
    # device_type (BSTR): The type of the device for which the name is to be retrieved.
    # index (DWORD): The index of the device within the specified type.

    # Returns:
    # str: The name of the device if successful. If an error occurs, the function
    # prints the error message and exits the program.

    dll.MLAPI_GetDeviceNameEx.argtypes = [
        BSTR,
        DWORD,
        POINTER(BSTR),
    ]
    dll.MLAPI_GetDeviceNameEx.restype = c_int
    pDevName = BSTR()
    status = dll.MLAPI_GetDeviceNameEx(device_type, index, byref(pDevName))
    if status != 0:
        error_message(status)
        exit()
    else:
        print(index, ":", pDevName.value)
        return pDevName.value


# int MLAPI_GetLedInfo(BSTR type, DWORD index, BSTR* pName, SAFEARRAY** pLedStyles)
def get_led_info(device_type, index):

    # Retrieves information about a specific LED on a device.

    # This function calls the MLAPI_GetLedInfo function from the DLL to obtain
    # the name and available styles of a specific LED identified by the given
    # device type and index.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.

    # Returns:
    # POINTER(SAFEARRAY): A pointer to a SAFEARRAY containing the available styles
    # for the specified LED. If an error occurs, the function prints the error
    # message and exits the program.

    dll.MLAPI_GetLedInfo.argtypes = [
        BSTR,
        DWORD,
        POINTER(BSTR),
        POINTER(POINTER(SAFEARRAY)),
    ]
    dll.MLAPI_GetLedInfo.restype = c_int
    pName = BSTR()
    pLedStyles = POINTER(SAFEARRAY)()
    status = dll.MLAPI_GetLedInfo(device_type, index, byref(pName), byref(pLedStyles))
    if status != 0:
        error_message(status)
        exit()
    else:
        print(f"\n{pName.value}")
        return pLedStyles


# int MLAPI_GetLedColor(BSTR type, DWORD index, DWORD* R, DWORD* G, DWORD* B)
def get_led_color(device_type, index):

    # Retrieves the current color of a specific LED on a device.

    # This function calls the MLAPI_GetLedColor function from the DLL to obtain
    # the RGB color values of a specific LED identified by the given device type
    # and index.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.

    # Returns:
    # dict: A dictionary containing the RGB color values with keys 'r', 'g', and 'b'.
    # If an error occurs, the function prints the error message and exits the program.

    dll.MLAPI_GetLedColor.argtypes = [
        BSTR,
        DWORD,
        POINTER(DWORD),
        POINTER(DWORD),
        POINTER(DWORD),
    ]
    dll.MLAPI_GetLedColor.restype = c_int
    r = DWORD()
    g = DWORD()
    b = DWORD()
    status = dll.MLAPI_GetLedColor(device_type, index, byref(r), byref(g), byref(b))
    if status != 0:
        error_message(status)
        exit()
    else:
        return {"r": r.value, "g": g.value, "b": b.value}


# int MLAPI_GetLedStyle(BSTR type, DWORD index, BSTR* style)
def get_led_style(device_type, index):

    # Retrieves the current style of a specific LED on a device.

    # This function calls the MLAPI_GetLedStyle function from the DLL to obtain
    # the style of a specific LED identified by the given device type and index.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.

    # Returns:
    # str: The current style of the LED. If an error occurs, the function prints
    # the error message and exits the program.

    dll.MLAPI_GetLedStyle.argtypes = [
        BSTR,
        DWORD,
        POINTER(BSTR),
    ]
    dll.MLAPI_GetLedStyle.restype = c_int
    style = BSTR()
    status = dll.MLAPI_GetLedStyle(device_type, index, byref(style))
    if status != 0:
        error_message(status)
        exit()
    else:
        return style.value


# int MLAPI_GetLedMaxBright(BSTR type, DWORD index, DWORD* maxLevel)
def get_led_max_bright(device_type, index):

    # Retrieves the maximum brightness level of a specific LED on a device.

    # This function calls the MLAPI_GetLedMaxBright function from the DLL to obtain
    # the maximum brightness level of a specific LED identified by the given device
    # type and index.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.

    # Returns:
    # int: The maximum brightness level of the LED. If an error occurs, the function
    # prints the error message and exits the program.

    dll.MLAPI_GetLedMaxBright.argtypes = [
        BSTR,
        DWORD,
        POINTER(DWORD),
    ]
    dll.MLAPI_GetLedMaxBright.restype = c_int
    max_level = DWORD()
    status = dll.MLAPI_GetLedMaxBright(device_type, index, byref(max_level))
    if status != 0:
        error_message(status)
        exit()
    else:
        return max_level.value


# int MLAPI_GetLedBright(BSTR type, DWORD index, DWORD* currentLevel)
def get_led_bright(device_type, index):

    # Retrieves the current brightness level of a specific LED on a device.

    # This function calls the MLAPI_GetLedBright function from the DLL to obtain
    # the current brightness level of a specific LED identified by the given device
    # type and index.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.

    # Returns:
    # int: The current brightness level of the LED. If an error occurs, the function
    # prints the error message and exits the program.

    dll.MLAPI_GetLedBright.argtypes = [
        BSTR,
        DWORD,
        POINTER(DWORD),
    ]
    dll.MLAPI_GetLedBright.restype = c_int
    current_level = DWORD()
    status = dll.MLAPI_GetLedBright(device_type, index, byref(current_level))
    if status != 0:
        error_message(status)
        exit()
    return current_level.value


# int MLAPI_GetLedMaxSpeed(BSTR type, DWORD index, DWORD* maxLevel)
def get_led_max_speed(device_type, index):

    # Retrieves the maximum speed level of a specific LED on a device.

    # This function calls the MLAPI_GetLedMaxSpeed function from the DLL to obtain
    # the maximum effect speed level (i.e. frequency of glow, flickering, etc)
    # of a specific LED identified by the given device type and index.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.

    # Returns:
    # int: The maximum speed level of the LED. If an error occurs, the function
    # prints the error message and exits the program

    dll.MLAPI_GetLedMaxSpeed.argtypes = [
        BSTR,
        DWORD,
        POINTER(DWORD),
    ]
    dll.MLAPI_GetLedMaxSpeed.restype = c_int
    max_level = DWORD()
    status = dll.MLAPI_GetLedMaxSpeed(device_type, index, byref(max_level))
    if status != 0:
        error_message(status)
        exit()
    return max_level.value


# int MLAPI_GetLedSpeed(BSTR type, DWORD index, DWORD* currentLevel)
def get_led_speed(device_type, index):

    # Retrieves the current speed level of a specific LED on a device.

    # This function calls the MLAPI_GetLedSpeed function from the DLL to obtain
    # the current effect speed level (i.e. frequency of glow, flickering, etc)
    # of a specific LED identified by the given device type and index.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.

    # Returns:
    # int: The current speed level of the LED. If an error occurs, the function
    # prints the error message and exits the program

    dll.MLAPI_GetLedSpeed.argtypes = [
        BSTR,
        DWORD,
        POINTER(DWORD),
    ]
    dll.MLAPI_GetLedSpeed.restype = c_int
    current_level = DWORD()
    status = dll.MLAPI_GetLedSpeed(device_type, index, byref(current_level))
    if status != 0:
        error_message(status)
        exit()
    return current_level.value


# int MLAPI_SetLedStyle(BSTR type, DWORD index, BSTR style)
def set_led_style(device_type, index, pLedStyles):

    # Sets the style of the specified LED on the a device.

    # This function calls the MLAPI_SetLedStyle function from the DLL to set
    # the style of the specified LED identified by the given device type and index.
    # The function accepts a pointer to an array of BSTRs representing the LED styles.
    # The user is prompted to choose a style from the list and the LED style is
    # updated accordingly.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.
    # pLedStyles (SAFEARRAY): A safe array containing BSTRs representing the LED styles.

    # Returns:
    # BSTR: The selected LED style. If an error occurs, the function prints the error
    # message and returns None.

    led_styles = cast(pLedStyles.contents.pvData, POINTER(BSTR))
    for i in range(pLedStyles.contents.rgsabound[0].cElements):
        print(f"{i}: {led_styles[i]}")
    n = int(input("Choose style: "))
    dll.MLAPI_SetLedStyle.argtypes = [
        BSTR,
        DWORD,
        BSTR,
    ]
    dll.MLAPI_SetLedStyle.restype = c_int
    status = dll.MLAPI_SetLedStyle(device_type, index, led_styles[n])
    if status != 0:
        error_message(status)
        print("Style setting failed.")
        return None
    else:
        print("LED Style updated.\n")
        return led_styles[n]


# int MLAPI_SetLedBright(BSTR type, DWORD index, DWORD level)
def set_led_bright(device_type, index, level):

    # Sets the brightness level of the specified LED on the a device.

    # This function calls the MLAPI_SetLedBright function from the DLL to set
    # the brightness level of the specified LED identified by the given device type
    # and index.
    # The function accepts the brightness level as a parameter. The LED brightness
    # is updated accordingly.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.
    # level (DWORD): The brightness level to set for the LED.

    # Returns:
    # bool: True if the brightness level was successfully set, False otherwise. If an
    # error occurs, the function prints the error message and returns False

    dll.MLAPI_SetLedBright.argtypes = [
        BSTR,
        DWORD,
        DWORD,
    ]
    dll.MLAPI_SetLedBright.restype = c_int
    status = dll.MLAPI_SetLedBright(device_type, index, level)
    if status != 0:
        error_message(status)
        print("Bright setting failed.")
        return None
    else:
        print("Bright updated.\n")
        return True


# int MLAPI_SetLedSpeed(BSTR type, DWORD index, DWORD level)
def set_led_speed(device_type, index, level):

    # Sets the effect speed level of the specified LED on the a device.

    # This function calls the MLAPI_SetLedSpeed function from the DLL to set
    # the effect speed level (i.e. frequency of glow, flickering, etc)
    # of the specified LED identified by the given device type and index.
    # The function accepts the effect speed level as a parameter. The LED effect
    # speed is updated accordingly.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.
    # level (DWORD): The effect speed level to set for the LED.

    # Returns:
    # bool: True if the effect speed level was successfully set, False otherwise. If an
    # error occurs, the function prints the error message and returns False

    dll.MLAPI_SetLedSpeed.argtypes = [
        BSTR,
        DWORD,
        DWORD,
    ]
    dll.MLAPI_SetLedSpeed.restype = c_int
    status = dll.MLAPI_SetLedSpeed(device_type, index, level)
    if status != 0:
        error_message(status)
        print("Speed setting failed.")
        return None
    else:
        print("Speed updated.\n")
        return True


# int MLAPI_SetLedColor(BSTR type, DWORD index, DWORD R, DWORD G, DWORD B)
def set_led_color(device_type, index, r, g, b):

    # Sets the color of the specified LED on the a device.

    # This function calls the MLAPI_SetLedColor function from the DLL to set the color
    # of the specified LED identified by the given device type and index.
    # The function accepts the RGB color values as parameters. The LED color is updated
    # accordingly.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.
    # R (DWORD): The red component of the RGB color.
    # G (DWORD): The green component of the RGB color.
    # B (DWORD): The blue component of the RGB color.

    # Returns:
    # bool: True if the color was successfully set, False otherwise. If an error occurs,
    # the function prints the error message and returns False

    dll.MLAPI_SetLedColor.argtypes = [
        BSTR,
        DWORD,
        DWORD,
        DWORD,
        DWORD,
    ]
    dll.MLAPI_SetLedColor.restype = c_int

    status = dll.MLAPI_SetLedColor(device_type, index, r, g, b)
    if status != 0:
        error_message(status)
        print("Color setting failed.")
        return None
    else:
        print("Color updated.\n")
        return True


# int MLAPI_SetLedColorsSync(BSTR type, DWORD R, DWORD G, DWORD B)
def set_led_colors_sync(device_type, r, g, b):

    # Sets the color of all LEDs on the specified device synchronously.
    # (In theory, had no chance to check it on my computer.)

    # This function calls the MLAPI_SetLedColorsSync function from the DLL to set
    # the color of all LEDs on the specified device synchronously.
    # The function accepts the RGB color values as parameters. The LED colors are updated
    # accordingly.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LEDs.
    # R (DWORD): The red component of the RGB color.
    # G (DWORD): The green component of the RGB color.
    # B (DWORD): The blue component of the RGB color.

    # Returns:
    # bool: True if the colors were successfully set, False otherwise. If an error occurs,
    # the function prints the error message and returns False

    dll.MLAPI_SetLedColorsSync.argtypes = [
        BSTR,
        DWORD,
        DWORD,
        DWORD,
    ]
    dll.MLAPI_SetLedColorsSync.restype = c_int
    status = dll.MLAPI_SetLedColorsSync(device_type, r, g, b)
    if status != 0:
        error_message(status)
        print("Color setting failed.")
        return None
    else:
        print("Colors updated synchronously.\n")
        return True


# int MLAPI_SetLedColors(BSTR type, DWORD index, SAFEARRAY** pLedName, DWORD* R, DWORD* G, DWORD* B)
def set_led_colors(device_type, index, pLedNames, pR, pG, pB):

    # Sets the color of the specified LED on the a device.
    # (Really don`t know how it works. :)

    # This function calls the MLAPI_SetLedColors function from the DLL to set
    # the color of the specified LED identified by the given device type and index.
    # The function accepts an array of LED names and RGB color values as parameters.
    # The LED color is updated accordingly.

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.
    # pLedNames (SAFEARRAY**): A pointer to an array of LED names.
    # pR (DWORD*): A pointer to an array of red color values.
    # pG (DWORD*): A pointer to an array of green color values.
    # pB (DWORD*): A pointer to an array of blue color values.

    # Returns:
    # bool: True if the colors were successfully set, False otherwise. If an error occurs,
    # the function prints the error message and returns False

    dll.MLAPI_SetLedColors.argtypes = [
        BSTR,
        DWORD,
        POINTER(POINTER(BSTR)),
        POINTER(DWORD),
        POINTER(DWORD),
        POINTER(DWORD),
    ]
    dll.MLAPI_SetLedColors.restype = c_int
    led_names = cast(pLedNames.contents.pvData, POINTER(BSTR))
    for i in range(pLedNames.contents.rgsabound[0].cElements):
        print(f"{i}: {led_names[i]}")
    n = int(input("Choose LED: "))
    status = dll.MLAPI_SetLedColors(device_type, index, led_names[n], pR, pG, pB)
    if status != 0:
        error_message(status)
        print("Color setting failed.")
        return None
    else:
        print("Color setting completed.")
        return True


# int MLAPI_SetLedColorEx(BSTR type, DWORD index, BSTR pLedName, DWORD R, DWORD G, DWORD B, DWORD Sync)
def set_led_color_ex(device_type, index, pLedName, r, g, b, sync):

    # Sets the color of the specified LED on the a device.
    # (Had no chance to check it on my computer)

    # This function calls the MLAPI_SetLedColorEx function from the DLL to set
    # the color of the specified LED identified by the given device type and index.
    # The function accepts the LED name, RGB color values, and a flag indicating whether
    # to update the LED color synchronously. The LED color is updated accordingly

    # Parameters:
    # device_type (BSTR): The type of the device containing the LED.
    # index (DWORD): The index of the LED within the specified device type.
    # pLedName (BSTR): The name of the LED.
    # R (DWORD): The red component of the RGB color.
    # G (DWORD): The green component of the RGB color.
    # B (DWORD): The blue component of the RGB color.
    # sync (DWORD): A flag indicating whether to update the LED color synchronously
    # Returns:
    # bool: True if the color was successfully set, False otherwise. If an error occurs,
    # the function prints the error message and returns False

    dll.MLAPI_SetLedColorEx.argtypes = [
        BSTR,
        DWORD,
        BSTR,
        DWORD,
        DWORD,
        DWORD,
        DWORD,
    ]
    dll.MLAPI_SetLedColorEx.restype = c_int
    status = dll.MLAPI_SetLedColorEx(device_type, index, pLedName, r, g, b, sync)
    if status != 0:
        error_message(status)
        print("Color setting failed.")
        return None
    else:
        print("Color updated.\n")
        return True


if __name__ == "__main__":
    initialize_dll()
    device_type, leds_count = get_device_info()
    if device_type:
        for i in range(leds_count):
            get_device_name_ex(device_type, i)
        index = int(input("Choose LED: "))
        led_styles = get_led_info(device_type, index)
        style = get_led_style(device_type, index)
        color = get_led_color(device_type, index)
        max_bright = get_led_max_bright(device_type, index)
        led_bright = get_led_bright(device_type, index)
        max_speed = get_led_max_speed(device_type, index)
        led_speed = get_led_speed(device_type, index)
        while True:
            print(f"\nStyle: {style}")
            print(f"Color: R-{color["r"]} G-{color["g"]} B-{color["b"]}")
            print(f"Max Bright Level: {max_bright}")
            print(f"Current Bright Level: {led_bright}")
            print(f"Max Speed Level: {max_speed}")
            print(f"Current Speed Level: {led_speed}")
            print(
                "\n1: Change LED color.\n"
                + "2: Change LED style.\n"
                + "3: Change LED Brightness.\n"
                + "4: Change LED Speed.\n"
                + "0: Exit"
            )
            a = int(input())
            if a == 1:
                try:
                    r = int(input("Enter red (0-255): "))
                    g = int(input("Enter green (0-255): "))
                    b = int(input("Enter blue (0-255): "))
                except ValueError:
                    print("Invalid input. Please enter integer values.")
                    continue
                if set_led_color(device_type, index, r, g, b):
                    color = {"r": r, "g": g, "b": b}
            elif a == 2:
                new_style = set_led_style(device_type, index, led_styles)
                if new_style:
                    style = new_style
            elif a == 3:
                try:
                    bright = int(input(f"Enter brightness level (1-{max_bright}): "))
                except ValueError:
                    print("Invalid input. Please enter integer values.")
                    continue
                if set_led_bright(device_type, index, bright):
                    led_bright = bright
            elif a == 4:
                try:
                    speed = int(input(f"Enter speed level (1-{max_speed}): "))
                except ValueError:
                    print("Invalid input. Please enter integer values.")
                    continue
                if set_led_speed(device_type, index, speed):
                    led_speed = speed
            elif a == 0:
                break
    release_dll()
