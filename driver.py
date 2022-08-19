import config as cfg

def get_active_window_name():
    import win32gui
    w = win32gui
    name = w.GetWindowText( w.GetForegroundWindow() )
    name = name.strip()
    return name


def resize_window(window_name, width, height):
    import pygetwindow
    win = pygetwindow.getWindowsWithTitle(window_name)[0]
    win.resizeTo(width, height)
    win.activate()


""" Generate a stochastic tree model and save it to a folder specified by the user """
def generate_stochastic_tree_model(output_folder_path, filename):
    import pygetwindow
    import time

    # Grab the x and y coordinates of the SpeedTree Modeler window
    win = pygetwindow.getWindowsWithTitle("SpeedTree Modeler v8.3.0 (Cinema Edition)")[0]
    win.moveTo(0,0)
    WINDOW_X = win.left #x:71
    WINDOW_Y = win.top  #y:-8

    # Click the "Random" button
    import pyautogui
    pyautogui.click(x= WINDOW_X + 214, y= WINDOW_Y + 98)
    time.sleep(0.5)

    # Export the model as a mesh by using Ctrl+E
    pyautogui.hotkey('ctrl', 'e', interval=0.1)

    # Wait for File Explorer to open
    while (get_active_window_name() != "Export Mesh"):
        time.sleep(0.5)

    time.sleep(0.3)

    # Check if output folder exists, if not create it
    import os
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)
    
    # Select the output folder by using Ctrl+L 
    pyautogui.hotkey('ctrl','l')   
    # Paste the output folder path into the File Explorer window
    pyautogui.typewrite(output_folder_path)
    pyautogui.press('enter')
    time.sleep(0.1)

    # Press Tab key followed by Esc to circumvent a bug in the File Explorer window
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('esc')
    time.sleep(0.3)

    # Select the tree model filename by using Alt+N
    pyautogui.hotkey('alt','n',interval=0.1)
    # Paste the tree model filename into the File Explorer window
    pyautogui.typewrite(filename)

    # Save the mesh by using Alt+S
    pyautogui.hotkey('alt','s',interval=0.1)

    # Click OK to export the mesh
    pyautogui.press('enter')

    # Wait for the SpeedTree Modeler to finish exporting the mesh
    while (get_active_window_name() != "SpeedTree Modeler v8.3.0 (Cinema Edition)"):
        import time
        time.sleep(0.5)
    
    time.sleep(1.0)


""" Open the seed file and run the simulation for the given number of itterations."""
def main(seed_spm_filepath, output_folder, number_of_itterations):
    # Open the seed file
    import os
    import time

    print("Starting macro simulation in 10 seconds...")
    time.sleep(10)

    os.startfile(r'{}'.format(seed_spm_filepath))
    
    # Wait for SpeedTree to load
    while (get_active_window_name() != "SpeedTree Modeler v8.3.0 (Cinema Edition)" and get_active_window_name() != "Alert"):
        time.sleep(0.5)

    time.sleep(0.3)

    # Progress through the Alerts if they exist
    while (get_active_window_name() == "Alert"):
        # Press the enter key to continue
        import pyautogui
        pyautogui.press('enter')
        time.sleep(0.5)

    # Wait until the SpeedTree Modeler is active
    while (get_active_window_name() != "SpeedTree Modeler v8.3.0 (Cinema Edition)"):
        time.sleep(0.5)

    # Wait for SpeedTree Software to load in completely
    time.sleep(5.0)

    # Resize the window to the desired size
    resize_window(window_name="SpeedTree Modeler v8.3.0 (Cinema Edition)", width=cfg.WINDOW_WIDTH, height=cfg.WINDOW_HEIGHT)
    
    # append the seed_spm_filepath filename to the output_folder
    output_folder_path = output_folder + "\\" + seed_spm_filepath.split("\\")[-1].split(".")[0]

    for i in range(number_of_itterations):
        generate_stochastic_tree_model(output_folder_path, "{}".format(i))

    # Run the simulation for the given number of itterations
    pass




# if this is run as a script, execute the main function
if __name__ == '__main__':
    # grab the arguments from the command line
    import sys
    seed_spm_filepath = sys.argv[1]
    output_folder = sys.argv[2]
    number_of_itterations = int(sys.argv[3])
    # execute the main function
    main(seed_spm_filepath, output_folder, number_of_itterations)
