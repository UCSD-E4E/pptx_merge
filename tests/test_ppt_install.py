'''Tests that PowerPoint is installed correctly
'''
import faulthandler

import win32com.client


def test_ppt_install():
    """Tests that PowerPoint can be created and quit
    """
    faulthandler.disable()
    ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
    ppt_instance.Activate()
    ppt_instance.Quit()
    faulthandler.enable()
